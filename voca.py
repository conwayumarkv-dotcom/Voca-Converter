import streamlit as st
from google import genai
from google.genai import types
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import io

# ==========================================
# 1. 워드 파일 디자인 & 생성 헬퍼 함수
# ==========================================
def set_cell_borders(cell, color="D9D9D9", sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_cell_shading(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def create_word_document(all_word_data):
    """메모리 버퍼에 스타일링된 워드 파일을 빌드하여 반환 (클라우드용)"""
    doc = docx.Document()
    
    # 여백 설정 (1인치)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    # 기본 폰트 설정
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style.font.size = Pt(10.5)
    
    # 제목 추가
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("통합 본문 단어장")
    title_run.font.bold = True
    title_run.font.size = Pt(18)
    title_p.paragraph_format.space_after = Pt(24)
    
    # 표 생성
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    col_widths = [Inches(1.5), Inches(2.0), Inches(4.0)]
    
    # 헤더 설정
    hdr_cells = table.rows[0].cells
    headers = ["본문 단어", "우리말 뜻", "영영 풀이"]
    for i, text in enumerate(headers):
        hdr_cells[i].text = text
        hdr_cells[i].width = col_widths[i]
        set_cell_shading(hdr_cells[i], "4F81BD")
        set_cell_borders(hdr_cells[i], color="A6A6A6")
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = docx.shared.RGBColor(255, 255, 255)
            
    # 통합된 데이터 행 추가
    for row_idx, item in enumerate(all_word_data):
        row_cells = table.add_row().cells
        row_data = [item.get("word", ""), item.get("meaning", ""), item.get("definition", "")]
        
        for i, text in enumerate(row_data):
            row_cells[i].text = text
            row_cells[i].width = col_widths[i]
            if row_idx % 2 == 1:
                set_cell_shading(row_cells[i], "F2F5F8")
            set_cell_borders(row_cells[i], color="D9D9D9")
            
            p = row_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            for run in p.runs:
                run.font.size = Pt(10)
                
    # 파일 객체를 바이트 버퍼로 반환
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 2. Streamlit 메인 UI 대시보드
# ==========================================
st.set_page_config(page_title="Voca 변환기 Cloud", layout="centered", page_icon="📝")

st.title("📝 멀티 이미지 단어장 워드 변환기")
st.write("24시간 언제 어디서나 단어장 사진들을 업로드하여 하나의 깔끔한 워드(.docx) 파일로 통합 다운로드하세요.")

# 사이드바 설정
with st.sidebar:
    st.header("⚙️ API 설정")
    api_key = st.text_input("Google Gemini API Key", type="password")
    st.markdown("[Gemini API Key 발급받기](https://aistudio.google.com/)")

# 여러 장 업로드
uploaded_files = st.file_uploader(
    "단어장 사진 파일들을 선택하세요 (여러 장 동시 선택 가능)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📂 **선택된 파일 수:** {len(uploaded_files)}개")
    
    if not api_key:
        st.warning("👈 프로그램을 실행하려면 왼쪽 사이드바에 Gemini API Key를 입력해 주세요.")
    else:
        if st.button("🚀 전체 사진 분석 및 하나의 Word 파일로 병합 생성"):
            all_word_data = []
            file_progress = st.progress(0)
            
            for idx, file in enumerate(uploaded_files):
                st.write(f"🔄 ({idx+1}/{len(uploaded_files)}) `{file.name}` 분석 중...")
                try:
                    client = genai.Client(api_key=api_key)
                    image_bytes = file.read()
                    
                    prompt = """
                    이 이미지에서 영어 단어, 우리말 뜻, 영영 풀이를 추출해서 정확한 JSON 배열 형식으로 출력해줘.
                    필기구로 수정한 흔적이나 추가로 적은 필기는 무시하고, 원래 인쇄되어 있던 텍스트만 추출해줘.
                    결과는 오직 아래 구조를 가진 JSON 데이터만 반환해야 해:
                    [
                      {"word": "단어", "meaning": "품사 및 뜻", "definition": "영영 풀이 내용"}
                    ]
                    """
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[
                            types.Part.from_bytes(data=image_bytes, mime_type=file.type),
                            prompt
                        ],
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    
                    page_data = json.loads(response.text)
                    all_word_data.extend(page_data)
                    
                except Exception as e:
                    st.error(f"`{file.name}` 처리 중 오류 발생: {e}")
                
                file_progress.progress((idx + 1) / len(uploaded_files))
            
            # 파일 빌드 및 브라우저 다운로드 제공
            if all_word_data:
                st.success("🎉 모든 사진의 단어 통합 완료!")
                
                # 화면에 전체 프리뷰 표 출력
                st.write("### 🔍 통합 추출 데이터 미리보기")
                st.dataframe(all_word_data, use_container_width=True)
                
                # 메모리에서 생성된 워드 파일 가져오기
                word_file_buffer = create_word_document(all_word_data)
                
                # 다운로드 버튼 생성 (이 버튼을 누르면 사용자의 로컬 컴퓨터 다운로드 폴더로 들어갑니다)
                st.download_button(
                    label="📥 깔끔한 Word 문서 다운로드 (.docx)",
                    data=word_file_buffer,
                    file_name="🔮_통합_영어단어장.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )