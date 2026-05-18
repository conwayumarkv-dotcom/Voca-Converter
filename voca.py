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
import time

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
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style.font.size = Pt(10.5)
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("통합 본문 단어장")
    title_run.font.bold = True
    title_run.font.size = Pt(18)
    title_p.paragraph_format.space_after = Pt(24)
    
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    col_widths = [Inches(1.5), Inches(2.0), Inches(4.0)]
    
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
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 2. 이미지 안전 분석 및 트래픽/제한 처리 로직
# ==========================================
def process_images_safely(client, uploaded_files, api_key, progress_bar, status_text):
    all_data = []
    total_files = len(uploaded_files)
    
    prompt = """
    이 이미지에서 영어 단어, 우리말 뜻, 영영 풀이를 추출해서 정확한 JSON 배열 형식으로 출력해줘.
    필기구로 수정한 흔적이나 추가로 적은 필기는 무시하고, 원래 인쇄되어 있던 텍스트만 추출해줘.
    결과는 오직 아래 구조를 가진 JSON 데이터만 반환해야 해:
    [
      {"word": "단어", "meaning": "품사 및 뜻", "definition": "영영 풀이 내용"}
    ]
    """
    
    current_displayed_percent = 0
    
    for idx, file in enumerate(uploaded_files):
        target_percent = int(((idx + 1) / total_files) * 100)
        
        # 처리 대기 상태 애니메이션 및 참고 파일의 정확한 모래시계 문구 표시
        pre_target = target_percent - 3 if target_percent > 3 else 0
        while current_displayed_percent < pre_target:
            current_displayed_percent += 1
            progress_bar.progress(current_displayed_percent / 100)
            status_text.markdown(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({total_files - idx}초)")
            time.sleep(0.01)
            
        page_data = None
        
        try:
            file.seek(0)
            image_bytes = file.read()
            
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type=file.type),
                    prompt
                ],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            page_data = json.loads(response.text)
            all_data.extend(page_data)
            
        except Exception as e:
            # 일일 한도 초과 안내 문구
            if idx > 0:
                st.warning("⚠️ 구글 계정의 하루 무료 사용량(20장)이 모두 마감되었습니다. 프로그램 보호를 위해 현재까지 변환된 파일들로만 워드를 생성합니다.")
                break
            else:
                st.error("❌ 오늘 사용 가능한 구글 무료 제공량(20장)을 모두 초과하여 변환을 시작할 수 없습니다. 내일 다시 시도해 주세요.")
                return None
            
        while current_displayed_percent < target_percent:
            current_displayed_percent += 1
            progress_bar.progress(current_displayed_percent / 100)
            status_text.markdown(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({total_files - idx}초)")
            time.sleep(0.01)
            
        time.sleep(0.2)
        
    if all_data:
        progress_bar.progress(1.0)
        status_text.success("🌿 정제 프로세스가 성공적으로 완료되었습니다. (100%)")
    return all_data

# ==========================================
# 3. Streamlit 메인 UI 대시보드
# ==========================================
st.set_page_config(page_title="Voca-converter", layout="centered", page_icon="📝")

# 🎨 디자인 커스텀 브랜딩 CSS (배경색, 폰트 색상, 버튼 및 크레딧 정렬)
st.markdown("""
    <style>
    /* 전체 배경 크림 베이지 톤 */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #FBF9F4 !important;
    }
    
    /* 레이아웃 폭 조정 */
    [data-testid="stMainBlockContainer"] {
        background-color: transparent !important;
        max-width: 720px !important;
        margin: 0 auto !important;
        padding-top: 50px !important;
    }
    
    /* 기본 테두리 제거 */
    [data-testid="stVerticalBlockBorderContainer"] {
        border: none !important;
        background: transparent !important;
        box-shadow: none !important;
        padding: 0 !important;
    }

    /* 상단 타이틀 로고 무드 */
    .brand-title {
        font-size: 52px !important;
        font-weight: 700 !important;
        color: #556B2F !important;
        text-align: center !important;
        margin-bottom: 5px !important;
        letter-spacing: -1px !important;
    }
    
    /* 서브 한글 설명 문구 */
    .brand-caption {
        font-size: 15px !important;
        color: #8C9A86 !important;
        text-align: center !important;
        margin-bottom: 5px !important;
        font-weight: 500 !important;
    }
    
    /* [수정] 아래줄 우측 정렬로 배치한 (Made by Manju) 스타일 */
    .brand-author {
        font-size: 13px !important;
        color: #A0ABA2 !important;
        text-align: right !important;
        margin-bottom: 45px !important;
        font-weight: 500 !important;
        padding-right: 5px;
    }

    /* 파일 업로더 박스 색상 */
    [data-testid="stFileUploader"] {
        border: none !important;
        background-color: #EEF1F6 !important;
        border-radius: 14px !important;
        padding: 20px 25px !important;
    }
    
    /* 카키 민트 변환 버튼 */
    div.stButton > button:first-child {
        background-color: #85A392 !important; 
        color: white !important;
        border: none !important;
        font-size: 16px !important;
        font-weight: 500 !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        box-shadow: none !important;
        width: auto !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #6C8B7A !important;
    }
    
    /* 다운로드 버튼 블루그레이 스타일 */
    [data-testid="stDownloadButton"]>button {
        background-color: #78909C !important;
        color: white !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 12px 24px !important;
    }
    [data-testid="stDownloadButton"]>button:hover {
        background-color: #607D8B !important;
    }
    
    /* 파스텔 블루 안내상자 커스텀 */
    div[data-testid="stNotification"] {
        background-color: #E8F1FC !important;
        border: none !important;
        border-radius: 12px !important;
    }
    div[data-testid="stNotification"] p {
        color: #1E60B4 !important;
        font-weight: 500 !important;
    }
    
    /* 게이지 진행률 바 색상 */
    .stProgress > div > div > div > div {
        background-color: #85A392 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 🏷️ 브랜드 헤더 섹션 (소문자 괄호 표기 및 위치 정밀 조정 완료)
st.markdown("<div class='brand-title'>Voca-converter</div>", unsafe_allow_html=True)
st.markdown("<div class='brand-caption'>사진 속 지문을 인식하여 편집 가능한 워드 문서(.docx)로 변환합니다.</div>", unsafe_allow_html=True)
st.markdown("<div class='brand-author'>(Made by Manju)</div>", unsafe_allow_html=True)

if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    st.error("❌ Streamlit Cloud 설정의 Secrets에 GEMINI_API_KEY가 등록되지 않았습니다.")
    st.stop()

uploaded_files = st.file_uploader(
    "변환할 영어 지문 사진을 업로드하세요 (복수 선택 가능)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write("")
    st.markdown(f"📂 **{len(uploaded_files)}개의 파일이 선택되었습니다.**")
    
    if st.button("Word 파일로 변환하기 ✨", type="primary"):
        client = genai.Client(api_key=api_key)
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        all_word_data = process_images_safely(client, uploaded_files, api_key, progress_bar, status_text)
        
        if all_word_data:
            st.toast("단어 데이터 정제가 완료되었습니다!")
            st.write("---")
            st.write("### 🔍 데이터 통합 미리보기")
            st.dataframe(all_word_data, use_container_width=True)
            
            word_file_buffer = create_word_document(all_word_data)
            
            st.write("")
            st.download_button(
                label="📥 정제된 수업용 Word 문서 다운로드 (.docx)",
                data=word_file_buffer,
                file_name="🔮_통합_영어단어장.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
