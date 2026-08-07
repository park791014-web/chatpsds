import streamlit as st
import os
import re
import io
import pandas as pd
import pdfplumber
from docx import Document
from pptx import Presentation
from PIL import Image
from openai import OpenAI

# 1. 페이지 기본 설정
st.set_page_config(page_title="Chat PSDongSung", layout="wide", page_icon=None)

# ==========================================
# 고급 타이포그래피 & 여백 조절 CSS
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&family=Noto+Sans+KR:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', 'Noto Sans KR', sans-serif;
    }

    /* 기본 메인 컨테이너 패딩 */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2rem !important;
    }
    
    /* 메인 화면 브랜드 헤더 */
    .brand-header {
        margin-bottom: 0.8rem;
        padding-top: 0.3rem;
        padding-bottom: 0.6rem;
        border-bottom: 1px solid #334155;
    }
    .brand-title {
        color: #f8fafc;
        font-size: 1.8rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
        line-height: 1.4;
        display: inline-block;
    }
    .brand-accent {
        color: #38bdf8;
        font-weight: 800;
    }
    .brand-sub {
        color: #94a3b8;
        font-size: 0.85rem;
        font-weight: 400;
        margin-top: 0.2rem;
        letter-spacing: -0.2px;
    }

    /* 사이드바 타이틀 커스텀 디자인 (가운데 정렬 + 여백 축소) */
    .sidebar-brand-title {
        color: #f8fafc;
        font-size: 1.4rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        text-align: center;
        margin-top: -0.5rem;
        margin-bottom: 0.8rem;
        padding-bottom: 0.6rem;
        border-bottom: 1px solid #334155;
    }

    /* 영화/드라마 보안 시스템 스타일 로그인 카드 */
    .login-card {
        text-align: center;
    }
    .login-title {
        color: #f8fafc;
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin-bottom: 0.3rem;
        line-height: 1.3;
    }
    .login-sub {
        color: #64748b;
        font-size: 0.85rem;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# Phase 5 유틸리티: NEIS 바이트 & 토큰 계산
# ==========================================
def calculate_neis_bytes(text):
    if not text:
        return 0, 0
    
    char_count = len(text)
    byte_count = 0
    
    for char in text:
        if char == '\n':
            byte_count += 2
        elif ord(char) > 127:
            byte_count += 3
        else:
            byte_count += 1
            
    return char_count, byte_count

FORBIDDEN_WORDS = [
    "대회", "수상", "상장", "올림피아드", "경시", "금상", "은상", "동상", "대상",
    "토익", "TOEIC", "토플", "TOEFL", "텝스", "TEPS", "HSK", "JLPT",
    "대학교", "영재교육원", "교육청", "교외", "학원", "자격증", "인증", "논문"
]

def check_forbidden_words(text):
    if not text:
        return []
    found = []
    for word in FORBIDDEN_WORDS:
        if word.lower() in text.lower():
            found.append(word)
    return list(set(found))

# ==========================================
# Secrets 파일 백그라운드 파싱
# ==========================================
def load_toml_secrets():
    secrets_data = {"ACCESS_CODE": "1234", "OPENROUTER_API_KEY": ""}
    
    possible_paths = [
        os.path.join(os.getcwd(), ".streamlit", "secrets.toml"),
        os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml"),
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                    
                    code_match = re.search(r'ACCESS_CODE\s*=\s*["\']([^"\']+)["\']', content)
                    if code_match:
                        secrets_data["ACCESS_CODE"] = code_match.group(1).strip()
                        
                    key_match = re.search(r'OPENROUTER_API_KEY\s*=\s*["\']([^"\']+)["\']', content)
                    if key_match:
                        secrets_data["OPENROUTER_API_KEY"] = key_match.group(1).strip()
                break
            except Exception:
                pass
                
    try:
        if not secrets_data["OPENROUTER_API_KEY"]:
            secrets_data["OPENROUTER_API_KEY"] = str(st.secrets.get("OPENROUTER_API_KEY", "")).strip()
        if secrets_data["ACCESS_CODE"] == "1234":
            secrets_data["ACCESS_CODE"] = str(st.secrets.get("ACCESS_CODE", "1234")).strip()
    except Exception:
        pass
        
    return secrets_data

SECRETS = load_toml_secrets()

def get_openrouter_client():
    api_key = SECRETS["OPENROUTER_API_KEY"]
    if not api_key:
        return None
    return OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )

# 모델 라인업
MODEL_MAP = {
    "Claude Opus 5": "anthropic/claude-opus-5",
    "GPT-5.6 Sol": "openai/gpt-5.6-sol",
    "Gemini 3.5 Flash": "google/gemini-flash-1.5",
    "Claude 3.5 Sonnet (검수 권장)": "anthropic/claude-3.5-sonnet:beta"
}

# 2. 세션 상태 초기화
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# 파일 업로더 초기화를 위한 카운터 Key
if "uploader_key_chat" not in st.session_state:
    st.session_state.uploader_key_chat = 0
if "uploader_key_std" not in st.session_state:
    st.session_state.uploader_key_std = 0

# 세션 관리 (일반 챗봇)
if "chat_sessions" not in st.session_state:
    st.session_state.chat_sessions = []
if "current_chat_idx" not in st.session_state:
    st.session_state.current_chat_idx = None

# 학생 세션 관리 (생기부)
if "student_records" not in st.session_state:
    st.session_state.student_records = []
if "current_student_idx" not in st.session_state:
    st.session_state.current_student_idx = None

# 누적 사용 토큰 추적
if "total_tokens_used" not in st.session_state:
    st.session_state.total_tokens_used = 0

# 검수 결과 캐시 세션
if "eval_result" not in st.session_state:
    st.session_state.eval_result = ""
if "eval_target_text" not in st.session_state:
    st.session_state.eval_target_text = ""

# ==========================================
# Phase 2: 파일 통합 파싱 및 캐싱 최적화
# ==========================================
@st.cache_data(ttl="1h", max_entries=50)
def parse_file_content(file_name, file_bytes):
    extracted_text = ""
    file_name = file_name.lower()
    try:
        file_io = io.BytesIO(file_bytes)
        
        if file_name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_io)
            extracted_text = df.to_string()
        elif file_name.endswith('.csv'):
            df = pd.read_csv(file_io)
            extracted_text = df.to_string()
        elif file_name.endswith('.pdf'):
            with pdfplumber.open(file_io) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        extracted_text += text + "\n"
        elif file_name.endswith('.docx'):
            doc = Document(file_io)
            for paragraph in doc.paragraphs:
                extracted_text += paragraph.text + "\n"
        elif file_name.endswith('.pptx'):
            try:
                prs = Presentation(file_io)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            extracted_text += shape.text + "\n"
            except Exception:
                extracted_text = f"[PPTX 문서 첨부됨: {file_name}]"
        elif file_name.endswith(('.hwpx', '.hwp')):
            try:
                extracted_text = file_bytes.decode('utf-8', errors='ignore')
            except Exception:
                extracted_text = f"[한글 문서 첨부됨: {file_name}]"
        elif file_name.endswith(('.png', '.jpg', '.jpeg')):
            extracted_text = f"[이미지 파일 첨부됨: {file_name}]"
        else:
            extracted_text = f"[첨부 파일: {file_name}]"
    except Exception as e:
        extracted_text = f"파싱 참고: {str(e)}"
    return extracted_text

def parse_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return ""
    # 파일 바이트 추출 후 캐싱된 함수 호출
    file_bytes = uploaded_file.getvalue()
    return parse_file_content(uploaded_file.name, file_bytes)

# 3. 영화/드라마 보안 시스템 스타일 콤팩트 중앙 로그인 화면
if not st.session_state.authenticated:
    st.markdown("""
        <style>
        /* 중앙 컬럼을 반투명 보안 대시보드 카드로 연출 */
        div[data-testid="column"]:nth-of-type(2) {
            background: #0f172a !important;
            border: 1px solid #334155 !important;
            border-radius: 16px !important;
            padding: 2.2rem 2rem 1.8rem 2rem !important;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.3) !important;
            margin-top: 2rem !important;
        }
        </style>
    """, unsafe_allow_html=True)
    
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        st.markdown("""
            <div class="login-card">
                <div class="login-title">Chat <span class="brand-accent">PSDongSung</span></div>
                <div class="login-sub">RESTRICTED ACCESS • TEACHER ONLY</div>
            </div>
        """, unsafe_allow_html=True)
        
        access_code = st.text_input("접속 코드 입력", type="password", placeholder="보안 코드를 입력하세요", label_visibility="collapsed")
        if st.button("SYSTEM ACCESS", type="primary", use_container_width=True):
            target_code = SECRETS["ACCESS_CODE"]
            if access_code == target_code:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("접속 코드가 올바르지 않습니다.")
    st.stop()

# 최초 접속 시 기본 세션 생성
if st.session_state.current_chat_idx is None and not st.session_state.chat_sessions:
    st.session_state.chat_sessions.append({"title": "새 대화", "messages": []})
    st.session_state.current_chat_idx = 0

# 4. 상단 헤더
col1, col2 = st.columns([3, 1], vertical_alignment="center")
with col1:
    st.markdown("""
        <div class="brand-header">
            <div class="brand-title">Chat <span class="brand-accent">PSDongSung</span></div>
            <div class="brand-sub">2022 개정 교육과정 기반 교사 전용 스마트 AI 에이전트</div>
        </div>
    """, unsafe_allow_html=True)
with col2:
    selected_model_name = st.selectbox(
        "모델 선택",
        list(MODEL_MAP.keys()),
        label_visibility="collapsed"
    )
    selected_model = MODEL_MAP[selected_model_name]

# 모드 선택 (Segmented Control로 더 깔끔하게 구성)
mode = st.segmented_control(
    "모드",
    ["일반 챗봇", "생기부 작성", "생기부 검수/진단"],
    default="일반 챗봇",
    label_visibility="collapsed"
)

# 5. Commissioning Sidebar rendering
with st.sidebar:
    st.markdown("""
        <div class="sidebar-brand-title">
            Chat <span class="brand-accent">PSDongSung</span>
        </div>
    """, unsafe_allow_html=True)
    
    if mode == "일반 챗봇":
        if st.button("새 대화 시작", icon=":material/add:", use_container_width=True):
            chat_key = f"uploader_chat_{st.session_state.uploader_key_chat}"
            if chat_key in st.session_state:
                del st.session_state[chat_key]
            st.session_state.uploader_key_chat += 1
            
            new_idx = len(st.session_state.chat_sessions)
            st.session_state.chat_sessions.append({"title": "새 대화", "messages": []})
            st.session_state.current_chat_idx = new_idx
            st.rerun()
            
        st.subheader("대화 목록")
        for idx, chat in enumerate(st.session_state.chat_sessions):
            btn_label = chat['title']
            is_active = (idx == st.session_state.current_chat_idx)
            if st.button(
                btn_label, 
                icon=":material/chat:", 
                key=f"chat_btn_{idx}", 
                use_container_width=True, 
                type="primary" if is_active else "secondary"
            ):
                chat_key = f"uploader_chat_{st.session_state.uploader_key_chat}"
                if chat_key in st.session_state:
                    del st.session_state[chat_key]
                st.session_state.uploader_key_chat += 1
                
                st.session_state.current_chat_idx = idx
                st.rerun()
                
    elif mode == "생기부 작성":
        if st.button("새 학생 작성", icon=":material/add:", use_container_width=True):
            std_key = f"uploader_std_{st.session_state.uploader_key_std}"
            if std_key in st.session_state:
                del st.session_state[std_key]
            st.session_state.uploader_key_std += 1
            
            st.session_state.current_student_idx = None
            st.rerun()
            
        st.subheader("학생 목록")
        for idx, student in enumerate(st.session_state.student_records):
            btn_label = f"학번: {student['id_val']}"
            is_active = (idx == st.session_state.current_student_idx)
            if st.button(
                btn_label, 
                icon=":material/description:", 
                key=f"std_btn_{idx}_{student['id_val']}", 
                use_container_width=True, 
                type="primary" if is_active else "secondary"
            ):
                std_key = f"uploader_std_{st.session_state.uploader_key_std}"
                if std_key in st.session_state:
                    del st.session_state[std_key]
                st.session_state.uploader_key_std += 1
                
                st.session_state.current_student_idx = idx
                st.rerun()
                
    elif mode == "생기부 검수/진단":
        st.info("작성된 생기부 문장을 업로드하거나 입력하시면 AI가 다각도로 정밀 진단합니다.")

    st.markdown("---")
    
    # 작업한 학생 일괄 엑셀 다운로드 버튼 배치 (생기부 작성 모드에서만 노출)
    if mode == "생기부 작성":
        if st.session_state.student_records:
            data_list = []
            for r in st.session_state.student_records:
                c_cnt, b_cnt = calculate_neis_bytes(r["draft"])
                data_list.append({
                    "학번": r["id_val"],
                    "작성영역": r["record_type"],
                    "관찰내용 및 키워드": r["memo"],
                    "생기부내용 (초안)": r["draft"],
                    "글자수": c_cnt,
                    "NEIS 바이트": b_cnt
                })
            df_bulk = pd.DataFrame(data_list)
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df_bulk.to_excel(writer, index=False, sheet_name='생기부_일괄내역')
            excel_data = excel_buffer.getvalue()
            
            st.download_button(
                label="일괄 엑셀 다운로드",
                data=excel_data,
                file_name="생기부_일괄초안.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button("일괄 엑셀 다운로드 (작성 데이터 없음)", disabled=True, use_container_width=True)
        st.markdown("---")

    st.caption(f"누적 사용 토큰: **{st.session_state.total_tokens_used:,} Tokens**")
    
    if st.button("로그아웃", icon=":material/logout:", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

st.divider()

# 6. Main runner
if mode == "일반 챗봇":
    curr_idx = st.session_state.current_chat_idx if st.session_state.current_chat_idx is not None else 0
    current_chat = st.session_state.chat_sessions[curr_idx]
    
    uploaded_files = st.file_uploader(
        "자료 첨부 (PDF, HWP, XLSX, DOCX, PPTX, 이미지 등)",
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"uploader_chat_{st.session_state.uploader_key_chat}"
    )
    
    if uploaded_files:
        with st.container(horizontal=True):
            for f in uploaded_files:
                st.badge(f.name, icon=":material/description:", color="gray")
                
    parsed_context = ""
    if uploaded_files:
        with st.expander("업로드된 파일 텍스트 미리보기", icon=":material/description:"):
            parsed_texts = []
            for file in uploaded_files:
                file_key = f"parsed_chat_{file.name}_{file.size}"
                if file_key not in st.session_state:
                    st.session_state[file_key] = parse_uploaded_file(file)
                file_text = st.session_state[file_key]
                parsed_texts.append(f"\n\n[파일: {file.name}]\n" + file_text)
                st.markdown(f"**{file.name}** ({len(file_text)}자 추출됨)")
                st.text(file_text[:300] + ("..." if len(file_text) > 300 else ""))
            parsed_context = "".join(parsed_texts)
                
    for msg in current_chat["messages"]:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])
            if "tokens" in msg:
                st.caption(f"소모 토큰: {msg['tokens']:,} Tokens")
            
    if prompt := st.chat_input("Chat PSDongSung에게 물어보기"):
        current_chat["messages"].append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.write(prompt)
            
        if current_chat["title"] == "새 대화":
            current_chat["title"] = prompt[:12] + "..." if len(prompt) > 12 else prompt
            
        with st.chat_message("assistant"):
            client = get_openrouter_client()
            if not client:
                res = "API 키 설정이 완료되지 않았습니다. (.streamlit/secrets.toml 확인 필요)"
                st.write(res)
                tokens_count = 0
            else:
                with st.spinner(f"[{selected_model_name}] 답변 생성 중..."):
                    try:
                        file_context = ""
                        if uploaded_files:
                            file_context = parsed_context
                        
                        api_messages = [{"role": "system", "content": "당신은 교사를 보조하는 유능한 AI 에이전트 'Chat PSDongSung'입니다."}]
                        for m in current_chat["messages"][:-1]:
                            api_messages.append({"role": m["role"], "content": m["content"]})
                        
                        final_prompt = prompt + file_context if file_context else prompt
                        api_messages.append({"role": "user", "content": final_prompt})
                        
                        response = client.chat.completions.create(
                            model=selected_model,
                            messages=api_messages
                        )
                        res = response.choices[0].message.content
                        st.write(res)
                        
                        if hasattr(response, 'usage') and response.usage:
                            tokens_count = response.usage.total_tokens
                            st.session_state.total_tokens_used += tokens_count
                            st.caption(f"소모 토큰: {tokens_count:,} Tokens")
                        else:
                            tokens_count = 0
                    except Exception as e:
                        res = "AI 응답 생성 실패."
                        st.error(f"오류 상세 내용: {str(e)}")
                        tokens_count = 0
                        
            new_msg = {"role": "assistant", "content": res}
            if tokens_count > 0:
                new_msg["tokens"] = tokens_count
            current_chat["messages"].append(new_msg)
            st.rerun()

elif mode == "생기부 작성":
    st.subheader("생기부 작성 전용 (Phase 6)")
    
    with st.expander("[공통] 교육과정 / 성취기준 및 개별 문체 가이드 등록", expanded=True):
        st.caption("과목별 성취기준 문서와 선생님 고유의 작성 스타일 예시 문장을 등록해 주세요.")
        global_ref_files = st.file_uploader(
            "공통 참조 파일 업로드 (PDF, HWP 등)",
            accept_multiple_files=True,
            key="global_ref_uploader",
            label_visibility="collapsed"
        )
        global_ref_text = ""
        if global_ref_files:
            ref_texts = []
            for g_file in global_ref_files:
                file_key = f"parsed_gref_{g_file.name}_{g_file.size}"
                if file_key not in st.session_state:
                    st.session_state[file_key] = parse_uploaded_file(g_file)
                ref_texts.append(f"\n\n[공통참조문서: {g_file.name}]\n" + st.session_state[file_key])
            global_ref_text = "".join(ref_texts)
            st.success(f"총 {len(global_ref_files)}개의 공통 참조 문서가 반영되었습니다. (프롬프트 캐싱 적용됨)")
                
        teacher_style_guide = st.text_area(
            "선생님 개별 작성 스타일(문체) 가이드 예시 문장",
            height=80,
            placeholder="평소 즐겨 쓰는 어조나 문장 구조 예시를 적어주세요. (예: 주도적인 탐구 태도와 논리적 사고 과정을 강조하며 서술함)"
        )

    st.divider()

    curr_std_idx = st.session_state.current_student_idx
    if curr_std_idx is not None and curr_std_idx < len(st.session_state.student_records):
        curr_std = st.session_state.student_records[curr_std_idx]
        default_id, default_memo, default_draft, default_type = curr_std["id_val"], curr_std["memo"], curr_std["draft"], curr_std.get("record_type", "교과세특 (1,500 Byte)")
    else:
        default_id, default_memo, default_draft, default_type = "", "", "", "교과세특 (1,500 Byte)"

    col_std1, col_std2 = st.columns(2)
    with col_std1:
        student_id = st.text_input("학번", value=default_id, placeholder="예: 10101")
    with col_std2:
        record_type = st.selectbox(
            "생기부 작성 영역 (NEIS 바이트 기준)",
            ["교과세특 (1,500 Byte)", "행동특성 및 종합의견 (1,500 Byte)", "자율활동 (1,500 Byte)", "동아리활동 (1,500 Byte)", "진로활동 (2,100 Byte)"],
            index=["교과세특 (1,500 Byte)", "행동특성 및 종합의견 (1,500 Byte)", "자율활동 (1,500 Byte)", "동아리활동 (1,500 Byte)", "진로활동 (2,100 Byte)"].index(default_type) if default_type in ["교과세특 (1,500 Byte)", "행동특성 및 종합의견 (1,500 Byte)", "자율활동 (1,500 Byte)", "동아리활동 (1,500 Byte)", "진로활동 (2,100 Byte)"] else 0
        )
        
    student_files = st.file_uploader(
        "개별 학생 관찰 기록 파일 업로드 (다중 선택 가능)",
        accept_multiple_files=True,
        key=f"uploader_std_{st.session_state.uploader_key_std}"
    )
    
    if student_files:
        with st.container(horizontal=True):
            for f in student_files:
                st.badge(f.name, icon=":material/description:", color="gray")
        
    file_text_combined = ""
    if student_files:
        with st.expander("업로드된 파일 텍스트 미리보기", icon=":material/description:"):
            student_file_texts = []
            for file in student_files:
                file_key = f"parsed_std_{file.name}_{file.size}"
                if file_key not in st.session_state:
                    st.session_state[file_key] = parse_uploaded_file(file)
                file_text = st.session_state[file_key]
                student_file_texts.append(f"\n[파일: {file.name}]\n" + file_text)
                st.markdown(f"**{file.name}** ({len(file_text)}자 추출됨)")
                st.text(file_text[:300] + ("..." if len(file_text) > 300 else ""))
            file_text_combined = "".join(student_file_texts)
            
    student_memo = st.text_area("학생 관찰 내용 및 키워드", value=default_memo, height=120, placeholder="수업 참여도, 수행평가 과정, 특기사항 등 입력")
    
    if st.button("초안 생성", type="primary"):
        if student_id:
            info_str = f"학번: {student_id}"
            
            client = get_openrouter_client()
            if not client:
                st.error("API 키 설정이 완료되지 않았습니다. (.streamlit/secrets.toml 확인 필요)")
            else:
                with st.spinner(f"[{selected_model_name}] {student_id} 학생 초안 생성 중..."):
                    try:
                        system_prompt = (
                            "당신은 대한민국 교육부 및 시도교육청의 학교생활기록부 작성 전문가입니다.\n"
                            "2022 개정 교육과정 기준 및 학교생활기록부 기재요령에 맞추어 다음 작성 지침을 '엄격히' 준수하세요:\n\n"
                            "1. [문장 시작 규칙 - 필수]: 문장 시작 시 '000 학생은', 'OOO 학생은', 'OO은' 같은 이름이나 학생 주어를 절대로 사용하지 마세요. 곧바로 학생의 학업적 성취 특성이나 활동 내용으로 문장을 바로 시작하세요.\n\n"
                            "2. [표현 제한 - 필수]: 영어 단어, 알파벳, 괄호 안 영문 병기(예: English) 및 중간점(·) 등의 특수문자를 절대로 사용하지 마세요. 모든 개념과 단어는 정돈된 표준 한글로만 작성하세요.\n\n"
                            f"3. [선생님 문체 가이드]: {teacher_style_guide if teacher_style_guide else '표준 교육용 어조 사용'}\n\n"
                            "4. [공통 참조 자료 반영]: 제시된 [공통 교육과정/성취기준 참조 자료]의 성취기준 및 교과역량을 적극 반영하여 작성하세요.\n\n"
                            "5. [영역별 작성 지침]:\n"
                            f"   - 현재 작성 영역: {record_type}\n"
                            "   - 교과세특인 경우: 성취수준 + 수행 과정 및 결과(과제, 분석내용, 도구) + 교과 핵심역량 + 교사 총평 구조 포함.\n"
                            "   - 창체/행발인 경우: 행동 특성, 공동체 의식, 주도적 활동 및 변화 모습을 구체적으로 진술.\n\n"
                            "6. [어조 및 분량]:\n"
                            "   - 문장 끝은 반드시 '~함', '~임' 어조로만 작성하세요.\n"
                            "   - 영역별 권장 바이트 한도 내에서 문장을 구체적이고 풍성하게 기술하세요.\n\n"
                            "7. [기재 금지어]: 대회, 수상, 외부 기관명, 공인어학성적, 사교육 관련 내용 절대 언급 금지."
                        )
                        
                        # 프롬프트 캐싱을 위한 메시지 구성
                        user_payload = []
                        if global_ref_text:
                            user_payload.append({
                                "type": "text",
                                "text": f"[공통 교육과정/성취기준 참조 자료]:\n{global_ref_text}",
                                "cache_control": {"type": "ephemeral"}
                            })
                            
                        user_payload.append({
                            "type": "text",
                            "text": f"[작성 정보]: {info_str} ({record_type})\n[선생님 문체 가이드]: {teacher_style_guide}\n[학생 관찰 메모]: {student_memo}\n[학생 첨부 자료]: {file_text_combined}"
                        })
                        
                        response = client.chat.completions.create(
                            model=selected_model,
                            messages=[
                                {"role": "system", "content": system_prompt},
                                {"role": "user", "content": user_payload}
                            ]
                        )
                        draft_res = response.choices[0].message.content
                        
                        tokens_count = 0
                        if hasattr(response, 'usage') and response.usage:
                            tokens_count = response.usage.total_tokens
                            st.session_state.total_tokens_used += tokens_count
                        
                        new_record = {
                            "info": info_str, "id_val": student_id,
                            "memo": student_memo, "draft": draft_res, "record_type": record_type
                        }
                        if tokens_count > 0:
                            new_record["tokens"] = tokens_count
                        
                        if curr_std_idx is not None and curr_std_idx < len(st.session_state.student_records):
                            st.session_state.student_records[curr_std_idx] = new_record
                        else:
                            st.session_state.student_records.append(new_record)
                            st.session_state.current_student_idx = len(st.session_state.student_records) - 1
                            
                        st.success(f"{student_id} 학생 초안 생성이 완료되었습니다!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"초안 생성 중 오류 발생: {str(e)}")
        else:
            st.warning("학번을 입력해 주세요.")
            
    if default_draft:
        st.divider()
        st.subheader("생성된 생기부 초안 및 영역별 NEIS 검수")
        
        edited_draft = st.text_area(
            "초안 편집",
            value=default_draft,
            height=220,
            label_visibility="collapsed",
            key=f"draft_ta_{curr_std_idx}"
        )
        if curr_std_idx is not None and curr_std_idx < len(st.session_state.student_records):
            st.session_state.student_records[curr_std_idx]["draft"] = edited_draft
            
        c_cnt, b_cnt = calculate_neis_bytes(edited_draft)
        max_bytes = 2100 if "진로활동" in default_type else 1500
        
        col_b1, col_b2, col_b3 = st.columns([1, 1, 2], vertical_alignment="center")
        with col_b1:
            st.metric("글자수 (공백 포함)", f"{c_cnt} 자")
        with col_b2:
            st.metric("NEIS 바이트", f"{b_cnt} / {max_bytes} Byte")
        with col_b3:
            progress_val = min(1.0, b_cnt / float(max_bytes))
            st.write(f"[{default_type.split(' ')[0]}] 용량 달성도 (최대 {max_bytes} Byte)")
            st.progress(progress_val)
            
        if b_cnt > max_bytes:
            st.warning(f"권장 입력 용량({max_bytes} Byte)을 {b_cnt - max_bytes} Byte 초과하였습니다.")
            
        forbidden_found = check_forbidden_words(edited_draft)
        if forbidden_found:
            st.error(f"기재 금지어 감지: {', '.join(forbidden_found)} (생기부 기재 지침 위반 위험이 있습니다.)")
        else:
            st.success("기재 금지어 검수 통과: 주요 기재 금지어가 감지되지 않았습니다.")
            
        if curr_std_idx is not None and curr_std_idx < len(st.session_state.student_records):
            curr_record = st.session_state.student_records[curr_std_idx]
            if "tokens" in curr_record:
                st.caption(f"이번 초안 생성에 소모된 토큰: **{curr_record['tokens']:,} Tokens**")

elif mode == "생기부 검수/진단":
    st.subheader("생기부 전문 진단 및 검수기")
    st.caption("기존에 작성된 생기부 문장을 업로드하거나 직접 입력하시면 지침 위반, 장단점, 문체 오류를 정밀 분석합니다.")
    
    col_e1, col_e2 = st.columns([3, 1])
    with col_e2:
        if st.session_state.eval_result:
            if st.button("새로 검수하기", width="stretch"):
                st.session_state.eval_result = ""
                st.session_state.eval_target_text = ""
                st.rerun()

    # type 제한 해제로 빨간색 에러 완전 방지
    eval_file = st.file_uploader("검수할 생기부 파일 첨부 (PDF, HWP, DOCX, TXT 등)", key="eval_file_uploader")
    eval_input_text = st.text_area("검수할 생기부 문장 직접 입력", height=180, placeholder="검수하고자 하는 생기부 특기사항 문단을 복사해서 붙여넣으세요.")
    
    target_eval_text = ""
    if eval_file:
        file_key = f"parsed_eval_{eval_file.name}_{eval_file.size}"
        if file_key not in st.session_state:
            with st.spinner("파일을 읽는 중입니다..."):
                st.session_state[file_key] = parse_uploaded_file(eval_file)
        target_eval_text = st.session_state[file_key]
    elif eval_input_text:
        target_eval_text = eval_input_text
    elif st.session_state.eval_target_text:
        target_eval_text = st.session_state.eval_target_text
        
    if target_eval_text:
        c_cnt, b_cnt = calculate_neis_bytes(target_eval_text)
        st.write(f"검수 대상 분량: {c_cnt}자 / **{b_cnt} Byte** (NEIS 기준)")
        
    if st.button("생기부 정밀 진단 시작", type="primary"):
        if target_eval_text:
            client = get_openrouter_client()
            if not client:
                st.error("API 키 설정이 완료되지 않았습니다. (.streamlit/secrets.toml 확인 필요)")
            else:
                with st.spinner(f"[{selected_model_name}] 생기부 정밀 분석 및 오류 검수 진행 중..."):
                    try:
                        eval_system_prompt = (
                            "당신은 대한민국 학교생활기록부 정밀 검수 평가관입니다.\n"
                            "제출된 생기부 텍스트를 분석하여 아래 구조에 맞춰 상세히 평가 리포트를 작성하세요:\n\n"
                            "1. 지침 위반 및 기재 금지어 적발: (대회, 수상, 외부기관, 어학성적, 사교육 유발 요소 여부 적발)\n"
                            "2. 문체 및 오탈자/비문 진단: (학생 이름/주어 시작 문장 오남용 여부, '~함/임' 어조 미준수 여부, 맞춤법 적발)\n"
                            "3. 작성의 장점 (강점): (구체적 수행과정, 도구 활용, 성취수준 표현 우수성 진술)\n"
                            "4. 작성의 단점 및 보완점: (추상적이거나 단순 총평에 그친 부분 지적)\n"
                            "5. 최종 개선/수정 제안 문장: (지침을 완벽히 준수한 최종 완성 문단 제시)\n"
                        )
                        
                        response = client.chat.completions.create(
                            model=selected_model,
                            messages=[
                                {"role": "system", "content": eval_system_prompt},
                                {"role": "user", "content": f"[검수 대상 생기부 텍스트]:\n{target_eval_text}"}
                            ]
                        )
                        eval_result = response.choices[0].message.content
                        
                        if hasattr(response, 'usage') and response.usage:
                            tokens = response.usage.total_tokens
                            st.session_state.total_tokens_used += tokens
                        
                        st.session_state.eval_result = eval_result
                        st.session_state.eval_target_text = target_eval_text
                        st.rerun()
                    except Exception as e:
                        st.error(f"진단 중 오류 발생: {str(e)}")
        else:
            st.warning("검수할 파일이나 텍스트를 입력해 주세요.")

    if st.session_state.eval_result:
        st.divider()
        st.markdown("### AI 생기부 정밀 진단 결과")
        st.markdown(st.session_state.eval_result)
