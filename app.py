import streamlit as st
import os
import re
import io
import pandas as pd
import pdfplumber
from docx import Document
from pptx import Presentation
from openai import OpenAI
from utils.ui import hide_streamlit_chrome, paste_listener

# 1. 페이지 기본 설정 (사이드바 기본 열림 상태로 고정)
st.set_page_config(page_title="Chat PSDongSung", layout="wide", initial_sidebar_state="expanded")

# Streamlit UI 크롬 정밀 제거
hide_streamlit_chrome()

# ==========================================
# 컴팩트 레이아웃 & 고밀도 Sticky Header & 통합 CSS
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&family=Noto+Sans+KR:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', 'Noto Sans KR', sans-serif;
    }

    /* 메인 앱 컨테이너 상단 여백 최소화 (뷰어 툴바 이하 밀착) */
    .block-container,
    [data-testid="stAppViewContainer"] .block-container {
        padding-top: 2.5rem !important;
        padding-bottom: 8.5rem !important;
    }
    
    /* Streamlit 수직 요소 간격 */
    div[data-testid="stVerticalBlock"] {
        gap: 0.5rem !important;
    }
    
    div[data-testid="stExpander"] {
        margin-bottom: 0.4rem !important;
    }
    
    /* 상단 Sticky 고정 헤더 컨테이너 컴팩트화 & 잘림 방지 (overflow: visible) */
    div[data-testid="stVerticalBlock"] > div:has(div[key="sticky_header"]),
    div[key="sticky_header"],
    [data-testid="stElementContainer"]:has(div[key="sticky_header"]),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div[key="sticky_header"]) {
        position: sticky !important;
        top: 2.875rem !important;
        z-index: 9999 !important;
        background-color: #0f172a !important;
        padding-top: 0.4rem !important;
        padding-bottom: 0.4rem !important;
        margin-bottom: 0.4rem !important;
        border-bottom: 1px solid #334155 !important;
        overflow: visible !important;
    }
    
    /* 모델 selectbox 사방 둥근 테두리 잘림 방지 및 여백 확보 */
    div[key="sticky_header"] div[data-testid="stSelectbox"],
    div[key="sticky_header"] div[data-baseweb="select"],
    div[key="sticky_header"] div[data-baseweb="select"] > div {
        overflow: visible !important;
        border-radius: 8px !important;
        margin-top: 2px !important;
    }
    
    /* 일반 챗봇 전용 대기 첨부자료 칩스 패널 (st.chat_input 바로 위) */
    div[data-testid="stVerticalBlock"] > div:has(div[key="chat_pending_chips"]),
    div[key="chat_pending_chips"],
    [data-testid="stElementContainer"]:has(div[key="chat_pending_chips"]) {
        position: sticky !important;
        bottom: 4.2rem !important;
        z-index: 998 !important;
        background-color: #0f172a !important;
        padding: 0.3rem 0.6rem !important;
        margin-bottom: 0.2rem !important;
        border-radius: 8px 8px 0 0 !important;
        border-top: 1px solid #334155 !important;
        overflow: visible !important;
    }
    
    /* 메인 화면 브랜드 헤더 여백 */
    .brand-header {
        margin-bottom: 0.2rem;
        padding-top: 0.1rem;
        padding-bottom: 0.2rem;
    }
    
    /* 테마 고정에 따라 항상 밝고 선명한 텍스트 가시성 유지 */
    .brand-title {
        color: #f8fafc !important;
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
        line-height: 1.4;
        display: inline-block;
    }
    .brand-accent {
        color: #38bdf8 !important;
        font-weight: 800;
    }
    .brand-sub {
        color: #94a3b8 !important;
        font-size: 0.85rem;
        font-weight: 400;
        margin-top: 0.2rem;
        letter-spacing: -0.2px;
    }
    .sidebar-brand-title {
        color: #f8fafc !important;
        font-size: 1.4rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        text-align: center;
        margin-top: -0.5rem;
        margin-bottom: 0.8rem;
        padding-bottom: 0.6rem;
        border-bottom: 1px solid #334155;
    }

    /* 영화/드라마 보안 시스템 스타일 로그인 카드 (항상 다크 테마 고정) */
    .login-card {
        text-align: center;
    }
    .login-title {
        color: #f8fafc !important;
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin-bottom: 0.3rem;
        line-height: 1.3;
    }
    .login-sub {
        color: #64748b !important;
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

import hashlib

# 이미지 데이터의 MD5 해시값 계산을 통한 중복 방지 헬퍼 함수
def calculate_image_hash(base64_data):
    if not base64_data:
        return ""
    payload = base64_data.split(",")[-1]
    return hashlib.md5(payload.encode("utf-8")).hexdigest()

# 모델 라인업 설정 (기능 지원 여부 및 역할 일원화 관리)
MODEL_MAP = {
    "GPT-5.6 Luna": {
        "id": "openai/gpt-5.6-luna",
        "supports_images": True,
        "role": "빠른 검토 · 가성비"
    },
    "Gemini 3.6 Flash": {
        "id": "google/gemini-3.6-flash",
        "supports_images": True,
        "role": "빠른 초안 작성"
    },
    "Claude Sonnet 5": {
        "id": "anthropic/claude-sonnet-5",
        "supports_images": True,
        "role": "생기부 작성 · 검수 권장"
    },
    "Claude Opus 5": {
        "id": "anthropic/claude-opus-5",
        "supports_images": True,
        "role": "정밀 분석"
    },
    "GPT-5.6 Sol": {
        "id": "openai/gpt-5.6-sol",
        "supports_images": True,
        "role": "고급 추론 · 복잡한 업무"
    }
}

# OpenRouter 모델 활성/사용가능 상태 검증 함수
def check_openrouter_model_availability(model_id):
    if "openrouter_models_cache" not in st.session_state:
        st.session_state.openrouter_models_cache = None
        
    if st.session_state.openrouter_models_cache is None:
        try:
            import urllib.request
            import json
            url = "https://openrouter.ai/api/v1/models"
            req = urllib.request.Request(
                url, 
                headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req, timeout=4) as response:
                res_data = json.loads(response.read().decode())
                available_ids = {m.get("id") for m in res_data.get("data", []) if m.get("id")}
                st.session_state.openrouter_models_cache = available_ids
        except Exception:
            # 네트워크 타임아웃 등의 이슈로 조회가 실패하면 중단 방지를 위해 통과시킴
            return True
            
    if st.session_state.openrouter_models_cache:
        return model_id in st.session_state.openrouter_models_cache
    return True

# 2. 세션 상태 초기화
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "user_has_manually_chosen_model" not in st.session_state:
    st.session_state.user_has_manually_chosen_model = False

if "selected_model_name" not in st.session_state:
    st.session_state.selected_model_name = "GPT-5.6 Luna"

# 파일 업로더 초기화를 위한 카운터 Key
if "uploader_key_chat" not in st.session_state:
    st.session_state.uploader_key_chat = 0
if "uploader_key_std" not in st.session_state:
    st.session_state.uploader_key_std = 0
if "uploader_key_eval" not in st.session_state:
    st.session_state.uploader_key_eval = 0

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

# 검수 세션 관리 (생기부 검수/진단)
if "eval_records" not in st.session_state:
    st.session_state.eval_records = []
if "current_eval_idx" not in st.session_state:
    st.session_state.current_eval_idx = None
if "eval_text_widget" not in st.session_state:
    st.session_state.eval_text_widget = ""

# 누적 사용 토큰 추적
if "total_tokens_used" not in st.session_state:
    st.session_state.total_tokens_used = 0

# 검수 결과 캐시 세션 (기존 단일 세션 호환 유지)
if "eval_result" not in st.session_state:
    st.session_state.eval_result = ""
if "eval_target_text" not in st.session_state:
    st.session_state.eval_target_text = ""

# 모드별 독립된 첨부파일 큐 초기화 (기존 list 세션 방어 및 마이그레이션 적용)
if "attachments" not in st.session_state or not isinstance(st.session_state.attachments, dict):
    st.session_state.attachments = {"chat": [], "student": [], "eval": []}
else:
    # 각 scope key가 누락된 경우 자동 보완
    st.session_state.attachments.setdefault("chat", [])
    st.session_state.attachments.setdefault("student", [])
    st.session_state.attachments.setdefault("eval", [])

def calculate_bytes_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def add_attachment(name, mime_type, data_base64, size_bytes, source="upload", file_type="image", scope="chat"):
    img_hash = calculate_image_hash(data_base64)
    exists = any(item.get("hash") == img_hash for item in st.session_state.attachments[scope])
    if not exists:
        st.session_state.attachments[scope].append({
            "name": name,
            "mime_type": mime_type,
            "data": data_base64,
            "size": size_bytes,
            "size_kb": size_bytes / 1024.0,
            "source": source,
            "type": file_type,
            "hash": img_hash
        })

def add_document_attachment(name, mime_type, text_content, size_bytes, file_hash, scope="chat"):
    exists = any(item.get("hash") == file_hash for item in st.session_state.attachments[scope])
    if not exists:
        st.session_state.attachments[scope].append({
            "name": name,
            "mime_type": mime_type,
            "data": text_content,
            "size": size_bytes,
            "size_kb": size_bytes / 1024.0,
            "source": "upload",
            "type": "document",
            "hash": file_hash
        })

def render_attachments_panel(uploader_key, scope="chat"):
    paste_key = f"paste_listener_{scope}"
    
    # 1. paste listener callback (scope 인식)
    def handle_pasted_image():
        state = st.session_state.get(paste_key)
        if state and hasattr(state, "pasted_image") and state.pasted_image:
            img_data = state.pasted_image
            add_attachment(
                name=img_data["name"],
                mime_type=img_data["type"],
                data_base64=img_data["data"],
                size_bytes=img_data["size"],
                source="paste",
                file_type="image",
                scope=scope
            )
            
    paste_listener(key=paste_key, on_pasted_image=handle_pasted_image)
    
    current_list = st.session_state.attachments[scope]
    
    # 2. Scope별 구분 라벨 (요청사항 7 반영)
    if scope == "student":
        st.caption("현재 학생 첨부자료 (파일 선택 또는 Ctrl+V 캡처)")
    elif scope == "eval":
        st.caption("검토 첨부자료 (파일 선택 또는 Ctrl+V 캡처)")

    # 3. 항상 노출형 콤팩트 파일 업로더 (요청사항 5 & 6 반영: expander/popover 래핑 없음)
    uploaded_files = st.file_uploader(
        "파일 선택 또는 Ctrl+V로 이미지 붙여넣기",
        accept_multiple_files=True,
        key=uploader_key,
        label_visibility="collapsed"
    )
    
    if uploaded_files:
        for f in uploaded_files:
            file_bytes = f.getvalue()
            file_hash = calculate_bytes_hash(file_bytes)
            
            if f.name.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                import base64
                b64_data = f"data:{f.type};base64," + base64.b64encode(file_bytes).decode("utf-8")
                add_attachment(
                    name=f.name,
                    mime_type=f.type,
                    data_base64=b64_data,
                    size_bytes=f.size,
                    source="upload",
                    file_type="image",
                    scope=scope
                )
            else:
                text_content = parse_uploaded_file(f)
                add_document_attachment(
                    name=f.name,
                    mime_type=f.type,
                    text_content=text_content,
                    size_bytes=f.size,
                    file_hash=file_hash,
                    scope=scope
                )

    # 4. 첨부자료가 존재할 때만 콤팩트 칩스 바 렌더링
    if current_list:
        max_cols = min(len(current_list), 6)
        cols = st.columns(max_cols)
        for idx, item in enumerate(current_list):
            with cols[idx % max_cols]:
                if item["type"] == "image":
                    st.image(item["data"], width=45)
                    btn_label = f"✕ {item['name'][:8]}"
                else:
                    btn_label = f"✕ [문서]{item['name'][:8]}"
                if st.button(btn_label, key=f"del_att_{uploader_key}_{idx}", help=f"{item['name']} ({item['size_kb']:.0f}KB) 삭제"):
                    current_list.pop(idx)
                    st.rerun()

def compile_api_payload(prompt, selected_model_name, scope="chat"):
    images = []
    text_contexts = []
    
    for item in st.session_state.attachments[scope]:
        if item["type"] == "image":
            images.append(item)
        elif item["type"] == "document":
            text_contexts.append(f"\n\n[첨부문서: {item['name']}]\n" + item["data"])
            
    model_supports_img = MODEL_MAP[selected_model_name].get("supports_images", False)
    
    if images and not model_supports_img:
        st.error(f"현재 선택한 모델 ({selected_model_name})은 이미지 분석을 지원하지 않습니다. 이미지를 제외하거나 다른 모델을 선택해 주세요.")
        st.stop()
        
    doc_text_combined = "".join(text_contexts)
    final_text = prompt + doc_text_combined if doc_text_combined else prompt
    
    content_payload = []
    content_payload.append({"type": "text", "text": final_text})
    
    for img in images:
        clean_b64 = img["data"].split(",")[-1]
        content_payload.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:{img['mime_type']};base64,{clean_b64}"
            }
        })
        
    return content_payload

# ==========================================
# 파일 통합 파싱 및 캐싱 최적화 (버퍼링 완벽 해결)
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
    
    # 텍스트 입력창 타이핑 중 랙 걸리는 현상을 방지하는 st.session_state 2차 캐싱
    file_key = f"parsed_{uploaded_file.name}_{uploaded_file.size}"
    if file_key not in st.session_state:
        file_bytes = uploaded_file.getvalue()
        st.session_state[file_key] = parse_file_content(uploaded_file.name, file_bytes)
        
    return st.session_state[file_key]

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

# 4. 상단 고정 헤더 (제목 / 모델 선택 / 모드 탭)
with st.container(key="sticky_header"):
    mode_from_state = st.session_state.get("mode_control_widget", "일반 챗봇")
    model_keys = list(MODEL_MAP.keys())

    # 사용자가 수동으로 변경하지 않은 경우, 생기부 작성/검수 진단 모드는 Claude Sonnet 5를 기본 추천
    if not st.session_state.get("user_has_manually_chosen_model", False):
        if mode_from_state in ["생기부 작성", "생기부 검수/진단"]:
            st.session_state.selected_model_name = "Claude Sonnet 5"
        else:
            st.session_state.selected_model_name = "GPT-5.6 Luna"

    # 세션에 기록된 모델명의 인덱스 검색
    default_idx = 0
    if st.session_state.selected_model_name in model_keys:
        default_idx = model_keys.index(st.session_state.selected_model_name)

    col1, col2 = st.columns([3, 1], vertical_alignment="center")
    with col1:
        st.markdown("""
            <div class="brand-header">
                <div class="brand-title">Chat <span class="brand-accent">PSDongSung</span></div>
                <div class="brand-sub">교사 전용 스마트 AI 에이전트</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        def on_model_change():
            st.session_state.user_has_manually_chosen_model = True

        selected_model_name = st.selectbox(
            "모델 선택",
            model_keys,
            index=default_idx,
            label_visibility="collapsed",
            key="model_selectbox_widget",
            on_change=on_model_change
        )
        st.session_state.selected_model_name = selected_model_name
        selected_model_info = MODEL_MAP[selected_model_name]
        selected_model = selected_model_info["id"]
        role_text = selected_model_info["role"]
        
        st.caption(role_text)

    # 모드 선택
    mode = st.segmented_control(
        "모드",
        ["일반 챗봇", "생기부 작성", "생기부 검수/진단"],
        default="일반 챗봇",
        label_visibility="collapsed",
        key="mode_control_widget"
    )

# 5. 사이드바 구성
with st.sidebar:
    st.markdown("""
        <div class="sidebar-brand-title">
            Chat <span class="brand-accent">PSDongSung</span>
        </div>
    """, unsafe_allow_html=True)
    
    if mode == "일반 챗봇":
        if st.button("새 대화 시작", use_container_width=True):
            chat_key = f"uploader_chat_{st.session_state.uploader_key_chat}"
            if chat_key in st.session_state:
                del st.session_state[chat_key]
            st.session_state.uploader_key_chat += 1
            st.session_state.attachments["chat"] = []
            
            new_idx = len(st.session_state.chat_sessions)
            st.session_state.chat_sessions.append({"title": "새 대화", "messages": []})
            st.session_state.current_chat_idx = new_idx
            st.rerun()
            
        st.subheader("대화 목록")
        for idx, chat in enumerate(st.session_state.chat_sessions):
            btn_label = f"{chat['title']}"
            is_active = (idx == st.session_state.current_chat_idx)
            if st.button(
                btn_label, 
                key=f"chat_btn_{idx}", 
                use_container_width=True, 
                type="primary" if is_active else "secondary"
            ):
                chat_key = f"uploader_chat_{st.session_state.uploader_key_chat}"
                if chat_key in st.session_state:
                    del st.session_state[chat_key]
                st.session_state.uploader_key_chat += 1
                st.session_state.attachments["chat"] = []
                
                st.session_state.current_chat_idx = idx
                st.rerun()
                
    elif mode == "생기부 작성":
        if st.button("새 학생 작성", use_container_width=True):
            std_key = f"uploader_std_{st.session_state.uploader_key_std}"
            if std_key in st.session_state:
                del st.session_state[std_key]
            st.session_state.uploader_key_std += 1
            st.session_state.attachments["student"] = []
            
            st.session_state.current_student_idx = None
            st.rerun()
            
        st.subheader("학생 목록")
        for idx, student in enumerate(st.session_state.student_records):
            btn_label = f"학번: {student['id_val']}"
            is_active = (idx == st.session_state.current_student_idx)
            if st.button(
                btn_label, 
                key=f"std_btn_{idx}_{student['id_val']}", 
                use_container_width=True, 
                type="primary" if is_active else "secondary"
            ):
                std_key = f"uploader_std_{st.session_state.uploader_key_std}"
                if std_key in st.session_state:
                    del st.session_state[std_key]
                st.session_state.uploader_key_std += 1
                st.session_state.attachments["student"] = []
                
                st.session_state.current_student_idx = idx
                st.rerun()
                
    elif mode == "생기부 검수/진단":
        if st.button("새 검토", use_container_width=True):
            eval_key = f"uploader_eval_{st.session_state.uploader_key_eval}"
            if eval_key in st.session_state:
                del st.session_state[eval_key]
            st.session_state.uploader_key_eval += 1
            st.session_state.attachments["eval"] = []
            
            st.session_state.current_eval_idx = None
            st.session_state.eval_result = ""
            st.session_state.eval_target_text = ""
            st.session_state.eval_text_widget = ""
            st.rerun()
            
        st.subheader("검토 목록")
        for idx, rec in enumerate(st.session_state.eval_records):
            btn_label = f"{rec.get('title', f'검토 {idx+1}')}"
            is_active = (idx == st.session_state.current_eval_idx)
            if st.button(
                btn_label, 
                key=f"eval_btn_{idx}", 
                use_container_width=True, 
                type="primary" if is_active else "secondary"
            ):
                eval_key = f"uploader_eval_{st.session_state.uploader_key_eval}"
                if eval_key in st.session_state:
                    del st.session_state[eval_key]
                st.session_state.uploader_key_eval += 1
                st.session_state.attachments["eval"] = []
                
                st.session_state.current_eval_idx = idx
                st.session_state.eval_target_text = rec.get("target_text", "")
                st.session_state.eval_text_widget = rec.get("target_text", "")
                st.session_state.eval_result = rec.get("result", "")
                if rec.get("model") and rec.get("model") in MODEL_MAP:
                    st.session_state.selected_model_name = rec.get("model")
                st.rerun()

    st.markdown("---")
    
    # 일괄 엑셀 다운로드 (생기부 작성 모드에서만 노출)
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
                label="전체 학생 기록 다운로드 (Excel)",
                data=excel_data,
                file_name="전체생기부기록.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button("일괄 엑셀 다운로드 (작성 데이터 없음)", disabled=True, use_container_width=True)
        st.markdown("---")

    st.caption(f"누적 사용 토큰: **{st.session_state.total_tokens_used:,} Tokens**")
    
    if st.button("로그아웃", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

st.divider()

# 6. 메인 동작 영역
if mode == "일반 챗봇":
    curr_idx = st.session_state.current_chat_idx if st.session_state.current_chat_idx is not None else 0
    current_chat = st.session_state.chat_sessions[curr_idx]
    
    # 1. 대화 히스토리 화면 먼저 표시
    for msg in current_chat["messages"]:
        with st.chat_message(msg["role"]):
            if isinstance(msg["content"], list):
                for part in msg["content"]:
                    if part["type"] == "text":
                        st.write(part["text"])
                    elif part["type"] == "image_url":
                        st.image(part["image_url"]["url"], width=250)
            else:
                st.write(msg["content"])
            if "tokens" in msg:
                st.caption(f"소모 토큰: {msg['tokens']:,} Tokens")

    # 2. Ctrl+V 캡처 이미지 또는 대기 첨부자료 칩스 표시 (st.chat_input 바로 위)
    paste_key = "paste_listener_chat"
    def handle_pasted_image():
        state = st.session_state.get(paste_key)
        if state and hasattr(state, "pasted_image") and state.pasted_image:
            img_data = state.pasted_image
            add_attachment(
                name=img_data["name"],
                mime_type=img_data["type"],
                data_base64=img_data["data"],
                size_bytes=img_data["size"],
                source="paste",
                file_type="image",
                scope="chat"
            )
    paste_listener(key=paste_key, on_pasted_image=handle_pasted_image)

    current_attachments = st.session_state.attachments["chat"]
    if current_attachments:
        with st.container(key="chat_pending_chips"):
            st.caption("대기 중인 첨부자료 (Ctrl+V / 파일 첨부):")
            max_cols = min(len(current_attachments), 6)
            cols = st.columns(max_cols)
            for idx, item in enumerate(current_attachments):
                with cols[idx % max_cols]:
                    if item["type"] == "image":
                        st.image(item["data"], width=45)
                        btn_label = f"✕ {item['name'][:8]}"
                    else:
                        btn_label = f"✕ [문서]{item['name'][:8]}"
                    if st.button(btn_label, key=f"del_chat_att_{idx}", help=f"{item['name']} ({item['size_kb']:.0f}KB) 삭제"):
                        current_attachments.pop(idx)
                        st.rerun()

    # 3. Streamlit 1.61.1 통합 채팅 입력창 (accept_file="multiple" 네이티브 파일 첨부 지원)
    user_input = st.chat_input(
        "Chat PSDongSung에게 물어보기 (파일 첨부 및 Ctrl+V 캡처 지원)",
        accept_file="multiple",
        file_type=["png", "jpg", "jpeg", "webp", "pdf", "hwp", "xlsx", "docx", "pptx"],
        key="chat_input_widget"
    )

    if user_input:
        prompt = ""
        # ChatInputValue 반환값 처리 (user_input.text 또는 dict 또는 str)
        if hasattr(user_input, "text"):
            prompt = user_input.text or ""
        elif isinstance(user_input, dict):
            prompt = user_input.get("text", "")
        elif isinstance(user_input, str):
            prompt = user_input

        # 파일 첨부 처리 (user_input.files 또는 dict)
        submitted_files = []
        if hasattr(user_input, "files") and user_input.files:
            submitted_files = user_input.files
        elif isinstance(user_input, dict) and user_input.get("files"):
            submitted_files = user_input.get("files")

        if submitted_files:
            for f in submitted_files:
                file_bytes = f.getvalue()
                file_hash = calculate_bytes_hash(file_bytes)
                if f.name.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                    import base64
                    b64_data = f"data:{f.type};base64," + base64.b64encode(file_bytes).decode("utf-8")
                    add_attachment(f.name, f.type, b64_data, f.size, "upload", "image", scope="chat")
                else:
                    text_content = parse_uploaded_file(f)
                    add_document_attachment(f.name, f.type, text_content, f.size, file_hash, scope="chat")

        # 공통 API 페이로드 컴파일 (chat 스코프 지정)
        user_content = compile_api_payload(prompt, selected_model_name, scope="chat")
        
        # 대화 세션에 추가 및 화면 출력
        current_chat["messages"].append({"role": "user", "content": user_content})
        
        with st.chat_message("user"):
            for part in user_content:
                if part["type"] == "text":
                    st.write(part["text"])
                elif part["type"] == "image_url":
                    st.image(part["image_url"]["url"], width=250)
                    
        if current_chat["title"] == "새 대화":
            current_chat["title"] = prompt[:12] + "..." if len(prompt) > 12 else (prompt if prompt else "첨부자료 분석")
            
        with st.chat_message("assistant"):
            if not check_openrouter_model_availability(selected_model):
                st.error("현재 선택한 AI 모델을 OpenRouter에서 사용할 수 없습니다. 관리자에게 모델 설정을 확인해 주세요.")
                st.stop()
            client = get_openrouter_client()
            if not client:
                res = "API 키 설정이 완료되지 않았습니다. (.streamlit/secrets.toml 확인 필요)"
                st.write(res)
                tokens_count = 0
            else:
                with st.spinner(f"[{selected_model_name}] 답변 생성 중..."):
                    try:
                        api_messages = [{"role": "system", "content": "당신은 교사를 보조하는 유능한 AI 에이전트 'Chat PSDongSung'입니다."}]
                        
                        # 이전 대화 목록 히스토리 메시지 팩킹
                        for m in current_chat["messages"][:-1]:
                            api_messages.append({"role": m["role"], "content": m["content"]})
                            
                        # 현재 메시지 추가
                        api_messages.append({"role": "user", "content": user_content})
                        
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
                            
                        # 성공 시 첨부자료 큐 클리어 및 uploader key 갱신으로 상태 완전 초기화
                        st.session_state.attachments["chat"] = []
                        chat_key = f"uploader_chat_{st.session_state.uploader_key_chat}"
                        if chat_key in st.session_state:
                            del st.session_state[chat_key]
                        st.session_state.uploader_key_chat += 1
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
    with st.expander("공통 참조자료", expanded=False):
        st.caption("과목별 성취기준이나 참고자료가 있을 경우 첨부하세요.")
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
                ref_texts.append(f"\n\n[공통참조문서: {g_file.name}]\n" + parse_uploaded_file(g_file))
            global_ref_text = "".join(ref_texts)
            st.success(f"총 {len(global_ref_files)}개의 공통 참조 문서가 반영되었습니다.")

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
        
    student_memo = st.text_area("학생 관찰 내용 및 키워드", value=default_memo, height=95, placeholder="수업 참여도, 수행평가 과정, 특기사항 등 입력")
    
    render_attachments_panel(uploader_key=f"uploader_std_{st.session_state.uploader_key_std}", scope="student")
    
    if st.button("초안 생성", type="primary", use_container_width=True):
        if student_id:
            info_str = f"학번: {student_id}"
            
            if not check_openrouter_model_availability(selected_model):
                st.error("현재 선택한 AI 모델을 OpenRouter에서 사용할 수 없습니다. 관리자에게 모델 설정을 확인해 주세요.")
                st.stop()
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
                            "3. [공통 참조 자료 반영]: 제시된 [공통 교육과정/성취기준 참조 자료]의 성취기준 및 교과역량을 적극 반영하여 작성하세요.\n\n"
                            "4. [영역별 작성 지침]:\n"
                            f"   - 현재 작성 영역: {record_type}\n"
                            "   - 교과세특인 경우: 성취수준 + 수행 과정 및 결과(과제, 분석내용, 도구) + 교과 핵심역량 + 교사 총평 구조 포함.\n"
                            "   - 창체/행발인 경우: 행동 특성, 공동체 의식, 주도적 활동 및 변화 모습을 구체적으로 진술.\n\n"
                            "5. [어조 및 분량]:\n"
                            "   - 문장 끝은 반드시 '~함', '~임' 어조로만 작성하세요.\n"
                            "   - 영역별 권장 바이트 한도 내에서 문장을 구체적이고 풍성하게 기술하세요.\n\n"
                            "6. [기재 금지어]: 대회, 수상, 외부 기관명, 공인어학성적, 사교육 관련 내용 절대 언급 금지."
                        )
                        
                        # 프롬프트 구성 및 공통 첨부자료 빌더 통합
                        prompt_base = f"[작성 정보]: {info_str} ({record_type})\n[학생 관찰 메모]: {student_memo}"
                        if global_ref_text:
                            prompt_base += f"\n[공통 교육과정/성취기준 참조 자료]:\n{global_ref_text}"
                            
                        user_payload = compile_api_payload(prompt_base, selected_model_name, scope="student")
                        
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
                        st.session_state.attachments["student"] = []
                        std_key = f"uploader_std_{st.session_state.uploader_key_std}"
                        if std_key in st.session_state:
                            del st.session_state[std_key]
                        st.session_state.uploader_key_std += 1
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
    
    curr_eval_idx = st.session_state.current_eval_idx
    if curr_eval_idx is not None and curr_eval_idx < len(st.session_state.eval_records):
        curr_eval_rec = st.session_state.eval_records[curr_eval_idx]
        default_eval_text = curr_eval_rec.get("target_text", "")
        default_eval_result = curr_eval_rec.get("result", "")
    else:
        default_eval_text = st.session_state.get("eval_target_text", "")
        default_eval_result = st.session_state.get("eval_result", "")
    
    col_e1, col_e2 = st.columns([3, 1])
    with col_e2:
        if default_eval_result or default_eval_text:
            if st.button("새로 검수하기", use_container_width=True):
                eval_key = f"uploader_eval_{st.session_state.uploader_key_eval}"
                if eval_key in st.session_state:
                    del st.session_state[eval_key]
                st.session_state.uploader_key_eval += 1
                st.session_state.attachments["eval"] = []
                st.session_state.current_eval_idx = None
                st.session_state.eval_result = ""
                st.session_state.eval_target_text = ""
                st.session_state.eval_text_widget = ""
                st.rerun()

    eval_input_text = st.text_area(
        "검수할 생기부 문장 직접 입력",
        height=140,
        placeholder="검수하고자 하는 생기부 특기사항 문단을 복사해서 붙여넣으세요.",
        key="eval_text_widget"
    )
    
    # 공통 첨부파일 패널 렌더링 (eval 스코프 지정 - 입력창 바로 아래 배치)
    render_attachments_panel(uploader_key=f"uploader_eval_{st.session_state.uploader_key_eval}", scope="eval")
    
    target_eval_text = eval_input_text
    
    if target_eval_text:
        c_cnt, b_cnt = calculate_neis_bytes(target_eval_text)
        st.caption(f"검수 대상 분량: {c_cnt}자 / **{b_cnt} Byte** (NEIS 기준)")
        
    if st.button("생기부 정밀 진단 시작", type="primary", use_container_width=True):
        if target_eval_text or any(item["type"] == "image" for item in st.session_state.attachments["eval"]) or any(item["type"] == "document" for item in st.session_state.attachments["eval"]):
            if not check_openrouter_model_availability(selected_model):
                st.error("현재 선택한 AI 모델을 OpenRouter에서 사용할 수 없습니다. 관리자에게 모델 설정을 확인해 주세요.")
                st.stop()
            client = get_openrouter_client()
            if not client:
                st.error("API 키 설정이 완료되지 않았습니다. (.streamlit/secrets.toml 확인 필요)")
            else:
                with st.spinner(f"[{selected_model_name}] 생기부 정밀 분석 및 오류 검수 진행 중..."):
                    try:
                        eval_system_prompt = (
                            "당신은 대한민국 학교생활기록부 정밀 검수 평가관입니다.\n"
                            "제출된 생기부 텍스트 및 첨부자료를 분석하여 아래 구조에 맞춰 상세히 평가 리포트를 작성하세요:\n\n"
                            "1. 지침 위반 및 기재 금지어 적발: (대회, 수상, 외부기관, 어학성적, 사교육 유발 요소 여부 적발)\n"
                            "2. 문체 및 오탈자/비문 진단: (학생 이름/주어 시작 문장 오남용 여부, '~함/임' 어조 미준수 여부, 맞춤법 적발)\n"
                            "3. 작성의 장점 (강점): (구체적 수행과정, 도구 활용, 성취수준 표현 우수성 진술)\n"
                            "4. 작성의 단점 및 보완점: (추상적이거나 단순 총평에 그친 부분 지적)\n"
                            "5. 최종 개선/수정 제안 문장: (지침을 완벽히 준수한 최종 완성 문단 제시)\n"
                        )
                        
                        prompt_base = f"[검수 대상 생기부 텍스트]:\n{target_eval_text}"
                        user_payload = compile_api_payload(prompt_base, selected_model_name, scope="eval")
                        
                        response = client.chat.completions.create(
                            model=selected_model,
                            messages=[
                                {"role": "system", "content": eval_system_prompt},
                                {"role": "user", "content": user_payload}
                            ]
                        )
                        eval_res_str = response.choices[0].message.content
                        
                        tokens_count = 0
                        if hasattr(response, 'usage') and response.usage:
                            tokens_count = response.usage.total_tokens
                            st.session_state.total_tokens_used += tokens_count
                        
                        # 자동 제목 생성 (12~18자)
                        if target_eval_text and target_eval_text.strip():
                            clean_text = target_eval_text.strip().replace("\n", " ")
                            title_str = clean_text[:15] + "..." if len(clean_text) > 15 else clean_text
                        else:
                            title_str = "첨부자료 검토"

                        from datetime import datetime
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

                        if curr_eval_idx is not None and curr_eval_idx < len(st.session_state.eval_records):
                            old_rec = st.session_state.eval_records[curr_eval_idx]
                            created_at_val = old_rec.get("created_at", old_rec.get("updated_at", now_str))
                            new_eval_record = {
                                "title": title_str,
                                "target_text": target_eval_text,
                                "result": eval_res_str,
                                "model": selected_model_name,
                                "tokens": tokens_count,
                                "created_at": created_at_val,
                                "updated_at": now_str
                            }
                            st.session_state.eval_records[curr_eval_idx] = new_eval_record
                        else:
                            new_eval_record = {
                                "title": title_str,
                                "target_text": target_eval_text,
                                "result": eval_res_str,
                                "model": selected_model_name,
                                "tokens": tokens_count,
                                "created_at": now_str,
                                "updated_at": now_str
                            }
                            st.session_state.eval_records.append(new_eval_record)
                            st.session_state.current_eval_idx = len(st.session_state.eval_records) - 1

                        st.session_state.eval_result = eval_res_str
                        st.session_state.eval_target_text = target_eval_text
                        st.session_state.eval_text_widget = target_eval_text

                        # 성공 시 첨부 파일 리셋 및 uploader key 갱신으로 상태 완전 초기화
                        st.session_state.attachments["eval"] = []
                        eval_key = f"uploader_eval_{st.session_state.uploader_key_eval}"
                        if eval_key in st.session_state:
                            del st.session_state[eval_key]
                        st.session_state.uploader_key_eval += 1
                        
                        st.rerun()
                    except Exception as e:
                        st.error(f"진단 중 오류 발생: {str(e)}")
        else:
            st.warning("검수할 파일이나 텍스트를 입력해 주세요.")

    if default_eval_result:
        st.divider()
        st.markdown("### AI 생기부 정밀 진단 결과")
        if curr_eval_idx is not None and curr_eval_idx < len(st.session_state.eval_records):
            rec_info = st.session_state.eval_records[curr_eval_idx]
            st.caption(f"검토 모델: **{rec_info.get('model', selected_model_name)}** | 사용 토큰: **{rec_info.get('tokens', 0):,} Tokens**")
        st.markdown(default_eval_result)
