def hide_streamlit_chrome():
    import streamlit as st
    st.markdown("""
    <style>
    /* (1) 상단 툴바 계열만 제거 — Share/Star/Edit/GitHub/⋮/Deploy */
    div[data-testid="stToolbar"]{display:none !important;}
    [data-testid="stToolbarActions"]{display:none !important;}
    [data-testid="stActionButtonIcon"]{display:none !important;}
    [data-testid="stAppDeployButton"], .stAppDeployButton{display:none !important;}
    #MainMenu{display:none !important;}
    [data-testid="stDecoration"]{display:none !important;}

    /* (2) header는 절대 display:none 하지 말고, 투명 처리만 한다 */
    header[data-testid="stHeader"]{
        background:transparent !important;
        box-shadow:none !important;
        border:none !important;
        height:2.75rem !important;   /* 0으로 만들면 사이드바 토글이 잘림 */
    }

    /* (3) 사이드바 토글/네비게이션은 강제로 보이게 — Streamlit 버전별 셀렉터 전부 커버 */
    [data-testid="stSidebar"],
    [data-testid="stSidebarNav"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stExpandSidebarButton"]{
        display:block !important;
        visibility:visible !important;
        opacity:1 !important;
        pointer-events:auto !important;
        z-index:999999 !important;
    }

    /* (4) 하단 푸터 / 상태 위젯 / 배지 */
    footer{display:none !important;}
    [data-testid="stStatusWidget"]{display:none !important;}
    [class*="viewerBadge"]{display:none !important;}
    </style>
    """, unsafe_allow_html=True)
