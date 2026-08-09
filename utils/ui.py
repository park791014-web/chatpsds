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
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stExpandSidebarButton"]{
        display:block !important;
        visibility:visible !important;
        opacity:1 !important;
        pointer-events:auto !important;
        z-index:9999 !important;
    }

    /* (4) 하단 푸터 / 상태 위젯 / 배지 */
    footer{display:none !important;}
    [data-testid="stStatusWidget"]{display:none !important;}
    [class*="viewerBadge"]{display:none !important;}

    /* (5) 팝업, 다이얼로그, 툴팁, 드롭다운 등 오버레이 UI 정상 표시 및 최상단 보장 */
    div[data-baseweb="popover"],
    div[role="dialog"],
    div[role="menu"],
    [data-testid="stTooltipHoverTarget"],
    [data-testid="stTooltipContent"],
    div[class*="stTooltip"],
    div[class*="stDialog"],
    div[class*="stPortal"] {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
        z-index: 9999999 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# Ctrl+V 클립보드 이미지 복사/붙여넣기 리스너 (CCv2 적용)
# ==========================================
_PASTE_LISTENER = None

def get_paste_listener():
    global _PASTE_LISTENER
    import streamlit as st
    if _PASTE_LISTENER is None:
        PASTE_HTML = """<div style="display:none;" id="paste-listener-container"></div>"""
        
        # 브라우저 단에서 1920px(Full HD) 스케일링 및 PNG 무손실 압축으로 텍스트 선명도 극대화
        PASTE_JS = """
        export default function (component) {
          const { setTriggerValue } = component;
          
          const handlePaste = (e) => {
            const items = (e.clipboardData || e.originalEvent.clipboardData).items;
            for (const item of items) {
              if (item.kind === 'file' && item.type.startsWith('image/')) {
                const file = item.getAsFile();
                const reader = new FileReader();
                reader.onload = function (event) {
                  const img = new Image();
                  img.onload = function () {
                    const max_size = 1920;
                    let width = img.width;
                    let height = img.height;
                    
                    if (width > max_size || height > max_size) {
                      if (width > height) {
                        height = Math.round((height * max_size) / width);
                        width = max_size;
                      } else {
                        width = Math.round((width * max_size) / height);
                        height = max_size;
                      }
                    }
                    
                    const canvas = document.createElement("canvas");
                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext("2d");
                    ctx.drawImage(img, 0, 0, width, height);
                    
                    // 텍스트 깨짐을 방지하는 무손실 PNG 포맷 우선 적용 (미인식 시 원본 타입 사용)
                    const outputType = file.type || "image/png";
                    const compressedBase64 = canvas.toDataURL(outputType);
                    
                    setTriggerValue("pasted_image", {
                      name: file.name || "pasted_image.png",
                      type: outputType,
                      size: Math.round(compressedBase64.length * 3 / 4),
                      data: compressedBase64
                    });
                  };
                  img.src = event.target.result;
                };
                reader.readAsDataURL(file);
                e.preventDefault();
                break;
              }
            }
          };
          
          document.addEventListener("paste", handlePaste);
          
          return () => {
            document.removeEventListener("paste", handlePaste);
          };
        }
        """
        _PASTE_LISTENER = st.components.v2.component(
            "clipboard_paste_listener",
            html=PASTE_HTML,
            js=PASTE_JS,
        )
    return _PASTE_LISTENER

def paste_listener(key="paste-listener-default", on_pasted_image=None):
    listener = get_paste_listener()
    if on_pasted_image is None:
        on_pasted_image = lambda: None
    return listener(
        key=key,
        on_pasted_image_change=on_pasted_image,
    )
