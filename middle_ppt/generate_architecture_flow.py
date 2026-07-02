"""
Tro-Fit SW Architecture Flowchart Generator
===========================================
Design System: Meta-inspired light theme (white canvas, rounded boxes, thin hairlines)
Font: Pretendard ONLY
Slide Ratio: 16:9 (33.87 cm × 19.05 cm)
Purpose: Visualize the step-by-step operating flow with icons/blocks and arrows.
"""

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# ─────────────────────────────────────────────
#  DESIGN SYSTEM TOKENS (Meta Commerce Light)
# ─────────────────────────────────────────────
C = {
    "canvas":           RGBColor(0xFF, 0xFF, 0xFF),   # White background
    "surface_soft":     RGBColor(0xF6, 0xF8, 0xFA),   # Light gray card bg
    "surface_card":     RGBColor(0xFF, 0xFF, 0xFF),   # Inner white card bg
    "hairline":         RGBColor(0xE1, 0xE4, 0xE6),   # 1px border
    "hairline_soft":    RGBColor(0xEA, 0xEA, 0xEA),   # Soft border
    "hairline_strong":  RGBColor(0xD1, 0xD5, 0xDA),   # Darker border
    "ink":              RGBColor(0x1F, 0x23, 0x28),   # Primary text
    "slate":            RGBColor(0x57, 0x60, 0x6A),   # Secondary text
    "steel":            RGBColor(0x8C, 0x95, 0x9F),   # Caption text
    "on_dark":          RGBColor(0xFF, 0xFF, 0xFF),   # White on colored bg
    "primary":          RGBColor(0x01, 0x66, 0xFF),   # Meta Cobalt
    "primary_soft":     RGBColor(0xE8, 0xF2, 0xFF),   # Soft Cobalt bg
    "tro_teal":         RGBColor(0x00, 0xB8, 0x82),   # Tro-Fit Teal
    "tro_soft":         RGBColor(0xE6, 0xF8, 0xF3),   # Soft Teal bg
    "accent_orange":    RGBColor(0xFF, 0x6D, 0x00),   # Flow Step Orange
}

FONT_NAME = "Pretendard"

SLIDE_W = Cm(33.87)
SLIDE_H = Cm(19.05)

MARGIN_L = Cm(1.5)
MARGIN_R = Cm(1.5)
CONTENT_W = SLIDE_W - MARGIN_L - MARGIN_R

# Positions
TOP_NAV_H    = Cm(1.10)
CHAPTER_TOP  = Cm(1.35)
DIVIDER_TOP  = Cm(2.00)
TITLE_TOP    = Cm(2.15)
SUBTITLE_TOP = Cm(3.20)
CONTENT_TOP  = Cm(4.00)
FOOTER_TOP   = Cm(17.85)

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs

def blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])

def set_bg(slide, color: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def draw_rect(slide, l, t, w, h, fill_color=None, line_color=None, line_pt=1.0):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape

def add_tb(slide, l, t, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT, spc=0.0, line_sp=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = True
    p   = tf.paragraphs[0]
    p.alignment = align
    if line_sp:
        p.line_spacing = line_sp
    run = p.add_run()
    run.text           = text
    run.font.name      = FONT_NAME
    run.font.size      = size
    run.font.color.rgb = color
    run.font.bold      = bold
    if spc:
        run._r.get_or_add_rPr().set("spc", str(int(spc * 100)))
    return box

def draw_arrow_right(slide, l, t, w, h=Cm(0.3), color=C["primary"]):
    # MSO_SHAPE.RIGHT_ARROW is 33
    arrow = slide.shapes.add_shape(33, l, t, w, h)
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = color
    arrow.line.fill.background()
    return arrow

def draw_arrow_down(slide, l, t, w=Cm(0.3), h=Cm(1.0), color=C["primary"]):
    # MSO_SHAPE.DOWN_ARROW is 35
    arrow = slide.shapes.add_shape(35, l, t, w, h)
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = color
    arrow.line.fill.background()
    return arrow

def draw_arrow_left(slide, l, t, w, h=Cm(0.3), color=C["primary"]):
    # MSO_SHAPE.LEFT_ARROW is 36
    arrow = slide.shapes.add_shape(36, l, t, w, h)
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = color
    arrow.line.fill.background()
    return arrow

def draw_arrow_up(slide, l, t, w=Cm(0.3), h=Cm(1.0), color=C["primary"]):
    # MSO_SHAPE.UP_ARROW is 34
    arrow = slide.shapes.add_shape(34, l, t, w, h)
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = color
    arrow.line.fill.background()
    return arrow

def draw_chrome(slide, chapter: str, slide_num: str, source: str = ""):
    draw_rect(slide, 0, 0, SLIDE_W, TOP_NAV_H, fill_color=C["canvas"])
    draw_rect(slide, 0, TOP_NAV_H, SLIDE_W, Pt(1), fill_color=C["hairline"])
    add_tb(slide, MARGIN_L, Pt(4), Cm(10), TOP_NAV_H, "TRO-FIT", Pt(11), C["ink"], bold=True, spc=1.5)
    add_tb(slide, SLIDE_W - Cm(3.8), Pt(4), Cm(3.5), TOP_NAV_H, slide_num, Pt(9), C["steel"], align=PP_ALIGN.RIGHT)

    # Chapter
    add_tb(slide, MARGIN_L, CHAPTER_TOP, CONTENT_W, Cm(0.6), chapter.upper(), Pt(8.5), C["primary"], bold=True, spc=1.8)
    draw_rect(slide, MARGIN_L - Cm(0.35), CHAPTER_TOP + Pt(2), Pt(4), Pt(11), fill_color=C["primary"])

    # Divider
    draw_rect(slide, MARGIN_L, DIVIDER_TOP, CONTENT_W, Pt(1), fill_color=C["hairline"])

    # Footer
    draw_rect(slide, MARGIN_L, FOOTER_TOP, CONTENT_W, Pt(1), fill_color=C["hairline"])
    if source:
        add_tb(slide, MARGIN_L, FOOTER_TOP + Pt(3), CONTENT_W * 0.8, Cm(0.55), source, Pt(7.5), C["steel"])

# ─────────────────────────────────────────────
#  BUILD FLOW SLIDE
# ─────────────────────────────────────────────

def build_architecture_flow_slide(prs: Presentation):
    slide = blank_slide(prs)
    set_bg(slide, C["canvas"])
    draw_chrome(slide, "System Architecture", "04c / 09", "출처: 2026 피우다프로젝트 WBS (로컬 환경 버전) 데이터 및 제어 흐름")

    # Titles
    add_tb(slide, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.0), "Tro-Fit 데이터 및 서비스 흐름도 (SW 동작 구성도)", Pt(26), C["ink"], bold=True)
    add_tb(slide, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.6), "사용자 동작 측정부터 실시간 ROM 진단, RAG 기반 안무 처방 및 TV 연동 운동 실행까지의 전체 순서도", Pt(12), C["slate"])

    # ─────────────────────────────────────────
    # 1. COMPONENT BLOCKS CONFIGURATION (XY)
    # ─────────────────────────────────────────
    
    # 1.1 [사용자] Block (Left-Top)
    user_x, user_y, user_w, user_h = Cm(1.5), Cm(6.5), Cm(2.4), Cm(2.8)
    draw_rect(slide, user_x, user_y, user_w, user_h, fill_color=C["surface_soft"], line_color=C["hairline_strong"], line_pt=1.0)
    add_tb(slide, user_x + Pt(4), user_y + Pt(6), user_w - Pt(8), Pt(14), "사용자", Pt(11), C["ink"], bold=True, align=PP_ALIGN.CENTER)
    add_tb(slide, user_x + Pt(4), user_y + Pt(26), user_w - Pt(8), Pt(40), "(65세 이상\n노년층)", Pt(8.5), C["slate"], align=PP_ALIGN.CENTER)

    # 1.2 [스마트폰 앱 (React Native / Expo)] Block
    phone_x, phone_y, phone_w, phone_h = Cm(5.0), Cm(4.5), Cm(5.8), Cm(6.8)
    draw_rect(slide, phone_x, phone_y, phone_w, phone_h, fill_color=C["primary_soft"], line_color=C["primary"], line_pt=1.0)
    add_tb(slide, phone_x + Pt(10), phone_y + Pt(8), phone_w - Pt(20), Pt(16), "스마트폰 앱", Pt(12), C["primary"], bold=True)
    add_tb(slide, phone_x + Pt(10), phone_y + Pt(24), phone_w - Pt(20), Pt(12), "(React Native / Expo)", Pt(8.5), C["slate"])
    
    # Inner: On-device AI (MediaPipe / TFLite)
    ai_x, ai_y, ai_w, ai_h = phone_x + Cm(0.4), phone_y + Cm(1.4), phone_w - Cm(0.8), Cm(4.8)
    draw_rect(slide, ai_x, ai_y, ai_w, ai_h, fill_color=C["canvas"], line_color=C["hairline_strong"], line_pt=1.0)
    add_tb(slide, ai_x + Pt(8), ai_y + Pt(6), ai_w - Pt(16), Pt(14), "On-device AI 엔진", Pt(10), C["ink"], bold=True)
    ai_stack = (
        "• MediaPipe Pose (33개 3D)\n"
        "• RTMPose 모델 성능 검증\n"
        "• TFLite (INT8 양자화 경량화)\n"
        "• FMS 관절 각도 계산\n"
        "※ 촬영 영상 즉시 파기"
    )
    add_tb(slide, ai_x + Pt(8), ai_y + Pt(22), ai_w - Pt(16), Cm(3.6), ai_stack, Pt(8), C["slate"], line_sp=Pt(11))

    # 1.3 [API 서버 (FastAPI)] Block
    api_x, api_y, api_w, api_h = Cm(12.5), Cm(4.5), Cm(5.2), Cm(2.8)
    draw_rect(slide, api_x, api_y, api_w, api_h, fill_color=C["tro_soft"], line_color=C["tro_teal"], line_pt=1.0)
    add_tb(slide, api_x + Pt(10), api_y + Pt(8), api_w - Pt(20), Pt(16), "API 서버 (FastAPI)", Pt(12), C["tro_teal"], bold=True)
    api_stack = (
        "• RESTful API / JSON\n"
        "• JWT 인증 & WebSocket\n"
        "• 보호자 카카오톡 알림"
    )
    add_tb(slide, api_x + Pt(10), api_y + Pt(26), api_w - Pt(20), Cm(1.8), api_stack, Pt(8.5), C["slate"], line_sp=Pt(12))

    # 1.4 [비동기 큐 (Celery / Redis)] Block
    queue_x, queue_y, queue_w, queue_h = Cm(12.5), Cm(8.5), Cm(5.2), Cm(2.8)
    draw_rect(slide, queue_x, queue_y, queue_w, queue_h, fill_color=C["surface_soft"], line_color=C["hairline_strong"], line_pt=1.0)
    add_tb(slide, queue_x + Pt(10), queue_y + Pt(8), queue_w - Pt(20), Pt(16), "비동기 분석 (Celery / Redis)", Pt(11), C["ink"], bold=True)
    queue_stack = (
        "• LLM Fine-tuning 안무 자동 생성\n"
        "• FFmpeg 오디오/VOD 인코딩\n"
        "• 시계열 이상 탐지 예보"
    )
    add_tb(slide, queue_x + Pt(10), queue_y + Pt(26), queue_w - Pt(20), Cm(1.8), queue_stack, Pt(8.5), C["slate"], line_sp=Pt(12))

    # 1.5 [데이터베이스 & 스토리지] Blocks (Right Side Stack)
    db_x = Cm(19.2)
    db_w, db_h = Cm(4.8), Cm(2.0)
    
    # 1.5.1 PostgreSQL
    pg_y = Cm(4.5)
    draw_rect(slide, db_x, pg_y, db_w, db_h, fill_color=C["canvas"], line_color=C["hairline_strong"], line_pt=1.0)
    add_tb(slide, db_x + Pt(8), pg_y + Pt(6), db_w - Pt(16), Pt(14), "PostgreSQL (RDB)", Pt(10), C["ink"], bold=True)
    add_tb(slide, db_x + Pt(8), pg_y + Pt(22), db_w - Pt(16), Pt(24), "• 사용자 데이터\n• 누적 ROM 시계열 로그", Pt(8), C["slate"], line_sp=Pt(11))

    # 1.5.2 ChromaDB
    chroma_y = Cm(6.9)
    draw_rect(slide, db_x, chroma_y, db_w, db_h, fill_color=C["canvas"], line_color=C["hairline_strong"], line_pt=1.0)
    add_tb(slide, db_x + Pt(8), chroma_y + Pt(6), db_w - Pt(16), Pt(14), "ChromaDB (Vector DB)", Pt(10), C["ink"], bold=True)
    add_tb(slide, db_x + Pt(8), chroma_y + Pt(22), db_w - Pt(16), Pt(24), "• 36개 임상 지침 가이드\n• RAG 기반 유사도 검색", Pt(8), C["slate"], line_sp=Pt(11))

    # 1.5.3 MinIO
    minio_y = Cm(9.3)
    draw_rect(slide, db_x, minio_y, db_w, db_h, fill_color=C["canvas"], line_color=C["hairline_strong"], line_pt=1.0)
    add_tb(slide, db_x + Pt(8), minio_y + Pt(6), db_w - Pt(16), Pt(14), "MinIO (Object Storage)", Pt(10), C["ink"], bold=True)
    add_tb(slide, db_x + Pt(8), minio_y + Pt(22), db_w - Pt(16), Pt(24), "• 안무 JSON 패킷\n• 로컬 S3 API", Pt(8), C["slate"], line_sp=Pt(11))

    # 1.6 [스마트 TV (WebOS / Tizen)] Block (Right-End)
    tv_x, tv_y, tv_w, tv_h = Cm(25.5), Cm(5.8), Cm(6.8), Cm(4.2)
    draw_rect(slide, tv_x, tv_y, tv_w, tv_h, fill_color=C["primary_soft"], line_color=C["primary"], line_pt=1.0)
    add_tb(slide, tv_x + Pt(12), tv_y + Pt(10), tv_w - Pt(24), Pt(18), "스마트 TV (거실 화면)", Pt(13), C["primary"], bold=True)
    add_tb(slide, tv_x + Pt(12), tv_y + Pt(28), tv_w - Pt(24), Pt(12), "(WebOS / Tizen TV App)", Pt(8.5), C["slate"])
    tv_details = (
        "• Chromecast 수신 연동\n"
        "• HDMI 미러링 및 Wake Lock 활성\n"
        "• Lottie 캐릭터 안무 시각화 재생\n"
        "• TV 리모컨 조작 피드백 송수신"
    )
    add_tb(slide, tv_x + Pt(12), tv_y + Pt(42), tv_w - Pt(24), Cm(2.4), tv_details, Pt(8.5), C["slate"], line_sp=Pt(12))

    # ─────────────────────────────────────────
    # 2. DEVSECOPS & INFRASTRUCTURE ROW (BOTTOM)
    # ─────────────────────────────────────────
    infra_y = Cm(14.0)
    infra_h = Cm(2.2)
    draw_rect(slide, MARGIN_L, infra_y, CONTENT_W, infra_h, fill_color=C["surface_soft"], line_color=C["hairline"], line_pt=1.0)
    add_tb(slide, MARGIN_L + Pt(14), infra_y + Pt(8), CONTENT_W - Pt(28), Pt(14), "개발 인프라 및 신뢰성 검증 환경 (WBS 기재 도구)", Pt(10), C["slate"], bold=True, spc=1.0)

    infra_w = (CONTENT_W - Cm(0.9)) / 4
    infra_items = [
        ("CI/CD 빌드 자동화", "GitHub Actions\n로컬 Docker Compose 빌드 자동화"),
        ("멀티컨테이너 모니터링", "Prometheus + Grafana\nDocker 리소스 모니터링"),
        ("성능 및 부하 테스트", "Locust\nFastAPI API 서버 로컬 부하 테스트"),
        ("정적 및 동적 보안 진단", "OWASP ZAP 로컬 취약점 스캔\nBandit Python 코드 정적 보안 분석")
    ]
    for ii, (ititle, idesc) in enumerate(infra_items):
        ix = MARGIN_L + ii * (infra_w + Cm(0.3)) + Pt(10)
        draw_rect(slide, MARGIN_L + ii * (infra_w + Cm(0.3)) + Cm(0.2), infra_y + Cm(0.7), infra_w - Cm(0.4), Cm(1.3), fill_color=C["canvas"], line_color=C["hairline"], line_pt=1.0)
        add_tb(slide, ix + Pt(6), infra_y + Cm(0.75), infra_w - Cm(0.6), Pt(12), ititle, Pt(9), C["ink"], bold=True)
        add_tb(slide, ix + Pt(6), infra_y + Cm(1.15), infra_w - Cm(0.6), Pt(18), idesc.replace('\n', '  |  '), Pt(7.5), C["slate"])

    # ─────────────────────────────────────────
    # 3. DRAW CONNECTORS & LABELED STEPS (① ~ ⑧)
    # ─────────────────────────────────────────

    # ① User -> Smartphone (관절 측정)
    draw_arrow_right(slide, user_x + user_w, user_y + Cm(0.8), Cm(1.1), color=C["accent_orange"])
    add_tb(slide, user_x + user_w, user_y - Cm(0.6), Cm(2.2), Cm(1.2), "① 관절 영상 촬영\n3D 랜드마크 추출", Pt(7.5), C["accent_orange"], bold=True)

    # ② Smartphone -> FastAPI (진단 데이터 전송)
    draw_arrow_right(slide, phone_x + phone_w, phone_y + Cm(0.8), Cm(1.7), color=C["accent_orange"])
    add_tb(slide, phone_x + phone_w + Cm(0.1), phone_y - Cm(0.6), Cm(1.5), Cm(1.2), "② ROM 진단\nJSON 전송", Pt(7.5), C["accent_orange"], bold=True)

    # ③ FastAPI -> PostgreSQL (데이터 적재)
    draw_arrow_right(slide, api_x + api_w, pg_y + Cm(0.8), Cm(1.5), color=C["primary"])
    add_tb(slide, api_x + api_w + Cm(0.1), pg_y - Cm(0.5), Cm(1.3), Cm(1.0), "③ 누적 데이터\n저장 (SQL)", Pt(7.5), C["primary"], bold=True)

    # ④ FastAPI -> Celery/Redis (비동기 처방 요청)
    draw_arrow_down(slide, api_x + api_w / 2 - Cm(0.15), api_y + api_h, w=Cm(0.3), h=Cm(1.2), color=C["primary"])
    add_tb(slide, api_x + api_w / 2 + Cm(0.2), api_y + api_h + Cm(0.3), Cm(2.0), Cm(0.8), "④ 비동기 안무\n생성 대기열 전송", Pt(7.5), C["primary"], bold=True)

    # ⑤ Celery -> ChromaDB (RAG 임상지침 가이드 검색)
    draw_arrow_right(slide, queue_x + queue_w, chroma_y + Cm(0.8), Cm(1.5), color=C["tro_teal"])
    add_tb(slide, queue_x + queue_w + Cm(0.1), chroma_y - Cm(0.5), Cm(1.3), Cm(1.0), "⑤ 임상지침 RAG\n유사도 검색", Pt(7.5), C["tro_teal"], bold=True)

    # ⑥ Celery -> MinIO (안무 생성 및 저장)
    draw_arrow_right(slide, queue_x + queue_w, minio_y + Cm(0.8), Cm(1.5), color=C["tro_teal"])
    add_tb(slide, queue_x + queue_w + Cm(0.1), minio_y - Cm(0.5), Cm(1.3), Cm(1.0), "⑥ 안무 생성\nJSON 저장", Pt(7.5), C["tro_teal"], bold=True)

    # ⑦ FastAPI -> Smartphone (안무 메타데이터/URL 리턴)
    draw_arrow_left(slide, phone_x + phone_w, phone_y + Cm(1.8), Cm(1.7), color=C["accent_orange"])
    add_tb(slide, phone_x + phone_w + Cm(0.1), phone_y + Cm(2.1), Cm(1.5), Cm(1.2), "⑦ 맞춤 안무\nJSON URL 전달", Pt(7.5), C["accent_orange"], bold=True)

    # ⑧ Smartphone -> Smart TV (TV 캐스팅 / 미러링 연동)
    draw_arrow_right(slide, phone_x + phone_w, phone_y + Cm(5.0), Cm(14.7), color=C["accent_orange"])
    add_tb(slide, phone_x + phone_w + Cm(2.0), phone_y + Cm(4.3), Cm(10.0), Cm(0.6), "⑧ TV 연동 실행 (Chromecast 캐스팅 및 HDMI 미러링 전송 → Lottie 애니메이션 및 오디오 재생)", Pt(8.5), C["accent_orange"], bold=True)

    # Footer Stripe Accent
    draw_rect(slide, 0, SLIDE_H - Pt(4), SLIDE_W, Pt(4), fill_color=C["primary"])

if __name__ == "__main__":
    prs = new_prs()
    build_architecture_flow_slide(prs)
    out_file = "c:/workspace/trofit/middle_ppt/tro_fit_sw_architecture_flow.pptx"
    prs.save(out_file)
    print(f"[OK] Saved SW Architecture Flow slide to: {out_file}")
