"""
Tro-Fit Full Presentation Generator — Meta Design System
==========================================================
Design Spec  : Meta Commerce Surface (white canvas, pill buttons, rounded cards)
Font         : Pretendard ONLY  (Bold=700 / SemiBold=600 / Regular=400 / Light=300)
Slide Ratio  : 16:9  →  33.87 cm × 19.05 cm
Total Slides : 9

Platform Goal:
  PRIMARY  = Joint health prevention for seniors
  SECONDARY = Poor joint health causes falls → this platform also prevents falls.
              "Fall prevention" is NOT the headline goal.

Fixed chrome positions (every slide):
  TOP_BAR      : y=0,     h=1.1cm   — service name + slide number
  CHAPTER_TOP  : y=1.35cm           — chapter label (small caps, cobalt)
  DIVIDER_TOP  : y=2.00cm           — 1pt hairline divider
  TITLE_TOP    : y=2.15cm           — H1 headline (large, bold)
  SUBTITLE_TOP : y=4.50cm           — subtitle / lead paragraph
  CONTENT_TOP  : y=5.90cm           — main content area begins
  FOOTER_TOP   : y=17.85cm          — citation + slide number row
"""

from __future__ import annotations
import os
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─────────────────────────────────────────────────────────
#  DESIGN TOKENS  (Meta Commerce System)
# ─────────────────────────────────────────────────────────
C = {
    # Surfaces
    "canvas":      RGBColor(0xFF, 0xFF, 0xFF),   # #FFFFFF — page background
    "soft":        RGBColor(0xF5, 0xF5, 0xF5),   # #F5F5F5 — card / tag bg
    "soft2":       RGBColor(0xEE, 0xEE, 0xEE),   # slightly darker soft
    "ink_deep":    RGBColor(0x0A, 0x13, 0x17),   # #0A1317 — dark promo strip
    "hairline":    RGBColor(0xDD, 0xDD, 0xDD),   # #DDDDDD — 1px divider
    "hairline_s":  RGBColor(0xEA, 0xEA, 0xEA),   # #EAEAEA — soft divider
    # Text
    "ink":         RGBColor(0x1C, 0x1C, 0x1C),   # #1C1C1C — body text
    "charcoal":    RGBColor(0x3E, 0x3E, 0x3E),   # #3E3E3E — secondary text
    "slate":       RGBColor(0x5F, 0x5F, 0x5F),   # #5F5F5F — muted text
    "steel":       RGBColor(0x8A, 0x8A, 0x8A),   # #8A8A8A — caption
    "on_dark":     RGBColor(0xFF, 0xFF, 0xFF),    # white on dark strips
    # Accent
    "primary":     RGBColor(0x01, 0x66, 0xFF),   # #0166FF — cobalt CTA
    "primary_d":   RGBColor(0x00, 0x4D, 0xC7),   # #004DC7 — deep cobalt
    "primary_s":   RGBColor(0xE8, 0xF2, 0xFF),   # #E8F2FF — soft cobalt bg
    # Semantic
    "success":     RGBColor(0x00, 0x9C, 0x44),   # #009C44
    "warning":     RGBColor(0xFF, 0xC3, 0x00),   # #FFC300
    "critical":    RGBColor(0xD8, 0x2C, 0x0C),   # #D82C0C
    "attention":   RGBColor(0xFF, 0x6D, 0x00),   # #FF6D00
    # Tro-Fit brand teal (secondary accent)
    "tro":         RGBColor(0x00, 0xB8, 0x82),   # #00B882 — vitality teal
    "tro_soft":    RGBColor(0xE6, 0xF8, 0xF3),   # teal tint bg
}

FONT = "Pretendard"
FONT_DIR = os.path.join(
    os.path.dirname(__file__),
    "Pretendard-1.3.9", "public", "static"
)

# Slide 16:9
SLIDE_W = Cm(33.87)
SLIDE_H = Cm(19.05)

# Fixed chrome Y positions
TOP_BAR_H    = Cm(1.10)
CHAPTER_TOP  = Cm(1.35)
DIVIDER_TOP  = Cm(2.00)
TITLE_TOP    = Cm(2.15)
SUBTITLE_TOP = Cm(4.50)
CONTENT_TOP  = Cm(5.90)
FOOTER_TOP   = Cm(17.85)

MARGIN_L = Cm(2.2)
MARGIN_R = Cm(2.2)
CONTENT_W = SLIDE_W - MARGIN_L - MARGIN_R


# ─────────────────────────────────────────────────────────
#  PRIMITIVES
# ─────────────────────────────────────────────────────────

def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def set_bg(slide, color: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, l, t, w, h, fill=None, line=None, line_pt=1.0, radius_pt=0):
    """Add a rectangle; radius_pt > 0 for rounded corners (via XML patch)."""
    shape = slide.shapes.add_shape(1, l, t, w, h)
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    if radius_pt > 0:
        emu_r = int(Pt(radius_pt) * 914400 / 12700)
        sp = shape._element
        sp_pr = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}spPr')
        if sp_pr is not None:
            prstGeom = sp_pr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
            if prstGeom is not None:
                sp_pr.remove(prstGeom)
            from lxml import etree
            cust = etree.SubElement(sp_pr,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}custGeom')
            avLst = etree.SubElement(cust,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
            gdLst = etree.SubElement(cust,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}gdLst')
            ahLst = etree.SubElement(cust,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}ahLst')
            cxnLst = etree.SubElement(cust,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}cxnLst')
            rect_el = etree.SubElement(cust,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}rect',
                l='l', t='t', r='r', b='b')
            pathLst = etree.SubElement(cust,
                '{http://schemas.openxmlformats.org/drawingml/2006/main}pathLst')
    return shape


def tb(slide, l, t, w, h, text, size, color, bold=False,
       align=PP_ALIGN.LEFT, spc=0.0, line_sp=None, wrap=True):
    """Single-paragraph textbox."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    if line_sp:
        p.line_spacing = line_sp
    run = p.add_run()
    run.text           = text
    run.font.name      = FONT
    run.font.size      = size
    run.font.color.rgb = color
    run.font.bold      = bold
    if spc:
        run._r.get_or_add_rPr().set("spc", str(int(spc * 100)))
    return box


def multiline(slide, l, t, w, h, lines: list):
    """Multi-paragraph textbox.  lines = list of dicts."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = True
    for i, ld in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ld.get("align", PP_ALIGN.LEFT)
        if "line_sp" in ld:
            p.line_spacing = ld["line_sp"]
        if "sb" in ld:
            p.space_before = ld["sb"]
        if "sa" in ld:
            p.space_after  = ld["sa"]
        run = p.add_run()
        run.text           = ld["text"]
        run.font.name      = FONT
        run.font.size      = ld.get("size", Pt(14))
        run.font.color.rgb = ld.get("color", C["ink"])
        run.font.bold      = ld.get("bold", False)
        if ld.get("spc", 0):
            run._r.get_or_add_rPr().set("spc", str(int(ld["spc"] * 100)))
    return box


def hairline(slide, l, t, w, color=None):
    rect(slide, l, t, w, Pt(1), fill=color or C["hairline"])


def pill_badge(slide, l, t, text, bg, fg, size=Pt(9)):
    """Draw a pill-shaped badge (approximated with rounded rect + text overlay)."""
    badge_w = Cm(2.6)
    badge_h = Pt(18)
    r = rect(slide, l, t, badge_w, badge_h, fill=bg)
    tb(slide, l, t, badge_w, badge_h, text, size, fg, bold=True,
       align=PP_ALIGN.CENTER)
    return r


def stat_card(slide, l, t, w, h, value, label, value_color=None, bg=None):
    """Metric stat card: big value + small label."""
    rect(slide, l, t, w, h, fill=bg or C["soft"],
         line=C["hairline_s"], line_pt=1)
    add_vc = value_color or C["ink"]
    tb(slide, l + Pt(18), t + Pt(14), w - Pt(36), Pt(38),
       value, Pt(28), add_vc, bold=True)
    tb(slide, l + Pt(18), t + Pt(52), w - Pt(36), Pt(22),
       label, Pt(9.5), C["slate"])


def tag(slide, l, t, text, color=None):
    """Small uppercase tag pill."""
    bg = color or C["primary_s"]
    badge_w = Cm(3.0)
    badge_h = Pt(16)
    rect(slide, l, t, badge_w, badge_h, fill=bg)
    tb(slide, l + Pt(6), t + Pt(1), badge_w - Pt(12), badge_h,
       text.upper(), Pt(8), C["primary"], bold=True, spc=1.0)


# ─────────────────────────────────────────────────────────
#  SHARED CHROME  (top bar + chapter + divider + footer)
# ─────────────────────────────────────────────────────────

def chrome(slide, chapter: str, slide_num: str, source: str = ""):
    """Draw the fixed chrome elements present on every slide."""
    # ── Top bar (white strip with bottom hairline)
    rect(slide, 0, 0, SLIDE_W, TOP_BAR_H, fill=C["canvas"])
    hairline(slide, 0, TOP_BAR_H, SLIDE_W, color=C["hairline"])
    # Service name
    tb(slide, MARGIN_L, Pt(4), Cm(10), TOP_BAR_H,
       "TRO-FIT", Pt(11), C["ink"], bold=True, spc=1.5)
    # Slide number
    tb(slide, SLIDE_W - Cm(3.8), Pt(4), Cm(3.5), TOP_BAR_H,
       slide_num, Pt(9), C["steel"], align=PP_ALIGN.RIGHT)

    # ── Chapter label
    tb(slide, MARGIN_L, CHAPTER_TOP, CONTENT_W, Cm(0.6),
       chapter.upper(), Pt(8.5), C["primary"], bold=True, spc=1.8)

    # ── Hairline divider (below chapter)
    hairline(slide, MARGIN_L, DIVIDER_TOP, CONTENT_W, color=C["hairline"])

    # ── Footer hairline + citation
    hairline(slide, MARGIN_L, FOOTER_TOP, CONTENT_W, color=C["hairline"])
    if source:
        tb(slide, MARGIN_L, FOOTER_TOP + Pt(3), CONTENT_W * 0.8, Cm(0.55),
           source, Pt(7.5), C["steel"])

    # Cobalt accent dot next to chapter label
    rect(slide, MARGIN_L - Cm(0.35), CHAPTER_TOP + Pt(2),
         Pt(4), Pt(11), fill=C["primary"])


# ─────────────────────────────────────────────────────────
#  SLIDE 01 — Opening Hook
# ─────────────────────────────────────────────────────────
def slide_01(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Platform Overview",
           slide_num="01 / 09",
           source="출처: 2022 한국소비자원  |  2024 건강보험통계연보  |  2025 통계청")

    # H1 Title
    multiline(s, MARGIN_L, TITLE_TOP, CONTENT_W * 0.62, Cm(2.1),
        lines=[
            {"text": "관절 건강이 낙상을 막습니다.",
             "size": Pt(34), "color": C["ink"], "bold": True, "line_sp": Pt(40)},
            {"text": "Tro-Fit이 그 시작입니다.",
             "size": Pt(34), "color": C["primary"], "bold": True, "line_sp": Pt(40)},
        ])

    # Subtitle
    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W * 0.60, Cm(0.85),
       "관절 기능 저하 → 균형 감각 저하 → 낙상 위험 증가. 예방의 근본은 관절입니다.",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    # Cobalt pill divider line
    rect(s, MARGIN_L, SUBTITLE_TOP + Cm(1.0), Cm(1.2), Pt(3), fill=C["primary"])

    # Body text
    tb(s, MARGIN_L, CONTENT_TOP, CONTENT_W * 0.54, Cm(3.8),
       "65세 이상 어르신의 관절 가동범위(ROM)가 줄어들수록\n"
       "낙상 위험은 기하급수적으로 증가합니다.\n\n"
       "Tro-Fit은 낙상 예방 앱이 아닙니다.\n"
       "AI가 매주 관절 상태를 진단하고, 맞춤 트로트 안무로\n"
       "관절 건강을 근본부터 관리하는 플랫폼입니다.",
       Pt(12), C["charcoal"], line_sp=Pt(20))

    # CTA text link
    tb(s, MARGIN_L, CONTENT_TOP + Cm(4.2), Cm(14), Cm(0.6),
       "→  AI ROM 측정 + RAG 안무 처방 + TV 연동 운동  |  스마트폰 1대로 완성",
       Pt(10.5), C["primary"], bold=True)

    # Right: 3 stat cards
    cw, ch, cg = Cm(8.5), Cm(3.45), Cm(0.4)
    cl = SLIDE_W - MARGIN_R - cw
    stats = [
        ("62.7%",  "65세 이상 낙상 관련 사망 비율",   C["critical"]),
        ("551만원", "노인 1인 연평균 진료비",           C["ink"]),
        ("20.3%",  "초고령사회 진입 비율 (2025)",     C["tro"]),
    ]
    for i, (v, l, vc) in enumerate(stats):
        stat_card(s, cl, CONTENT_TOP + i*(ch+cg), cw, ch, v, l, vc)

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 02 — Problem: 3 Barriers
# ─────────────────────────────────────────────────────────
def slide_02(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Problem Definition",
           slide_num="02 / 09",
           source="출처: 2024 건강보험통계연보 756p  |  (NIA) 2024 디지털정보격차 실태조사  |  보건교육건강증진학회지")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "기존 솔루션이 어르신에게 닿지 못하는 이유",
       Pt(30), C["ink"], bold=True, line_sp=Pt(38))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "근골격계 위기 · 디지털 소외 · 기기 과부하 — 세 장벽의 동시 돌파가 필요합니다.",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    # 3 barrier cards
    card_w = (CONTENT_W - Cm(0.8)) / 3
    card_h = Cm(9.4)
    cards = [
        {
            "tag": "장벽 1",
            "tag_color": C["tro_soft"],
            "title": "근골격계 위기",
            "lines": [
                "65세↑ 다빈도 상병",
                "무릎관절증 4위",
                "척추병증 9위",
                "",
                "관절 기능 저하 →",
                "균형 감각 약화 →",
                "낙상 → 장기 입원",
                "악순환 구조",
            ],
            "stat": "4위",
            "stat_label": "무릎관절증 상병 순위",
            "stat_color": C["critical"],
        },
        {
            "tag": "장벽 2",
            "tag_color": RGBColor(0xFF, 0xF3, 0xE0),
            "title": "디지털 소외",
            "lines": [
                "노인 AI 서비스 경험률",
                "27.8%",
                "(일반 51% 대비 −23.2%p)",
                "",
                "TV 리모컨은 익숙,",
                "스마트폰 앱은 낯설다",
                "",
                "기존 앱은 지속률 0%",
            ],
            "stat": "27.8%",
            "stat_label": "노인 AI 서비스 경험률",
            "stat_color": C["attention"],
        },
        {
            "tag": "장벽 3",
            "tag_color": RGBColor(0xE8, 0xF2, 0xFF),
            "title": "기기 과부하",
            "lines": [
                "실시간 비전 AI →",
                "기기 발열 · 배터리 방전",
                "",
                "매번 카메라 각도",
                "맞추는 피로감",
                "",
                "보급형 폰에서",
                "지속 사용 불가",
            ],
            "stat": "0%",
            "stat_label": "기존 앱 지속 사용률",
            "stat_color": C["primary"],
        },
    ]

    for i, cd in enumerate(cards):
        cl = MARGIN_L + i * (card_w + Cm(0.4))
        ct = CONTENT_TOP
        rect(s, cl, ct, card_w, card_h, fill=C["canvas"],
             line=C["hairline_s"], line_pt=1)
        # tag
        tag_w = Cm(1.8)
        rect(s, cl + Pt(16), ct + Pt(16), tag_w, Pt(16),
             fill=cd["tag_color"])
        tb(s, cl + Pt(16), ct + Pt(16), tag_w, Pt(16),
           cd["tag"], Pt(8), C["charcoal"], bold=True, align=PP_ALIGN.CENTER)
        # title
        tb(s, cl + Pt(16), ct + Pt(38), card_w - Pt(32), Pt(26),
           cd["title"], Pt(15), C["ink"], bold=True, line_sp=Pt(22))
        # stat value
        tb(s, cl + Pt(16), ct + Pt(68), card_w - Pt(32), Pt(38),
           cd["stat"], Pt(28), cd["stat_color"], bold=True)
        tb(s, cl + Pt(16), ct + Pt(104), card_w - Pt(32), Pt(18),
           cd["stat_label"], Pt(9), C["slate"])
        # divider
        hairline(s, cl + Pt(16), ct + Pt(124), card_w - Pt(32))
        # body lines
        body_txt = "\n".join(ld for ld in cd["lines"] if ld)
        tb(s, cl + Pt(16), ct + Pt(132), card_w - Pt(32), Cm(4.0),
           body_txt, Pt(10.5), C["charcoal"], line_sp=Pt(17))

    # Bottom CTA strip
    strip_t = CONTENT_TOP + card_h + Cm(0.35)
    rect(s, MARGIN_L, strip_t, CONTENT_W, Cm(1.05), fill=C["ink_deep"])
    tb(s, MARGIN_L + Pt(16), strip_t + Pt(8), CONTENT_W - Pt(32), Pt(30),
       "Tro-Fit은 이 세 가지 장벽을 동시에 돌파하는 유일한 솔루션입니다.",
       Pt(11.5), C["on_dark"], bold=True)

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 03 — Solution: Async Architecture
# ─────────────────────────────────────────────────────────
def slide_03(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Solution Architecture",
           slide_num="03 / 09")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "진단과 운동을 분리한다 — 비동기 아키텍처",
       Pt(30), C["ink"], bold=True, line_sp=Pt(38))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W * 0.75, Cm(0.75),
       "스마트폰(주 1회 진단) ↔ TV(매일 운동)  |  기기 발열·인지 피로 동시 해결",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    # Two-column flow diagram
    col_w = (CONTENT_W - Cm(1.2)) / 2
    col_h = Cm(9.2)
    ct = CONTENT_TOP

    # Left: Phone column
    rect(s, MARGIN_L, ct, col_w, col_h, fill=C["soft"], line=C["hairline_s"], line_pt=1)
    tag(s, MARGIN_L + Pt(14), ct + Pt(14), "스마트폰 — 주 1회")
    tb(s, MARGIN_L + Pt(14), ct + Pt(36), col_w - Pt(28), Pt(22),
       "카메라 진단", Pt(16), C["ink"], bold=True)
    steps_phone = [
        "① 카메라로 5~7가지 동작 촬영",
        "② AI가 33개 관절 3D 좌표 추출",
        "③ ROM(관절 가동범위) 측정 & 등급 판정",
        "④ 측정 결과 JSON → 서버 전송",
        "⑤ 영상은 즉시 파기 (프라이버시 보호)",
    ]
    body = "\n".join(steps_phone)
    tb(s, MARGIN_L + Pt(14), ct + Pt(68), col_w - Pt(28), Cm(3.5),
       body, Pt(11), C["charcoal"], line_sp=Pt(19))
    # Advantage chip
    rect(s, MARGIN_L + Pt(14), ct + Cm(6.0), col_w - Pt(28), Pt(20),
         fill=C["tro_soft"])
    tb(s, MARGIN_L + Pt(14), ct + Cm(6.0), col_w - Pt(28), Pt(20),
       "비동기 처리 → 발열 없음  |  보급형 폰 50FPS 달성",
       Pt(9), C["tro"], bold=True, align=PP_ALIGN.CENTER)

    # Arrow
    arrow_l = MARGIN_L + col_w + Cm(0.25)
    arrow_t = ct + Cm(4.0)
    tb(s, arrow_l, arrow_t, Cm(0.7), Cm(1.2), "→", Pt(28), C["primary"], bold=True,
       align=PP_ALIGN.CENTER)

    # Right: TV column
    tv_l = MARGIN_L + col_w + Cm(0.7)
    rect(s, tv_l, ct, col_w, col_h, fill=C["primary_s"], line=C["hairline_s"], line_pt=1)
    tag(s, tv_l + Pt(14), ct + Pt(14), "TV 화면 — 매일")
    tb(s, tv_l + Pt(14), ct + Pt(36), col_w - Pt(28), Pt(22),
       "트로트 운동", Pt(16), C["ink"], bold=True)
    steps_tv = [
        "① Lottie 캐릭터 안무 재생 (TV 대화면)",
        "② 트로트 리듬에 맞춰 따라하기",
        "③ 하단 피토그래픽 동작 가이드 실시간 전환",
        "④ 리모컨 하나로 조작 (스마트폰 내려두기)",
        "⑤ 운동 완료 → 주간 리포트 자동 업데이트",
    ]
    body2 = "\n".join(steps_tv)
    tb(s, tv_l + Pt(14), ct + Pt(68), col_w - Pt(28), Cm(3.5),
       body2, Pt(11), C["charcoal"], line_sp=Pt(19))
    # Advantage chip
    rect(s, tv_l + Pt(14), ct + Cm(6.0), col_w - Pt(28), Pt(20),
         fill=C["primary_s"])
    tb(s, tv_l + Pt(14), ct + Cm(6.0), col_w - Pt(28), Pt(20),
       "스마트폰 내려두고 편하게 운동  |  인지 피로 0",
       Pt(9), C["primary"], bold=True, align=PP_ALIGN.CENTER)

    # Comparison table
    comp_t = ct + col_h + Cm(0.35)
    comp_h = Cm(1.8)
    rect(s, MARGIN_L, comp_t, CONTENT_W, comp_h, fill=C["ink_deep"])
    headers = ["항목", "기존 솔루션", "Tro-Fit"]
    rows    = ["하드웨어", "웨어러블 필수", "스마트폰 1개"]
    col3w   = CONTENT_W / 3
    for ci, (hd, rv) in enumerate(zip(headers, rows)):
        xp = MARGIN_L + ci * col3w + Pt(10)
        tb(s, xp, comp_t + Pt(4),  col3w - Pt(20), Pt(14), hd, Pt(8.5), C["steel"], bold=True)
        tb(s, xp, comp_t + Pt(20), col3w - Pt(20), Pt(16), rv, Pt(10.5), C["on_dark"], bold=(ci==2))

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 04 — Vision AI Pipeline
# ─────────────────────────────────────────────────────────
def slide_04(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Technology — Vision AI",
           slide_num="04 / 09",
           source="출처: Namdari 2012  |  Morrey 1981  |  Rowe 2000  |  Atkinson & Nevill 1998")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "임상 논문 기반 관절 가동범위(ROM) 측정 엔진",
       Pt(30), C["ink"], bold=True, line_sp=Pt(38))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "MediaPipe BlazePose Full — 33개 관절 3D 좌표  |  PC CPU 10.48ms / 95.42 FPS  |  보급형 폰 ~20ms / 50FPS",
       Pt(12), C["charcoal"], line_sp=Pt(19))

    ct = CONTENT_TOP

    # Left: Pipeline flow (50%)
    flow_w = CONTENT_W * 0.50
    steps = [
        ("STEP 1", "카메라 FMS 평가 촬영", "5~7가지 동작 가이드"),
        ("STEP 2", "On-device TFLite 추론", "33개 관절 3D (x, y, z) 추출 · 영상 즉시 파기"),
        ("STEP 3", "3D→2D 폴백 + AMA 변환", "측면 동작 Z축 왜곡 제거 · 각도 오차 −19° 교정"),
        ("STEP 4", "ROM 계산 & 등급 판정", "NORMAL / WARNING / CRITICAL · 임상 오차 마진 ±6°"),
    ]
    step_h = Cm(2.0)
    for i, (tag_t, title, desc) in enumerate(steps):
        sy = ct + i * (step_h + Cm(0.18))
        # step tag
        rect(s, MARGIN_L, sy + Pt(6), Cm(1.4), Pt(14), fill=C["primary"])
        tb(s, MARGIN_L, sy + Pt(6), Cm(1.4), Pt(14),
           tag_t, Pt(7), C["on_dark"], bold=True, align=PP_ALIGN.CENTER)
        # content
        tb(s, MARGIN_L + Cm(1.55), sy, flow_w - Cm(1.7), Pt(22),
           title, Pt(13), C["ink"], bold=True, line_sp=Pt(20))
        tb(s, MARGIN_L + Cm(1.55), sy + Pt(22), flow_w - Cm(1.7), Pt(18),
           desc, Pt(10), C["slate"], line_sp=Pt(16))
        # connector
        if i < len(steps) - 1:
            rect(s, MARGIN_L + Cm(0.55), sy + step_h - Pt(2),
                 Pt(2), Cm(0.28), fill=C["hairline"])

    # Right: Benchmark + ROM table (48%)
    right_l = MARGIN_L + flow_w + Cm(0.6)
    right_w = CONTENT_W - flow_w - Cm(0.6)

    # Benchmark cards
    bench = [
        ("10.48ms", "PC CPU 추론 속도", C["primary"]),
        ("95.42 FPS", "실측 처리 프레임", C["tro"]),
    ]
    bw = (right_w - Cm(0.3)) / 2
    for bi, (bv, bl, bc) in enumerate(bench):
        stat_card(s, right_l + bi*(bw+Cm(0.3)), ct, bw, Cm(1.7), bv, bl, bc)

    # ROM grade table
    tbl_t = ct + Cm(2.0)
    tb(s, right_l, tbl_t, right_w, Pt(18),
       "관절별 ROM 등급 기준 (ADL 논문 근거)", Pt(11), C["ink"], bold=True)
    tbl_t += Pt(22)
    rows = [
        ("어깨 굴곡", "≥120°  NORMAL", "Namdari 2012"),
        ("어깨 외전", "≥120°  NORMAL", "Namdari 2012"),
        ("팔꿈치 굴곡", "≥130°  NORMAL", "Morrey 1981"),
        ("무릎 굴곡", "≥120°  NORMAL", "Rowe 2000"),
        ("좌우 비대칭", "≤9°   NORMAL", "Atkinson 1998"),
    ]
    row_h = Cm(1.05)
    for ri, (joint, grade, ref) in enumerate(rows):
        ry = tbl_t + ri * row_h
        bg = C["soft"] if ri % 2 == 0 else C["canvas"]
        rect(s, right_l, ry, right_w, row_h, fill=bg)
        tb(s, right_l + Pt(8), ry + Pt(4), right_w * 0.35, row_h - Pt(4),
           joint, Pt(9.5), C["ink"], bold=True)
        tb(s, right_l + right_w * 0.35 + Pt(4), ry + Pt(4), right_w * 0.35, row_h - Pt(4),
           grade, Pt(9.5), C["tro"], bold=True)
        tb(s, right_l + right_w * 0.70 + Pt(4), ry + Pt(4), right_w * 0.30, row_h - Pt(4),
           ref, Pt(7.5), C["steel"])

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 05 — RAG + LLM Choreography
# ─────────────────────────────────────────────────────────
def slide_05(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Technology — RAG + LLM",
           slide_num="05 / 09",
           source="기술 스택: ChromaDB + LangChain + sentence-transformers + FastAPI  |  임베딩 36개 문서 완료")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "RAG로 개인 관절 상태에 맞는 트로트 안무를 처방하다",
       Pt(28), C["ink"], bold=True, line_sp=Pt(36))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "ROM 결과 → 임상 DB 검색 → LLM 안무 생성 → TV 출력  |  개인 맞춤형 완전 자동화 파이프라인",
       Pt(12), C["charcoal"], line_sp=Pt(19))

    ct = CONTENT_TOP
    # RAG pipeline boxes (horizontal flow)
    boxes = [
        ("ROM 측정 결과", "무릎 굴곡 115°\n어깨 외전 95°", C["tro_soft"], C["tro"]),
        ("ChromaDB 검색", "임상 ROM 기준 문서\n안전/위험 동작 DB\n노인 운동 처방 가이드", C["primary_s"], C["primary"]),
        ("LLM 프롬프트", "트로트 리듬에 맞춰\n안전한 7가지 동작\n조합 요청", C["soft"], C["ink"]),
        ("안무 JSON 생성", "Lottie 애니메이션 +\n피토그래픽 카드로\nTV 화면 출력", RGBColor(0xE8,0xFF,0xF5), C["tro"]),
    ]
    bw = (CONTENT_W - Cm(0.9)) / 4
    bh = Cm(5.2)
    for bi, (title, body, bg, tc) in enumerate(boxes):
        bl = MARGIN_L + bi * (bw + Cm(0.3))
        rect(s, bl, ct, bw, bh, fill=bg, line=C["hairline_s"], line_pt=1)
        tb(s, bl + Pt(12), ct + Pt(12), bw - Pt(24), Pt(22),
           title, Pt(12), tc, bold=True)
        tb(s, bl + Pt(12), ct + Pt(38), bw - Pt(24), Cm(3.0),
           body, Pt(10.5), C["charcoal"], line_sp=Pt(17))
        if bi < len(boxes) - 1:
            ax = bl + bw + Cm(0.05)
            tb(s, ax, ct + bh/2 - Pt(10), Cm(0.25), Pt(20),
               "→", Pt(14), C["primary"], bold=True)

    # Evidence section
    ev_t = ct + bh + Cm(0.5)
    rect(s, MARGIN_L, ev_t, CONTENT_W, Cm(2.5), fill=C["ink_deep"])

    tb(s, MARGIN_L + Pt(16), ev_t + Pt(10), CONTENT_W * 0.45, Pt(20),
       "RAG 구현 실증 데이터", Pt(11), C["on_dark"], bold=True)

    evidence = [
        "ChromaDB 임베딩 완료: 36개 임상 문서",
        '쿼리: "어깨 외전 정상 기준과 WARNING 기준 알려줘"',
        "응답: NORMAL ≥120° / WARNING 80~119°  (거리값 0.119 — 높은 정확도)",
        "Namdari 2012 논문 기반 임상적으로 정확한 검색 결과 확인",
    ]
    for ei, ev in enumerate(evidence):
        tb(s, MARGIN_L + Pt(16), ev_t + Pt(30) + ei * Pt(14),
           CONTENT_W - Pt(32), Pt(14),
           f"• {ev}", Pt(9), C["on_dark"], line_sp=Pt(14))

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 06 — Team
# ─────────────────────────────────────────────────────────
def slide_06(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Team",
           slide_num="06 / 09")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "실마리 팀 — 4인 4색 전문화 역할 분담",
       Pt(30), C["ink"], bold=True, line_sp=Pt(38))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "PM · Vision AI · Backend · Frontend · Cloud — 각자의 기여를 결과물 기준으로",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    ct = CONTENT_TOP
    members = [
        {
            "name": "박준형",
            "role": "PM & Vision AI",
            "color": C["tro"],
            "tasks": [
                "MediaPipe 파이프라인 설계·구현",
                "ROM 알고리즘 (AMA 변환 · ±6° 마진)",
                "PC CPU 실측 벤치마크 (10.48ms / 95FPS)",
                "프로젝트 총괄 (WBS · 발표)",
            ],
            "output": "snapshot_rom_pipeline.py",
        },
        {
            "name": "이태균",
            "role": "Backend & LLM",
            "color": C["primary"],
            "tasks": [
                "FastAPI 인증 서버 구축",
                "PostgreSQL 스키마 설계",
                "RAG 파이프라인 (LangChain + ChromaDB)",
                "Celery + Redis 비동기 큐",
            ],
            "output": "FastAPI + LLM 서버",
        },
        {
            "name": "한상협",
            "role": "Frontend & Cloud",
            "color": RGBColor(0x7B,0x2F,0xBF),
            "tasks": [
                "React Native 앱 UI/UX",
                "Docker Compose 인프라",
                "Chromecast TV 캐스팅 연동",
                "Prometheus + Grafana 모니터링",
            ],
            "output": "Mobile App + Infra",
        },
        {
            "name": "조영진",
            "role": "AI 설계 & QA",
            "color": C["attention"],
            "tasks": [
                "ChromaDB 36개 문서 임베딩",
                "포즈 추출 모듈 개발",
                "통합 테스트 시나리오 작성",
                "발표 자료 총괄",
            ],
            "output": "RAG 파이프라인 + QA",
        },
    ]

    card_w = (CONTENT_W - Cm(0.9)) / 4
    card_h = Cm(9.0)
    for i, m in enumerate(members):
        cl = MARGIN_L + i * (card_w + Cm(0.3))
        rect(s, cl, ct, card_w, card_h, fill=C["canvas"],
             line=C["hairline_s"], line_pt=1)
        # top color strip
        rect(s, cl, ct, card_w, Pt(4), fill=m["color"])
        # role tag
        rect(s, cl + Pt(12), ct + Pt(14), card_w - Pt(24), Pt(14), fill=C["soft"])
        tb(s, cl + Pt(12), ct + Pt(14), card_w - Pt(24), Pt(14),
           m["role"], Pt(7.5), m["color"], bold=True, align=PP_ALIGN.CENTER)
        # name
        tb(s, cl + Pt(12), ct + Pt(34), card_w - Pt(24), Pt(24),
           m["name"], Pt(16), C["ink"], bold=True)
        # divider
        hairline(s, cl + Pt(12), ct + Pt(60), card_w - Pt(24))
        # tasks
        tasks_txt = "\n".join(f"• {t}" for t in m["tasks"])
        tb(s, cl + Pt(12), ct + Pt(68), card_w - Pt(24), Cm(4.2),
           tasks_txt, Pt(9.5), C["charcoal"], line_sp=Pt(16))
        # output chip
        rect(s, cl + Pt(12), ct + Cm(7.1), card_w - Pt(24), Pt(18),
             fill=m["color"])
        tb(s, cl + Pt(12), ct + Cm(7.1), card_w - Pt(24), Pt(18),
           m["output"], Pt(8), C["on_dark"], bold=True, align=PP_ALIGN.CENTER)

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 07 — Demo Video Placeholder
# ─────────────────────────────────────────────────────────
def slide_07(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Live Demo",
           slide_num="07 / 09")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "지금 바로 보여드립니다",
       Pt(32), C["ink"], bold=True, line_sp=Pt(40))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "관절 측정 → ROM 등급 결과 → 맞춤 안무 처방 — 전체 동작 시연",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    ct = CONTENT_TOP
    # Video placeholder
    vid_w = CONTENT_W * 0.68
    vid_h = Cm(9.0)
    rect(s, MARGIN_L, ct, vid_w, vid_h, fill=C["ink_deep"])
    tb(s, MARGIN_L, ct, vid_w, vid_h,
       "[ 시연 영상 재생 ]\napp_test.mp4",
       Pt(18), C["on_dark"], bold=True, align=PP_ALIGN.CENTER, line_sp=Pt(28))

    # Right: preview checklist
    check_l = MARGIN_L + vid_w + Cm(0.5)
    check_w = CONTENT_W - vid_w - Cm(0.5)
    tb(s, check_l, ct, check_w, Pt(22),
       "시연 포인트", Pt(12), C["ink"], bold=True)
    points = [
        ("①", "스마트폰 카메라로 동작 수행", C["tro"]),
        ("②", "실시간 관절 오버레이\n(33개 랜드마크 시각화)", C["primary"]),
        ("③", "ROM 분석 결과\nNORMAL / WARNING / CRITICAL", C["attention"]),
        ("④", "안무 피토그래픽\n카드 화면 출력", C["tro"]),
    ]
    for pi, (num, txt, col) in enumerate(points):
        py = ct + Pt(30) + pi * Cm(2.1)
        rect(s, check_l, py, Pt(20), Pt(20), fill=col)
        tb(s, check_l, py, Pt(20), Pt(20),
           num, Pt(9), C["on_dark"], bold=True, align=PP_ALIGN.CENTER)
        tb(s, check_l + Pt(26), py, check_w - Pt(30), Cm(1.8),
           txt, Pt(10.5), C["charcoal"], line_sp=Pt(17))

    # Bottom note
    note_t = ct + vid_h + Cm(0.35)
    tb(s, MARGIN_L, note_t, CONTENT_W, Pt(18),
       '영상 재생 전 멘트: "실제 앱으로 관절 측정부터 ROM 등급 결과까지, 전체 동작을 보여드리겠습니다."',
       Pt(9.5), C["slate"], line_sp=Pt(15))

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 08 — Market & Commercialization
# ─────────────────────────────────────────────────────────
def slide_08(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Market & Business Model",
           slide_num="08 / 09",
           source="출처: 2025 통계청  |  2024 건강보험통계연보  |  돌봄통합지원법 2026.3 시행  |  일본 시니어 헬스케어 시장 약 10조 엔")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "초고령사회가 시장이다 — B2G 사업화 전략",
       Pt(30), C["ink"], bold=True, line_sp=Pt(38))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "지자체 협약 → 복지관 구독 대시보드 → 일본 시니어 시장 진출  |  추가 장비 비용 0원",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    ct = CONTENT_TOP
    col_w = (CONTENT_W - Cm(0.6)) / 2

    # Left: Market size stat cards
    mstats = [
        ("20.3%",   "한국 65세↑ 비율 (2025 초고령사회)", C["critical"]),
        ("551만원",  "노인 1인 연평균 진료비",             C["ink"]),
        ("10조 엔",  "일본 시니어 헬스케어 시장 규모",    C["tro"]),
    ]
    sc_h = Cm(2.9)
    for mi, (v, l, vc) in enumerate(mstats):
        stat_card(s, MARGIN_L, ct + mi * (sc_h + Cm(0.2)), col_w, sc_h, v, l, vc)

    # Policy highlight
    pol_t = ct + 3 * (sc_h + Cm(0.2))
    rect(s, MARGIN_L, pol_t, col_w, Cm(1.2), fill=C["tro_soft"],
         line=C["hairline_s"], line_pt=1)
    tb(s, MARGIN_L + Pt(12), pol_t + Pt(8), col_w - Pt(24), Pt(18),
       "2026년 3월 '돌봄통합지원법' 시행", Pt(11), C["tro"], bold=True)
    tb(s, MARGIN_L + Pt(12), pol_t + Pt(24), col_w - Pt(24), Pt(18),
       "예방 플랫폼 정책 수요 급증 → 정부 지원 시장 진입 최적 시점",
       Pt(9.5), C["charcoal"])

    # Right: B2G model
    right_l = MARGIN_L + col_w + Cm(0.6)
    tb(s, right_l, ct, col_w, Pt(22), "수익 모델 (B2G/B2B)", Pt(12), C["ink"], bold=True)

    phases = [
        ("Phase 1", "지자체 협약",
         "행정복지센터 · 노인복지관\n어르신 온보딩 + Chromecast 보급",
         C["tro"]),
        ("Phase 2", "구독 대시보드",
         "복지기관 단위 구독\n다수 어르신 건강 트렌드 모니터링",
         C["primary"]),
        ("Phase 3", "일본 시장 진출",
         "B2G 레퍼런스 기반\n65세↑ 29.4% 시장 공략",
         C["attention"]),
    ]
    phase_h = Cm(3.0)
    for pi, (phase, title, desc, col) in enumerate(phases):
        py = ct + Pt(28) + pi * (phase_h + Cm(0.2))
        rect(s, right_l, py, col_w, phase_h, fill=C["canvas"],
             line=C["hairline_s"], line_pt=1)
        rect(s, right_l, py, Pt(4), phase_h, fill=col)
        tb(s, right_l + Pt(14), py + Pt(8), col_w - Pt(24), Pt(14),
           phase, Pt(8), col, bold=True)
        tb(s, right_l + Pt(14), py + Pt(24), col_w - Pt(24), Pt(20),
           title, Pt(13), C["ink"], bold=True)
        tb(s, right_l + Pt(14), py + Pt(46), col_w - Pt(24), Cm(1.2),
           desc, Pt(9.5), C["charcoal"], line_sp=Pt(16))

    # Differentiation strip
    diff_t = FOOTER_TOP - Cm(1.45)
    rect(s, MARGIN_L, diff_t, CONTENT_W, Cm(1.3), fill=C["soft"])
    diffs = [
        "추가 장비 구매 0원",
        "앱 설치 없는 보호자 카카오톡 리포트",
        "가정 내 예방 → 국가 의료비 절감 기여",
    ]
    for di, d in enumerate(diffs):
        dw = CONTENT_W / 3
        tb(s, MARGIN_L + di * dw + Pt(16), diff_t + Pt(10), dw - Pt(32), Pt(20),
           f"✔  {d}", Pt(10), C["tro"], bold=True)

    return s


# ─────────────────────────────────────────────────────────
#  SLIDE 09 — Roadmap + Closing
# ─────────────────────────────────────────────────────────
def slide_09(prs):
    s = blank(prs)
    set_bg(s, C["canvas"])
    chrome(s,
           chapter="Roadmap & Closing",
           slide_num="09 / 09")

    tb(s, MARGIN_L, TITLE_TOP, CONTENT_W, Cm(1.3),
       "우리는 멈추지 않습니다 — Phase 2 로드맵",
       Pt(30), C["ink"], bold=True, line_sp=Pt(38))

    tb(s, MARGIN_L, SUBTITLE_TOP, CONTENT_W, Cm(0.75),
       "MVP 완료 → TV 앱 · 필드 테스트 · 특허 출원 → 지자체 파트너십 · 일본 진출",
       Pt(12.5), C["charcoal"], line_sp=Pt(19))

    ct = CONTENT_TOP
    phases = [
        {
            "phase": "현재 완료 — Phase 1 MVP",
            "color": C["tro"],
            "items": [
                "MediaPipe ROM 측정 엔진 (10.48ms / 95FPS)",
                "ChromaDB RAG 파이프라인 (36개 문서 임베딩)",
                "FastAPI 백엔드 + JWT 인증",
                "React Native UI 흐름",
            ],
            "icon": "DONE",
        },
        {
            "phase": "Phase 2 — 2026년 7~8월",
            "color": C["primary"],
            "items": [
                "TV 앱 개발 (WebOS / Tizen)",
                "Adaptive Compute (배터리 최적화)",
                "65세↑ 현장 필드 테스트 (10명)",
                "보호자 카카오톡 리포트 완성",
            ],
            "icon": "IN PROGRESS",
        },
        {
            "phase": "Phase 3 — 2026년 9월 이후",
            "color": C["attention"],
            "items": [
                "특허 출원",
                "지자체 파트너십 체결",
                "일본 시니어 요양 시장 탐색",
                "Federated Learning 도입 (분산 학습)",
            ],
            "icon": "PLANNED",
        },
    ]

    phase_w = (CONTENT_W - Cm(0.8)) / 3
    phase_h = Cm(8.0)
    for pi, ph in enumerate(phases):
        pl = MARGIN_L + pi * (phase_w + Cm(0.4))
        rect(s, pl, ct, phase_w, phase_h, fill=C["canvas"],
             line=C["hairline_s"], line_pt=1)
        rect(s, pl, ct, phase_w, Pt(5), fill=ph["color"])
        # icon badge
        rect(s, pl + Pt(12), ct + Pt(14), Cm(2.5), Pt(16), fill=ph["color"])
        tb(s, pl + Pt(12), ct + Pt(14), Cm(2.5), Pt(16),
           ph["icon"], Pt(8), C["on_dark"], bold=True, align=PP_ALIGN.CENTER)
        # title
        tb(s, pl + Pt(12), ct + Pt(36), phase_w - Pt(24), Pt(38),
           ph["phase"], Pt(11.5), C["ink"], bold=True, line_sp=Pt(18))
        hairline(s, pl + Pt(12), ct + Pt(76), phase_w - Pt(24))
        # items
        items_txt = "\n".join(f"• {it}" for it in ph["items"])
        tb(s, pl + Pt(12), ct + Pt(84), phase_w - Pt(24), Cm(4.5),
           items_txt, Pt(10.5), C["charcoal"], line_sp=Pt(18))

    # Closing message strip
    close_t = ct + phase_h + Cm(0.45)
    rect(s, MARGIN_L, close_t, CONTENT_W, Cm(2.0), fill=C["ink_deep"])
    tb(s, MARGIN_L + Pt(20), close_t + Pt(10), CONTENT_W - Pt(40), Cm(1.5),
       '"Tro-Fit은 트로트 한 곡과 스마트폰 하나로,\n'
       '어르신이 건강하고 오래 살 수 있는 세상을 만듭니다."',
       Pt(14), C["on_dark"], bold=True, line_sp=Pt(22))

    return s


# ─────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    prs = new_prs()

    slide_01(prs)
    slide_02(prs)
    slide_03(prs)
    slide_04(prs)
    slide_05(prs)
    slide_06(prs)
    slide_07(prs)
    slide_08(prs)
    slide_09(prs)

    out = os.path.join(os.path.dirname(__file__), "tro_fit_full.pptx")
    prs.save(out)
    print(f"[OK]  Saved  9 slides  →  {out}")
