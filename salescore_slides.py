from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Colors
NAVY = RGBColor(0x1E, 0x3A, 0x6E)
ROYAL_BLUE = RGBColor(0x27, 0x5B, 0xC4)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xF8)
CHARCOAL = RGBColor(0x37, 0x41, 0x51)
RED = RGBColor(0xEF, 0x44, 0x44)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
LIGHT_GRAY = RGBColor(0xF3, 0xF4, 0xF6)
MID_GRAY = RGBColor(0x9C, 0xA3, 0xAF)

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width = W
prs.slide_height = H

def add_slide(prs):
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)

def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_width=Pt(1)):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        x, y, w, h
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, text, x, y, w, h, font_size=Pt(12), bold=False,
                color=BLACK, align=PP_ALIGN.LEFT, wrap=True, font_name="Noto Sans JP"):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.color.rgb = color
    try:
        run.font.name = font_name
    except:
        pass
    return txBox

def add_text_in_shape(shape, text, font_size=Pt(12), bold=False, color=WHITE, align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.color.rgb = color

LOGO_PATH = "/home/user/-/salescore_logo.png"

def add_logo(slide, x, y, w, h):
    """SALESCOREロゴ画像をスライドに追加する"""
    slide.shapes.add_picture(LOGO_PATH, x, y, w, h)

def add_header_footer(slide, title, page_num):
    # Header line
    line = add_rect(slide, Inches(0.4), Inches(0.75), Inches(12.5), Pt(1.5), fill_color=BLACK)

    # Title
    add_textbox(slide, title, Inches(0.4), Inches(0.15), Inches(8), Inches(0.6),
                font_size=Pt(18), bold=True, color=BLACK)

    # Logo image (right)
    add_logo(slide, Inches(10.8), Inches(0.08), Inches(2.1), Inches(0.62))

    # Footer line
    add_rect(slide, Inches(0.4), Inches(7.0), Inches(12.5), Pt(1.5), fill_color=BLACK)

    # Footer text
    add_textbox(slide, "Confidential All Rights Reserved SALESCORE Inc.",
                Inches(0.4), Inches(7.05), Inches(10), Inches(0.35),
                font_size=Pt(8), color=MID_GRAY)

    # Page number
    add_textbox(slide, str(page_num), Inches(12.5), Inches(7.05), Inches(0.4), Inches(0.35),
                font_size=Pt(10), color=BLACK, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────
# SLIDE 1: 表紙
# ─────────────────────────────────────────
slide1 = add_slide(prs)

# Background white (default)
# Header line only
add_rect(slide1, Inches(0.4), Inches(0.75), Inches(12.5), Pt(1.5), fill_color=BLACK)
# Footer line
add_rect(slide1, Inches(0.4), Inches(7.0), Inches(12.5), Pt(1.5), fill_color=BLACK)

# Logo top right
add_logo(slide1, Inches(10.8), Inches(0.08), Inches(2.1), Inches(0.62))

# 宛先
add_textbox(slide1, "株式会社〇〇  御中", Inches(1.2), Inches(1.5), Inches(8), Inches(0.7),
            font_size=Pt(18), bold=False, color=BLACK)

# Main title
add_textbox(slide1, "営業生産性を2倍にする\n3つのアクション",
            Inches(1.2), Inches(2.5), Inches(10), Inches(2.0),
            font_size=Pt(40), bold=True, color=BLACK)

# Company name
add_textbox(slide1, "SALESCORE株式会社", Inches(1.2), Inches(5.2), Inches(6), Inches(0.6),
            font_size=Pt(16), bold=False, color=BLACK)

# Footer text
add_textbox(slide1, "Confidential All Rights Reserved SALESCORE Inc.",
            Inches(0.4), Inches(7.05), Inches(10), Inches(0.35),
            font_size=Pt(8), color=MID_GRAY)
add_textbox(slide1, "1", Inches(12.5), Inches(7.05), Inches(0.4), Inches(0.35),
            font_size=Pt(10), color=BLACK, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────
# SLIDE 2: 課題スライド
# ─────────────────────────────────────────
slide2 = add_slide(prs)
add_header_footer(slide2, "売上の8割を2割の「できる営業」が担う構造が、組織スケールを阻んでいる", 2)

# Section header band
band = add_rect(slide2, Inches(0.4), Inches(0.9), Inches(12.5), Inches(0.45), fill_color=CHARCOAL)
add_text_in_shape(band, "日本の営業組織が抱える構造的課題", font_size=Pt(12), bold=True, color=WHITE, align=PP_ALIGN.LEFT)
band.text_frame.paragraphs[0].runs[0].font.size = Pt(12)

# Subtitle
add_textbox(slide2, "売上の78%をトップ20%の営業が創出。属人化が組織の成長を制約している。",
            Inches(0.4), Inches(1.45), Inches(12.5), Inches(0.4),
            font_size=Pt(12), color=BLACK)

# LEFT column: 課題の構造
left_header = add_rect(slide2, Inches(0.4), Inches(2.0), Inches(5.8), Inches(0.4), fill_color=NAVY)
add_text_in_shape(left_header, "属人化が生む3つの弊害", font_size=Pt(12), bold=True)

# 3 boxes for弊害
issues = [
    ("❶", "育成が属人的で\nノウハウが継承されない"),
    ("❷", "マネジャーが「感覚」で\nマネジメントする"),
    ("❸", "組織が拡大しても\n売上が比例して伸びない"),
]
for i, (badge, text) in enumerate(issues):
    y = Inches(2.5 + i * 1.3)
    box = add_rect(slide2, Inches(0.4), y, Inches(5.8), Inches(1.1),
                   fill_color=LIGHT_BLUE, line_color=ROYAL_BLUE, line_width=Pt(1.5))
    # Badge circle (simulate with rect)
    badge_shape = add_rect(slide2, Inches(0.5), y + Inches(0.2), Inches(0.45), Inches(0.45), fill_color=NAVY)
    add_text_in_shape(badge_shape, badge, font_size=Pt(10), bold=True)
    add_textbox(slide2, text, Inches(1.1), y + Inches(0.1), Inches(5.0), Inches(0.9),
                font_size=Pt(12), color=BLACK)

# RIGHT column: データ
right_header = add_rect(slide2, Inches(6.7), Inches(2.0), Inches(6.1), Inches(0.4), fill_color=NAVY)
add_text_in_shape(right_header, "売上構成比の実態", font_size=Pt(12), bold=True)

# Bar chart simulation
add_textbox(slide2, "トップ20%の営業", Inches(6.7), Inches(2.55), Inches(3.0), Inches(0.35),
            font_size=Pt(11), color=BLACK)
bar1 = add_rect(slide2, Inches(6.7), Inches(2.9), Inches(4.9), Inches(0.55), fill_color=NAVY)
add_text_in_shape(bar1, "78%", font_size=Pt(14), bold=True)

add_textbox(slide2, "残り80%の営業", Inches(6.7), Inches(3.6), Inches(3.0), Inches(0.35),
            font_size=Pt(11), color=BLACK)
bar2 = add_rect(slide2, Inches(6.7), Inches(3.95), Inches(1.35), Inches(0.55), fill_color=ROYAL_BLUE)
add_text_in_shape(bar2, "22%", font_size=Pt(14), bold=True)

# So What box
so_what = add_rect(slide2, Inches(6.7), Inches(5.0), Inches(6.1), Inches(1.1),
                   fill_color=LIGHT_BLUE, line_color=ROYAL_BLUE, line_width=Pt(2))
add_textbox(slide2, "→ 個人の「才能」ではなく、\n　 仕組みで再現することが急務",
            Inches(6.9), Inches(5.05), Inches(5.7), Inches(1.0),
            font_size=Pt(13), bold=True, color=NAVY)


# ─────────────────────────────────────────
# SLIDE 3: 解決策スライド
# ─────────────────────────────────────────
slide3 = add_slide(prs)
add_header_footer(slide3, "営業生産性を2倍にする「3つのアクション」", 3)

# Subtitle
add_textbox(slide3, "データで行動を可視化し、再現可能な仕組みに変換することで、組織全体の底上げが実現できる",
            Inches(0.4), Inches(0.9), Inches(12.5), Inches(0.45),
            font_size=Pt(12), color=BLACK)

# Timeline label
add_textbox(slide3, "◀━━━━━━ 約1〜2ヶ月 ━━━━━━▶",
            Inches(0.4), Inches(1.45), Inches(12.5), Inches(0.35),
            font_size=Pt(11), color=ROYAL_BLUE, align=PP_ALIGN.CENTER)

# Chevron steps (simulate with rectangles + arrows)
steps = [
    ("ACTION\n01", "Keyアクション\n特定"),
    ("ACTION\n02", "仕組み\n実装"),
    ("ACTION\n03", "習慣化\n定着"),
]
chevron_w = Inches(3.8)
chevron_h = Inches(1.1)
for i, (badge, label) in enumerate(steps):
    x = Inches(0.4 + i * 4.3)
    y = Inches(1.9)
    shape = add_rect(slide3, x, y, chevron_w, chevron_h, fill_color=NAVY)
    add_textbox(slide3, f"{badge}\n{label}", x + Inches(0.1), y + Inches(0.05),
                chevron_w - Inches(0.2), chevron_h - Inches(0.1),
                font_size=Pt(13), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Arrow between
    if i < 2:
        add_textbox(slide3, "▶", Inches(4.1 + i * 4.3), y + Inches(0.3),
                    Inches(0.3), Inches(0.5), font_size=Pt(16), color=NAVY, align=PP_ALIGN.CENTER)

# Bullet points under each step
bullets = [
    "・お打ち合わせ段階で\n　録画視聴やインタビューを実施\n・勝ちパターンの仮説設計",
    "・KeyアクションをKPIに実装\n・会議体・ダッシュボード設計\n・現場への運用説明会",
    "・日々の会議体で進捗確認\n・ブロッカーをタイムリー排除\n・アジャイルに磨き込み",
]
for i, bullet in enumerate(bullets):
    x = Inches(0.4 + i * 4.3)
    add_textbox(slide3, bullet, x, Inches(3.1), Inches(3.8), Inches(1.2),
                font_size=Pt(10), color=BLACK)

# 利点 boxes
riyoten = [
    ("利点\n01", "短期で成果出しの\n検証まで実行できる"),
    ("利点\n02", "成功体験を経て\n要領がわかる"),
    ("利点\n03", "戦略を実行する習慣が\n形成され成果にHITする"),
]
for i, (badge_text, desc) in enumerate(riyoten):
    x = Inches(0.4 + i * 4.3)
    y = Inches(4.5)
    box = add_rect(slide3, x, y, Inches(3.8), Inches(1.6),
                   fill_color=WHITE, line_color=ROYAL_BLUE, line_width=Pt(2))

    # Badge
    badge_shape = add_rect(slide3, x + Inches(1.5), y - Inches(0.25), Inches(0.8), Inches(0.5), fill_color=NAVY)
    add_text_in_shape(badge_shape, badge_text, font_size=Pt(8), bold=True)

    add_textbox(slide3, desc, x + Inches(0.1), y + Inches(0.3), Inches(3.6), Inches(1.1),
                font_size=Pt(12), bold=True, color=NAVY, align=PP_ALIGN.CENTER)


# Save
output_path = "/home/user/-/SALESCORE_営業生産性向上_提案資料.pptx"
prs.save(output_path)
print(f"✅ 保存完了: {output_path}")
