"""
SALESCORE 提案資料 - コンサルティング品質デザイン
Abeam / Accenture スタイル参考
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import math, os

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# カラーパレット（洗練されたネイビー系）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NAVY_DEEP   = RGBColor(0x0C, 0x1A, 0x4B)   # 最も深いネイビー（表紙パネル）
NAVY        = RGBColor(0x1E, 0x3A, 0x8A)   # メインネイビー
BLUE        = RGBColor(0x25, 0x63, 0xEB)   # アクセントブルー
BLUE_MID    = RGBColor(0x3B, 0x82, 0xF6)   # ミディアムブルー
BLUE_LIGHT  = RGBColor(0xDB, 0xEA, 0xFE)   # 薄いブルー背景
BLUE_XLIGHT = RGBColor(0xEF, 0xF4, 0xFF)   # 極薄ブルー
TEXT        = RGBColor(0x0F, 0x17, 0x2A)   # ほぼ黒（本文）
TEXT_MID    = RGBColor(0x47, 0x55, 0x69)   # ミドルグレー（サブテキスト）
TEXT_LIGHT  = RGBColor(0x94, 0xA3, 0xB8)   # 薄グレー（フッター）
LINE        = RGBColor(0xCB, 0xD5, 0xE1)   # 罫線
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# レイアウト定数
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW = Inches(13.33)   # スライド幅
SH = Inches(7.5)     # スライド高さ

# 内側スライド共通マージン
ML  = Inches(0.65)   # 左マージン
MR  = Inches(0.65)   # 右マージン
CW  = Inches(12.03)  # コンテンツ幅 (13.33 - 0.65*2 - logo_space)
HDR_Y   = Inches(0.20)   # タイトルY
RULE_Y  = Inches(0.82)   # タイトル下ライン
BODY_Y  = Inches(1.05)   # ボディ開始Y
BODY_H  = Inches(5.75)   # ボディ高さ
FTR_Y   = Inches(6.92)   # フッターライン
FTR_TY  = Inches(6.97)   # フッターテキストY

# ロゴ
LOGO_PATH  = "/home/user/-/salescore_logo.png"
LOGO_PATH_W = "/home/user/-/salescore_logo_white.png"
LOGO_W = Inches(1.90)
LOGO_H = Inches(0.52)
LOGO_X = Inches(13.33 - 0.55 - 1.90)
LOGO_Y = Inches(0.18)

# 日本語フォント
JP_FONT    = "IPAGothic"
LATIN_FONT = "Calibri"

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ヘルパー関数
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH


def add_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])


def rect(slide, x, y, w, h, fill=None, line=None, line_w=Pt(0.75)):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape


def txt(slide, text, x, y, w, h,
        size=Pt(12), bold=False, color=TEXT,
        align=PP_ALIGN.LEFT, wrap=True,
        font=JP_FONT, italic=False,
        line_spacing=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    if line_spacing:
        tf.paragraphs[0].line_spacing = line_spacing
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    try:
        run.font.name = font
    except Exception:
        pass
    return tb


def txt_in_shape(shape, text, size=Pt(12), bold=False,
                 color=WHITE, align=PP_ALIGN.CENTER,
                 font=JP_FONT, v_anchor="middle"):
    tf = shape.text_frame
    tf.word_wrap = True
    if v_anchor == "middle":
        from pptx.enum.text import MSO_ANCHOR
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    try:
        run.font.name = font
    except Exception:
        pass


def logo(slide, x=LOGO_X, y=LOGO_Y, w=LOGO_W, h=LOGO_H, white=False):
    path = LOGO_PATH_W if white else LOGO_PATH
    if os.path.exists(path):
        slide.shapes.add_picture(path, x, y, w, h)


def add_inner_slide_chrome(slide, title_text, page_num):
    """内側スライド共通のヘッダー・ロゴ・フッターを追加"""
    # ──── ロゴ ────
    logo(slide)

    # ──── タイトル ────
    txt(slide, title_text,
        ML, HDR_Y, Inches(9.8), Inches(0.60),
        size=Pt(20), bold=True, color=TEXT, font=JP_FONT)

    # ──── タイトル下アクセントライン ────
    rect(slide, ML, RULE_Y, Inches(12.03), Pt(2.5), fill=NAVY)

    # ──── フッターライン ────
    rect(slide, ML, FTR_Y, Inches(12.03), Pt(1.0), fill=LINE)

    # ──── フッターテキスト ────
    txt(slide, "Confidential  All Rights Reserved SALESCORE Inc.",
        ML, FTR_TY, Inches(9.5), Inches(0.35),
        size=Pt(8), color=TEXT_LIGHT, font=LATIN_FONT)

    # ──── ページ番号 ────
    txt(slide, str(page_num),
        Inches(12.5), FTR_TY, Inches(0.5), Inches(0.35),
        size=Pt(9), color=TEXT_LIGHT, align=PP_ALIGN.RIGHT, font=LATIN_FONT)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 1 : 表紙 ─ 左ネイビーパネル + 右ホワイト
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
s1 = add_slide()

PANEL_W = Inches(4.60)

# 左ネイビーパネル
rect(s1, 0, 0, PANEL_W, SH, fill=NAVY_DEEP)

# 左パネル: ロゴ（白版）
logo(s1, x=Inches(0.45), y=Inches(0.42), w=Inches(1.80), h=Inches(0.50), white=True)

# 左パネル: 区切りライン
rect(s1, Inches(0.45), Inches(1.10), Inches(3.70), Pt(1.5), fill=BLUE)

# 左パネル: 会社名
txt(s1, "SALESCORE株式会社",
    Inches(0.45), Inches(1.22), Inches(3.8), Inches(0.5),
    size=Pt(12), color=WHITE, font=JP_FONT)

# 左パネル: 日付
txt(s1, "2026年3月",
    Inches(0.45), Inches(6.50), Inches(3.0), Inches(0.4),
    size=Pt(11), color=TEXT_LIGHT, font=JP_FONT)

# 左パネル: フッター
txt(s1, "Confidential",
    Inches(0.45), Inches(7.0), Inches(3.0), Inches(0.35),
    size=Pt(8), color=TEXT_LIGHT, font=LATIN_FONT)

# 右コンテンツエリア（白）は自動的に白
# 右エリアの縦アクセントライン（ネイビーとホワイトの境界に細いブルーライン）
rect(s1, PANEL_W, 0, Inches(0.06), SH, fill=BLUE)

# 宛先ラベル
RX = PANEL_W + Inches(0.90)
txt(s1, "株式会社〇〇  御中",
    RX, Inches(1.50), Inches(7.5), Inches(0.55),
    size=Pt(16), color=TEXT_MID, font=JP_FONT)

# デコレーションライン（タイトル上）
rect(s1, RX, Inches(2.18), Inches(7.5), Pt(3.0), fill=NAVY)

# メインタイトル
txt(s1, "営業生産性を2倍にする\n3つのアクション",
    RX, Inches(2.32), Inches(7.5), Inches(2.60),
    size=Pt(38), bold=True, color=TEXT, font=JP_FONT)

# タイトル下ライン
rect(s1, RX, Inches(5.12), Inches(7.5), Pt(1.5), fill=LINE)

# サブライン
txt(s1, "データ活用 × 行動変容 × 組織定着",
    RX, Inches(5.22), Inches(7.5), Inches(0.5),
    size=Pt(13), color=TEXT_MID, bold=False, italic=True, font=JP_FONT)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 2 : 課題スライド
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
s2 = add_slide()
add_inner_slide_chrome(
    s2,
    "売上の8割を2割の「できる営業」が担う構造が、組織スケールを阻んでいる",
    2
)

# セクションキャプション
txt(s2, "売上の78%をトップ20%の営業が創出。属人化が組織の成長を制約している。",
    ML, Inches(1.05), Inches(10.5), Inches(0.40),
    size=Pt(11), color=TEXT_MID, font=JP_FONT)

# ─ 左カラム ─────────────────────────────────────────
LCX = ML
LCW = Inches(5.60)
LCY = Inches(1.55)

# 左カラムヘッダー
hdr_l = rect(s2, LCX, LCY, LCW, Inches(0.42), fill=NAVY)
txt_in_shape(hdr_l, "属人化が生む 3 つの弊害",
             size=Pt(11), bold=True, color=WHITE, align=PP_ALIGN.LEFT, font=JP_FONT)
s2.shapes[-1].text_frame.paragraphs[0].runs[0].font.size = Pt(11)
# テキストの左マージン調整はインデントで対応
tx = s2.shapes[-1].text_frame.paragraphs[0]
tx.runs[0].font.size = Pt(11)
# 左詰めのためにtextboxを重ねる
txt(s2, "属人化が生む 3 つの弊害",
    LCX + Inches(0.20), LCY + Inches(0.06), LCW - Inches(0.3), Inches(0.32),
    size=Pt(11), bold=True, color=WHITE, font=JP_FONT)

issues = [
    ("01", "育成が属人的でノウハウが継承されない",
     "トップ営業の知識・経験が組織内に蓄積されず、\n離職や異動のたびにノウハウがリセットされる。"),
    ("02", "マネジャーが「感覚」でマネジメントする",
     "行動データが可視化されておらず、指示が\n経験値・直感に依存するため再現性がない。"),
    ("03", "組織が拡大しても売上が比例して伸びない",
     "人材追加コストが増加する一方、\n勝ちパターンが展開できず生産性が停滞する。"),
]

for i, (num, heading, detail) in enumerate(issues):
    y = LCY + Inches(0.52) + i * Inches(1.65)
    box = rect(s2, LCX, y, LCW, Inches(1.52),
               fill=BLUE_XLIGHT, line=BLUE_LIGHT, line_w=Pt(1.2))
    # バッジ
    badge = rect(s2, LCX + Inches(0.18), y + Inches(0.18),
                 Inches(0.46), Inches(0.46), fill=NAVY)
    txt_in_shape(badge, num, size=Pt(10), bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, font=LATIN_FONT)
    # 見出し
    txt(s2, heading,
        LCX + Inches(0.78), y + Inches(0.14), LCW - Inches(0.90), Inches(0.38),
        size=Pt(12), bold=True, color=TEXT, font=JP_FONT)
    # 詳細
    txt(s2, detail,
        LCX + Inches(0.78), y + Inches(0.52), LCW - Inches(0.90), Inches(0.92),
        size=Pt(9.5), color=TEXT_MID, font=JP_FONT)

# ─ 右カラム ──────────────────────────────────────────
RCX = ML + LCW + Inches(0.50)
RCW = Inches(5.88)
RCY = Inches(1.55)

# 右カラムヘッダー
hdr_r = rect(s2, RCX, RCY, RCW, Inches(0.42), fill=NAVY)
txt(s2, "売上構成比の実態",
    RCX + Inches(0.20), RCY + Inches(0.06), RCW - Inches(0.3), Inches(0.32),
    size=Pt(11), bold=True, color=WHITE, font=JP_FONT)

# 棒グラフエリア
BY = RCY + Inches(0.62)
BAR_MAX_W = RCW - Inches(0.30)

bars = [
    ("トップ 20% の営業", 0.78, "78%", NAVY, Inches(0.38)),
    ("残り 80% の営業",   0.22, "22%", BLUE, Inches(1.65)),
]
for label, ratio, pct_text, color, bar_y in bars:
    txt(s2, label, RCX, BY + bar_y - Inches(0.28), RCW, Inches(0.28),
        size=Pt(10), color=TEXT_MID, font=JP_FONT)
    bar_w = BAR_MAX_W * ratio
    bar = rect(s2, RCX, BY + bar_y, bar_w, Inches(0.55), fill=color)
    txt_in_shape(bar, pct_text, size=Pt(14), bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, font=LATIN_FONT)

# So-what ボックス
sw_y = RCY + Inches(2.80)
sw_box = rect(s2, RCX, sw_y, RCW, Inches(1.85),
              fill=BLUE_XLIGHT, line=BLUE, line_w=Pt(2.0))
txt(s2, "→ 結論",
    RCX + Inches(0.28), sw_y + Inches(0.18), RCW - Inches(0.40), Inches(0.32),
    size=Pt(10), bold=True, color=BLUE, font=JP_FONT)
txt(s2, "個人の「才能」に依存した営業から脱却し、\nデータと仕組みで誰もが成果を出せる組織へ転換することが急務。",
    RCX + Inches(0.28), sw_y + Inches(0.50), RCW - Inches(0.40), Inches(1.15),
    size=Pt(12), bold=True, color=NAVY, font=JP_FONT)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SLIDE 3 : 解決策スライド
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
s3 = add_slide()
add_inner_slide_chrome(
    s3,
    "営業生産性を2倍にする「3つのアクション」",
    3
)

# サブキャプション
txt(s3, "データで行動を可視化し、再現可能な仕組みに変換することで、組織全体の底上げが実現できる",
    ML, Inches(1.05), Inches(11.0), Inches(0.40),
    size=Pt(11), color=TEXT_MID, font=JP_FONT)

# タイムラインバー
TL_Y = Inches(1.58)
TL_H = Pt(2.0)
rect(s3, ML, TL_Y, Inches(12.03), TL_H, fill=LINE)

# タイムラインテキスト
txt(s3, "◀  導入から約 1〜2 ヶ月で成果検証フェーズへ  ▶",
    ML, TL_Y - Inches(0.30), Inches(12.03), Inches(0.28),
    size=Pt(9.5), color=BLUE, align=PP_ALIGN.CENTER, font=JP_FONT)

# 3ステップのカラム
COL_W  = Inches(3.80)
COL_GAP = Inches(0.21)
COL_Y  = Inches(1.65)

steps = [
    ("ACTION  01", "Key アクション特定",
     "・お打ち合わせ段階で\n　 録画視聴・インタビューを実施\n・トップ営業の行動パターン分析\n・勝ちパターンの仮説設計",
     "短期で成果出しの\n検証まで実行できる"),
    ("ACTION  02", "仕組み実装",
     "・Key アクションを KPI に実装\n・会議体・ダッシュボード設計\n・現場への運用説明会の実施\n・管理職トレーニング",
     "成功体験を経て\n要領がわかる"),
    ("ACTION  03", "習慣化・定着",
     "・日々の会議体で進捗確認\n・ブロッカーをタイムリー排除\n・アジャイルに磨き込み\n・効果測定と横展開",
     "戦略を実行する習慣が\n形成され成果に HIT する"),
]

for i, (action_label, step_name, bullets, benefit) in enumerate(steps):
    cx = ML + i * (COL_W + COL_GAP)

    # ── ステップヘッダーボックス（ネイビー） ──
    hdr_h = Inches(1.18)
    hdr = rect(s3, cx, COL_Y, COL_W, hdr_h, fill=NAVY)

    # ACTION番号（小さめ）
    txt(s3, action_label,
        cx + Inches(0.18), COL_Y + Inches(0.12), COL_W - Inches(0.20), Inches(0.28),
        size=Pt(9), bold=True, color=BLUE_MID, align=PP_ALIGN.CENTER, font=LATIN_FONT)

    # ステップ名
    txt(s3, step_name,
        cx + Inches(0.12), COL_Y + Inches(0.42), COL_W - Inches(0.24), Inches(0.68),
        size=Pt(16), bold=True, color=WHITE, align=PP_ALIGN.CENTER, font=JP_FONT)

    # ── 区切り矢印（最後以外） ──
    if i < 2:
        arr_x = cx + COL_W + Inches(0.04)
        txt(s3, "▶",
            arr_x, COL_Y + Inches(0.40),
            COL_GAP + Inches(0.12), Inches(0.40),
            size=Pt(11), color=NAVY, align=PP_ALIGN.CENTER, font=LATIN_FONT)

    # ── 施策リスト（ヘッダー下） ──
    bullet_y = COL_Y + hdr_h + Inches(0.14)
    bullet_h = Inches(1.72)
    bullet_box = rect(s3, cx, bullet_y, COL_W, bullet_h,
                      fill=BLUE_XLIGHT, line=BLUE_LIGHT, line_w=Pt(0.8))
    txt(s3, bullets,
        cx + Inches(0.18), bullet_y + Inches(0.12),
        COL_W - Inches(0.28), bullet_h - Inches(0.20),
        size=Pt(9.5), color=TEXT, font=JP_FONT)

    # ── 利点バッジ ──
    badge_y = bullet_y + bullet_h + Inches(0.30)
    badge_h = Inches(0.32)
    badge_w = Inches(1.20)
    badge = rect(s3, cx + (COL_W - badge_w) / 2, badge_y,
                 badge_w, badge_h, fill=NAVY)
    txt_in_shape(badge, f"利点  0{i+1}", size=Pt(8), bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER, font=JP_FONT)

    # ── 利点説明ボックス ──
    ben_y = badge_y + badge_h - Pt(1)
    ben_h = Inches(1.08)
    ben_box = rect(s3, cx, ben_y, COL_W, ben_h,
                   fill=WHITE, line=BLUE, line_w=Pt(1.8))
    txt(s3, benefit,
        cx + Inches(0.14), ben_y + Inches(0.14),
        COL_W - Inches(0.28), ben_h - Inches(0.22),
        size=Pt(12), bold=True, color=NAVY,
        align=PP_ALIGN.CENTER, font=JP_FONT)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 保存
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT = "/home/user/-/SALESCORE_営業生産性向上_提案資料.pptx"
prs.save(OUTPUT)
print(f"✅ 保存完了: {OUTPUT}")
