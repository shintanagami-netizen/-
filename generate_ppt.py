"""
Markdown → PowerPoint (.pptx) 変換スクリプト
使い方:
  python generate_ppt.py input.md output.pptx [logo.png]

Markdown記法:
  # タイトル                          → タイトルスライド
  サブタイトル or メタ情報テキスト     → タイトルのサブタイトルに追加
  ### セクション名                     → セクション区切りスライド（ダークネイビー全面）
  ## 1.1. 節番号 | スライドタイトル   → コンテンツスライド（節番号付き）
  ## スライドタイトル                  → コンテンツスライド（節番号なし）
  > リード文                           → タイトル直下の概要メッセージ（太字・強調スタイル）
  - 箇条書き                           → 標準箇条書き（●）
    - サブ項目                         → 字下げ箇条書き（–）
  1. 番号付きリスト                    → 番号付きリスト
  ---                                  → コンテンツ内の視覚的区切り
"""

import sys
import re
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree


# ===== SALESCORE カラーパレット =====
COLOR_DARK_NAVY  = RGBColor(0x1D, 0x34, 0x61)   # 濃いネイビー（タイトルスライド背景・セクション帯）
COLOR_NAVY       = RGBColor(0x1B, 0x4F, 0xA8)   # ミディアムネイビー（番号・記号アクセント）
COLOR_BLUE       = RGBColor(0x1E, 0x6F, 0xD8)   # ブランドブルー（セクション区切りアクセント）
COLOR_TEXT       = RGBColor(0x1A, 0x1A, 0x1A)   # ほぼ黒（本文テキスト）
COLOR_BLACK      = RGBColor(0x00, 0x00, 0x00)   # 黒（ラインなど）
COLOR_SECTION_NUM = RGBColor(0x55, 0x65, 0x7A)  # 節番号グレー
COLOR_GRAY       = RGBColor(0x80, 0x80, 0x80)   # グレー（フッター・サブ記号）
COLOR_GRAY_LIGHT = RGBColor(0xCC, 0xCC, 0xCC)   # 薄グレー（区切り線）
COLOR_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# スライドサイズ（16:9）
SLIDE_WIDTH  = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# ロゴ（右上・縦横比維持）
LOGO_W = Inches(2.2)
LOGO_X = SLIDE_WIDTH - LOGO_W - Inches(0.15)
LOGO_Y = Inches(0.08)

# 左アクセントバー
LEFT_BAR_W = Inches(0.08)

# コンテンツ左余白（バーの外側）
CONTENT_X  = Inches(0.50)
CONTENT_RW = SLIDE_WIDTH - Inches(0.50) - Inches(0.35)   # 右余白

# タイトルエリア
TITLE_AREA_X       = Inches(0.50)
TITLE_AREA_MAX_W   = SLIDE_WIDTH - Inches(2.95)   # ロゴとの重複を避ける

# ヘッダーライン Y（タイトル下）
HEADER_LINE_Y  = Inches(0.83)
HEADER_LINE_H  = Pt(1.8)

# フッターライン Y
FOOTER_LINE_Y  = SLIDE_HEIGHT - Inches(0.42)
FOOTER_LINE_H  = Pt(1.2)

FOOTER_TEXT = "Confidential All Rights Reserved SALESCORE Inc."


# ─────────────────────────────────────────────
#  ユーティリティ
# ─────────────────────────────────────────────

def set_font(run, size, bold=False, color=None, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color


def remove_shadow(shape):
    """シェイプ影エフェクトを除去（テーマ継承を上書き）"""
    spPr = shape._element.spPr
    for el in spPr.findall(qn('a:effectLst')):
        spPr.remove(el)
    etree.SubElement(spPr, qn('a:effectLst'))


def add_rect(slide, x, y, w, h, color, line_color=None, line_w=Pt(0)):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_w if line_w else Pt(0.75)
    else:
        shape.line.fill.background()
    remove_shadow(shape)
    return shape


def add_logo(slide, logo_path):
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, LOGO_X, LOGO_Y, width=LOGO_W)


def add_footer(slide, page_num=None):
    """フッターライン + テキスト + ページ番号"""
    add_rect(slide, 0, FOOTER_LINE_Y, SLIDE_WIDTH, FOOTER_LINE_H, COLOR_BLACK)

    # フッターテキスト（左）
    tb = slide.shapes.add_textbox(
        CONTENT_X, FOOTER_LINE_Y + Pt(3), Inches(8), Inches(0.35)
    )
    tf = tb.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = FOOTER_TEXT
    set_font(r, 8, color=COLOR_GRAY)

    # ページ番号（右）
    if page_num is not None:
        tb2 = slide.shapes.add_textbox(
            SLIDE_WIDTH - Inches(0.6), FOOTER_LINE_Y + Pt(3),
            Inches(0.45), Inches(0.35)
        )
        tf2 = tb2.text_frame
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.RIGHT
        r2 = p2.add_run()
        r2.text = str(page_num)
        set_font(r2, 9, color=COLOR_GRAY)


# ─────────────────────────────────────────────
#  スライド生成
# ─────────────────────────────────────────────

def make_title_slide(prs, title, subtitle, logo_path, page_num=1):
    """
    タイトルスライド
    ・上部 58%: ダークネイビー（ロゴ・タイトル）
    ・下部 42%: 白（サブタイトル・メタ情報）
    ・境界: ネイビーブルーのアクセントライン
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    split_y = SLIDE_HEIGHT * 0.58

    # 上部ネイビー背景
    add_rect(slide, 0, 0, SLIDE_WIDTH, split_y, COLOR_DARK_NAVY)
    # 下部白背景
    add_rect(slide, 0, split_y, SLIDE_WIDTH, SLIDE_HEIGHT - split_y, COLOR_WHITE)
    # 境界アクセントライン
    add_rect(slide, 0, split_y - Pt(3), SLIDE_WIDTH, Pt(6), COLOR_NAVY)

    # タイトルテキスト（白、左寄り・上部中央）
    title_y = split_y * 0.28
    tb = slide.shapes.add_textbox(
        Inches(1.0), title_y, SLIDE_WIDTH - Inches(3.8), split_y * 0.55
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.space_after = Pt(4)
    r = p.add_run()
    r.text = title
    set_font(r, 34, bold=True, color=COLOR_WHITE)

    # サブタイトル（白、タイトル直下）
    if subtitle:
        sub_y = split_y * 0.28 + split_y * 0.55
        tb2 = slide.shapes.add_textbox(
            Inches(1.0), sub_y, SLIDE_WIDTH - Inches(3.8), Inches(0.9)
        )
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = subtitle.strip()
        set_font(r2, 13, color=RGBColor(0xCC, 0xD8, 0xF0))

    # 下部エリアに日付などの補足（任意）
    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


def make_section_slide(prs, title, section_num, logo_path, page_num=None):
    """
    セクション区切りスライド
    ・全面ダークネイビー
    ・セクション番号（小）＋ タイトル（大・白）
    ・ブランドブルーのアクセントライン
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_DARK_NAVY)

    center_y = SLIDE_HEIGHT * 0.42

    if section_num:
        # セクション番号ラベル
        tb_num = slide.shapes.add_textbox(
            Inches(1.2), center_y - Inches(0.55),
            SLIDE_WIDTH - Inches(2.4), Inches(0.35)
        )
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        r_num = p_num.add_run()
        r_num.text = section_num
        set_font(r_num, 11, bold=True, color=COLOR_BLUE)

    # セクションタイトル
    tb = slide.shapes.add_textbox(
        Inches(1.2), center_y - Inches(0.2) if section_num else center_y - Inches(0.4),
        SLIDE_WIDTH - Inches(2.4), Inches(1.2)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    set_font(r, 36, bold=True, color=COLOR_WHITE)

    # アクセントライン（タイトル下）
    line_y = center_y + Inches(0.85) if not section_num else center_y + Inches(1.05)
    add_rect(slide, Inches(1.2), line_y, Inches(2.5), Pt(3), COLOR_BLUE)

    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


def make_content_slide(prs, title, section_num, items, logo_path, page_num=None):
    """
    コンテンツスライド（コンサルファームスタイル）
    ・純白背景
    ・左細ネイビーアクセントバー
    ・節番号（小グレー）＋ メインタイトル（大ネイビー）
    ・ダークネイビーヘッダーライン
    ・リード文（> で指定）
    ・箇条書き・番号付き・サブ箇条書き・区切り線
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_WHITE)

    # ── タイトルエリア ──
    if section_num:
        # 節番号（小・グレー）
        tb_sec = slide.shapes.add_textbox(
            TITLE_AREA_X, Inches(0.07),
            TITLE_AREA_MAX_W, Inches(0.24)
        )
        tf_sec = tb_sec.text_frame
        p_sec = tf_sec.paragraphs[0]
        r_sec = p_sec.add_run()
        r_sec.text = section_num
        set_font(r_sec, 9, bold=True, color=COLOR_SECTION_NUM)
        main_title_y = Inches(0.27)
    else:
        main_title_y = Inches(0.10)

    # メインタイトル（大・ダークネイビー・太字）
    tb_title = slide.shapes.add_textbox(
        TITLE_AREA_X, main_title_y, TITLE_AREA_MAX_W, Inches(0.62)
    )
    tf_title = tb_title.text_frame
    tf_title.word_wrap = False
    p_title = tf_title.paragraphs[0]
    r_title = p_title.add_run()
    r_title.text = title
    set_font(r_title, 26, bold=True, color=COLOR_TEXT)

    # ── ヘッダーライン（ブラック・全幅）──
    add_rect(slide, 0, HEADER_LINE_Y, SLIDE_WIDTH, HEADER_LINE_H, COLOR_BLACK)

    # ── コンテンツエリア ──
    lead_items    = [i for i in items if i.get('type') == 'lead']
    content_items = [i for i in items if i.get('type') != 'lead']

    current_y = HEADER_LINE_Y + Inches(0.18)

    # リード文（概要メッセージ）
    if lead_items:
        lead_text = '　'.join(i['text'] for i in lead_items)
        tb_lead = slide.shapes.add_textbox(
            CONTENT_X, current_y, CONTENT_RW, Inches(0.60)
        )
        tf_lead = tb_lead.text_frame
        tf_lead.word_wrap = True
        p_lead = tf_lead.paragraphs[0]
        r_lead = p_lead.add_run()
        r_lead.text = lead_text
        set_font(r_lead, 12, color=COLOR_TEXT)
        current_y += Inches(0.65)

    # 箇条書き・テキストブロック
    if content_items:
        # dividerで分割し、各ブロックをテキストボックスとして描画
        blocks = []
        current_block = []
        for item in content_items:
            if item.get('type') == 'divider':
                if current_block:
                    blocks.append(current_block)
                    current_block = []
                blocks.append(None)   # 区切り線マーカー
            else:
                current_block.append(item)
        if current_block:
            blocks.append(current_block)

        available_h = FOOTER_LINE_Y - current_y - Inches(0.1)

        # 区切り線の数を数えて高さを按分
        n_dividers = blocks.count(None)
        n_blocks   = len([b for b in blocks if b is not None])
        divider_h  = Pt(8)
        total_divider_h = n_dividers * divider_h
        block_h = (available_h - total_divider_h) / max(n_blocks, 1) if n_blocks else available_h

        block_y = current_y
        for block in blocks:
            if block is None:
                # 区切り線（薄グレー）
                div_center_y = block_y + divider_h / 2 - Pt(0.75)
                add_rect(slide, CONTENT_X, div_center_y,
                         CONTENT_RW, Pt(1.0), COLOR_GRAY_LIGHT)
                block_y += divider_h
            else:
                tb = slide.shapes.add_textbox(CONTENT_X, block_y, CONTENT_RW, block_h)
                tf = tb.text_frame
                tf.word_wrap = True

                first = True
                for item in block:
                    if first:
                        p = tf.paragraphs[0]
                        first = False
                    else:
                        p = tf.add_paragraph()

                    itype = item.get('type', 'bullet')
                    text  = item.get('text', '')
                    num   = item.get('num', '')

                    if itype == 'bullet':
                        p.space_before = Pt(6)
                        p.space_after  = Pt(2)
                        r_sym = p.add_run()
                        r_sym.text = "• "
                        set_font(r_sym, 12, color=COLOR_NAVY)
                        r_txt = p.add_run()
                        r_txt.text = text
                        set_font(r_txt, 13, color=COLOR_TEXT)

                    elif itype == 'sub_bullet':
                        p.space_before = Pt(3)
                        p.space_after  = Pt(1)
                        r_sym = p.add_run()
                        r_sym.text = "      – "
                        set_font(r_sym, 10, color=COLOR_GRAY)
                        r_txt = p.add_run()
                        r_txt.text = text
                        set_font(r_txt, 11, color=COLOR_TEXT)

                    elif itype == 'numbered':
                        p.space_before = Pt(6)
                        p.space_after  = Pt(2)
                        r_sym = p.add_run()
                        r_sym.text = f"{num}. "
                        set_font(r_sym, 13, bold=True, color=COLOR_NAVY)
                        r_txt = p.add_run()
                        r_txt.text = text
                        set_font(r_txt, 13, color=COLOR_TEXT)

                    else:   # plain text
                        p.space_before = Pt(4)
                        p.space_after  = Pt(2)
                        r_txt = p.add_run()
                        r_txt.text = text
                        set_font(r_txt, 13, color=COLOR_TEXT)

                block_y += block_h

    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


# ─────────────────────────────────────────────
#  Markdown パーサー
# ─────────────────────────────────────────────

def parse_markdown(md_text):
    """
    スライド構造のリストを返す。
    各要素:
      title   : {type, title, subtitle}
      section : {type, title, section_num}
      content : {type, title, section_num, items}
    """
    slides   = []
    current  = None

    for raw_line in md_text.splitlines():
        line = raw_line.strip()

        # ── H1 → タイトルスライド ──
        if line.startswith('# '):
            if current:
                slides.append(current)
            current = {
                'type': 'title',
                'title': line[2:].strip(),
                'subtitle': '',
                'items': [],
            }

        # ── H3 → セクション区切りスライド ──
        elif line.startswith('### '):
            if current:
                slides.append(current)
            raw_title = line[4:].strip()
            # "番号 | タイトル" 形式
            if ' | ' in raw_title:
                sec, ttl = raw_title.split(' | ', 1)
                current = {'type': 'section', 'title': ttl.strip(),
                           'section_num': sec.strip(), 'items': []}
            else:
                current = {'type': 'section', 'title': raw_title,
                           'section_num': None, 'items': []}

        # ── H2 → コンテンツスライド ──
        elif line.startswith('## '):
            if current:
                slides.append(current)
            raw_title = line[3:].strip()
            # "節番号 | タイトル" 形式
            if ' | ' in raw_title:
                sec, ttl = raw_title.split(' | ', 1)
                current = {'type': 'content', 'title': ttl.strip(),
                           'section_num': sec.strip(), 'items': []}
            else:
                current = {'type': 'content', 'title': raw_title,
                           'section_num': None, 'items': []}

        # ── リード文（> blockquote）──
        elif line.startswith('> '):
            text = line[2:].strip()
            if current and current.get('type') in ('content',):
                current['items'].append({'type': 'lead', 'text': text})

        # ── 区切り線 ---  ──
        elif line == '---':
            if current and current.get('type') == 'content':
                current['items'].append({'type': 'divider'})

        # ── 標準箇条書き（* または -）──
        elif re.match(r'^[\*\-] ', line):
            text = line[2:].strip()
            if current:
                current['items'].append({'type': 'bullet', 'text': text})

        # ── 字下げ箇条書き ──
        elif re.match(r'^\s{2,}[\*\-] ', raw_line):
            text = re.sub(r'^\s+[\*\-] ', '', raw_line).strip()
            if current:
                current['items'].append({'type': 'sub_bullet', 'text': text})

        # ── 番号付きリスト ──
        elif re.match(r'^\d+\. ', line):
            text = re.sub(r'^\d+\. ', '', line).strip()
            num = len([i for i in (current.get('items', [])) if i.get('type') == 'numbered']) + 1
            if current:
                current['items'].append({'type': 'numbered', 'text': text, 'num': num})

        # ── タイトルスライドのサブタイトル行 ──
        elif line and current and current['type'] == 'title':
            current['subtitle'] += line + ' '

        # ── その他テキスト（コンテンツスライドの平文）──
        elif line and current and current.get('type') == 'content':
            current['items'].append({'type': 'text', 'text': line})

    if current:
        slides.append(current)

    return slides


# ─────────────────────────────────────────────
#  メイン
# ─────────────────────────────────────────────

def generate_ppt(md_path, out_path, logo_path=None):
    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    slides_data = parse_markdown(md_text)

    prs = Presentation()
    prs.slide_width  = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    for i, s in enumerate(slides_data, start=1):
        stype = s['type']

        if stype == 'title':
            make_title_slide(
                prs, s['title'], s.get('subtitle', '').strip(),
                logo_path, page_num=i
            )
        elif stype == 'section':
            make_section_slide(
                prs, s['title'], s.get('section_num'),
                logo_path, page_num=i
            )
        else:  # content
            items = s.get('items', [])
            if items or s.get('title'):
                make_content_slide(
                    prs, s['title'], s.get('section_num'),
                    items, logo_path, page_num=i
                )
            else:
                make_section_slide(
                    prs, s['title'], s.get('section_num'),
                    logo_path, page_num=i
                )

    prs.save(out_path)
    print(f"✅ 保存完了: {out_path}")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("使い方: python generate_ppt.py input.md output.pptx [logo.png]")
        sys.exit(1)

    md_file   = sys.argv[1]
    out_file  = sys.argv[2]
    logo_file = sys.argv[3] if len(sys.argv) >= 4 else None

    generate_ppt(md_file, out_file, logo_file)
