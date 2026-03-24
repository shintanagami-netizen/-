"""
Markdown → PowerPoint (.pptx) 変換スクリプト
使い方:
  python generate_ppt.py input.md output.pptx [logo.png]
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
COLOR_DARK_NAVY  = RGBColor(0x1D, 0x34, 0x61)  # 濃いネイビー（タイトルスライド背景・セクション帯）
COLOR_NAVY       = RGBColor(0x1B, 0x4F, 0xA8)  # ミディアムネイビー（アクセント要素）
COLOR_BLUE       = RGBColor(0x1E, 0x6F, 0xD8)  # ブランドブルー（箇条書き記号・強調）
COLOR_TEXT       = RGBColor(0x1A, 0x1A, 0x1A)  # ほぼ黒（本文テキスト）
COLOR_GRAY       = RGBColor(0x80, 0x80, 0x80)  # グレー（フッターテキスト）
COLOR_BLACK      = RGBColor(0x00, 0x00, 0x00)  # 黒（ラインなど）
COLOR_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# スライドサイズ（16:9）
SLIDE_WIDTH  = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# ロゴサイズ・位置（右上）幅だけ指定、高さは縦横比を自動維持
LOGO_W = Inches(2.2)
LOGO_X = SLIDE_WIDTH - LOGO_W - Inches(0.15)
LOGO_Y = Inches(0.08)

# ヘッダーラインのY位置（タイトルの下）
HEADER_LINE_Y = Inches(0.72)
# フッターラインのY位置
FOOTER_LINE_Y = SLIDE_HEIGHT - Inches(0.42)
# ラインの太さ
LINE_H = Pt(1.2)

FOOTER_TEXT = "Confidential All Rights Reserved SALESCORE Inc."


def set_font(run, size, bold=False, color=None, italic=False, font_name=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    if font_name:
        run.font.name = font_name


def add_rect(slide, x, y, w, h, color, line_color=None):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def add_logo(slide, logo_path):
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, LOGO_X, LOGO_Y, width=LOGO_W)


def add_header_line(slide):
    """タイトル下の横線（全幅）"""
    add_rect(slide, Inches(0), HEADER_LINE_Y, SLIDE_WIDTH, LINE_H, COLOR_BLACK)


def add_footer(slide, page_num=None):
    """フッターライン＋テキスト＋ページ番号"""
    # フッターライン
    add_rect(slide, Inches(0), FOOTER_LINE_Y, SLIDE_WIDTH, LINE_H, COLOR_BLACK)

    # フッターテキスト（左）
    tb = slide.shapes.add_textbox(
        Inches(0.3), FOOTER_LINE_Y + Pt(3),
        Inches(8), Inches(0.35)
    )
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = FOOTER_TEXT
    set_font(run, 8, color=COLOR_GRAY)

    # ページ番号（右）
    if page_num is not None:
        tb2 = slide.shapes.add_textbox(
            SLIDE_WIDTH - Inches(0.6), FOOTER_LINE_Y + Pt(3),
            Inches(0.45), Inches(0.35)
        )
        tf2 = tb2.text_frame
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.RIGHT
        run2 = p2.add_run()
        run2.text = str(page_num)
        set_font(run2, 9, color=COLOR_GRAY)


def add_textbox(slide, text, x, y, w, h, size, bold=False, color=None,
                align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_font(run, size, bold=bold, color=color or COLOR_TEXT)
    return tb


def make_title_slide(prs, title, subtitle, logo_path, page_num=1):
    """タイトルスライド：ダークネイビー上半分＋白下半分"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 上部ダークネイビー背景
    split_y = Inches(4.5)
    add_rect(slide, 0, 0, SLIDE_WIDTH, split_y, COLOR_DARK_NAVY)
    # 下部白背景
    add_rect(slide, 0, split_y, SLIDE_WIDTH, SLIDE_HEIGHT - split_y, COLOR_WHITE)

    # タイトルテキスト（白、上部エリア中央）
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.4), SLIDE_WIDTH - Inches(3.5), Inches(2.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_font(run, 32, bold=True, color=COLOR_WHITE)

    # サブタイトル（白、上部エリア下）
    if subtitle:
        tb2 = slide.shapes.add_textbox(Inches(1.0), Inches(3.7), SLIDE_WIDTH - Inches(3.5), Inches(0.7))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        run2 = p2.add_run()
        run2.text = subtitle
        set_font(run2, 14, color=COLOR_WHITE)

    # ブランドブルーのアクセントバー（split_y付近）
    add_rect(slide, 0, split_y - Inches(0.06), SLIDE_WIDTH, Inches(0.06), COLOR_NAVY)

    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


def make_section_slide(prs, title, logo_path, page_num=None):
    """セクション区切りスライド：ダークネイビー全面"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_DARK_NAVY)

    # 中央テキスト
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.8), SLIDE_WIDTH - Inches(2.0), Inches(1.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_font(run, 32, bold=True, color=COLOR_WHITE)

    # ブルーのアクセントライン（タイトル下）
    add_rect(slide, Inches(1.0), Inches(3.9), Inches(2.0), Pt(3), COLOR_NAVY)

    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


def make_content_slide(prs, title, items, logo_path, page_num=None):
    """コンテンツスライド：SALESCORE実際の資料スタイル"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 純白背景（デフォルトは白なので不要だが明示）
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_WHITE)

    # ページタイトル（黒太字、左上）
    tb = slide.shapes.add_textbox(
        Inches(0.35), Inches(0.1),
        SLIDE_WIDTH - Inches(2.9), Inches(0.6)
    )
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_font(run, 22, bold=True, color=COLOR_TEXT)

    # ヘッダーライン
    add_header_line(slide)

    # コンテンツエリア
    content_x = Inches(0.5)
    content_y = HEADER_LINE_Y + Inches(0.2)
    content_w = SLIDE_WIDTH - Inches(0.8)
    content_h = FOOTER_LINE_Y - content_y - Inches(0.1)

    tb2 = slide.shapes.add_textbox(content_x, content_y, content_w, content_h)
    tf2 = tb2.text_frame
    tf2.word_wrap = True

    first = True
    for item in items:
        if first:
            p = tf2.paragraphs[0]
            first = False
        else:
            p = tf2.add_paragraph()

        item_type = item.get('type', 'bullet')
        text = item.get('text', '')
        num = item.get('num', '')

        p.space_before = Pt(5)
        p.space_after = Pt(2)

        if item_type == 'numbered':
            run_prefix = p.add_run()
            run_prefix.text = f"{num}. "
            set_font(run_prefix, 14, bold=True, color=COLOR_NAVY)
            run_text = p.add_run()
            run_text.text = text
            set_font(run_text, 14, color=COLOR_TEXT)

        elif item_type == 'bullet':
            run_prefix = p.add_run()
            run_prefix.text = "● "
            set_font(run_prefix, 10, color=COLOR_NAVY)
            run_text = p.add_run()
            run_text.text = text
            set_font(run_text, 14, color=COLOR_TEXT)

        elif item_type == 'sub_bullet':
            run_prefix = p.add_run()
            run_prefix.text = "    - "
            set_font(run_prefix, 11, color=COLOR_GRAY)
            run_text = p.add_run()
            run_text.text = text
            set_font(run_text, 12, color=COLOR_TEXT)

        else:
            # プレーンテキスト（概要文など）
            run_text = p.add_run()
            run_text.text = text
            set_font(run_text, 14, color=COLOR_TEXT)

    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


# ==================== Markdownパーサー ====================

def parse_markdown(md_text):
    slides = []
    current = None

    for raw_line in md_text.splitlines():
        line = raw_line.strip()

        if line.startswith('# '):
            if current:
                slides.append(current)
            current = {'type': 'title', 'title': line[2:].strip(), 'subtitle': '', 'items': []}

        elif line.startswith('## '):
            if current:
                slides.append(current)
            current = {'type': 'content', 'title': line[3:].strip(), 'subtitle': '', 'items': []}

        elif re.match(r'^[\*\-] ', line):
            text = line[2:].strip()
            if current:
                current['items'].append({'type': 'bullet', 'text': text})

        elif re.match(r'^\s{2,}[\*\-] ', raw_line):
            text = re.sub(r'^\s+[\*\-] ', '', raw_line).strip()
            if current:
                current['items'].append({'type': 'sub_bullet', 'text': text})

        elif re.match(r'^\d+\. ', line):
            text = re.sub(r'^\d+\. ', '', line).strip()
            num = len([i for i in (current['items'] if current else []) if i['type'] == 'numbered']) + 1
            if current:
                current['items'].append({'type': 'numbered', 'text': text, 'num': num})

        elif line.startswith('* ') and current and current['type'] == 'title':
            current['subtitle'] += line[2:].strip() + '  '

        elif line and current:
            if current['type'] == 'title':
                current['subtitle'] += line + ' '
            else:
                current['items'].append({'type': 'text', 'text': line})

    if current:
        slides.append(current)

    return slides


# ==================== メイン ====================

def generate_ppt(md_path, out_path, logo_path=None):
    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    slides_data = parse_markdown(md_text)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    for i, s in enumerate(slides_data, start=1):
        if s['type'] == 'title':
            make_title_slide(prs, s['title'], s.get('subtitle', '').strip(), logo_path, page_num=i)
        elif s['type'] == 'section':
            make_section_slide(prs, s['title'], logo_path, page_num=i)
        else:  # content
            items = s.get('items', [])
            if items:
                make_content_slide(prs, s['title'], items, logo_path, page_num=i)
            else:
                make_section_slide(prs, s['title'], logo_path, page_num=i)

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
