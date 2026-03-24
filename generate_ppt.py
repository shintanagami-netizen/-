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


# ===== カラー設定（ここを自社カラーに変更してください） =====
COLOR_PRIMARY   = RGBColor(0x00, 0x47, 0xAB)  # 濃いブルー（タイトル背景など）
COLOR_ACCENT    = RGBColor(0x00, 0x8B, 0xD8)  # 明るいブルー（アクセント）
COLOR_BG        = RGBColor(0xF5, 0xF8, 0xFF)  # 薄いブルー白（スライド背景）
COLOR_TEXT      = RGBColor(0x1A, 0x1A, 0x2E)  # ほぼ黒（本文テキスト）
COLOR_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_BULLET    = RGBColor(0x00, 0x47, 0xAB)  # 箇条書き記号の色

# スライドサイズ（16:9）
SLIDE_WIDTH  = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# ロゴサイズ・位置（右上）
LOGO_W = Inches(1.5)
LOGO_H = Inches(0.6)
LOGO_X = SLIDE_WIDTH - LOGO_W - Inches(0.2)
LOGO_Y = Inches(0.1)


def set_font(run, size, bold=False, color=None, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color


def add_rect(slide, x, y, w, h, color):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        x, y, w, h
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_logo(slide, logo_path):
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, LOGO_X, LOGO_Y, LOGO_W, LOGO_H)


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


def make_title_slide(prs, title, subtitle, logo_path):
    """タイトルスライド"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # 背景
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_BG)

    # 左側カラーバー
    add_rect(slide, 0, 0, Inches(0.35), SLIDE_HEIGHT, COLOR_PRIMARY)

    # タイトルエリア背景
    add_rect(slide, Inches(0.35), Inches(1.8), SLIDE_WIDTH - Inches(0.35), Inches(2.5), COLOR_PRIMARY)

    # タイトルテキスト
    tb = slide.shapes.add_textbox(Inches(0.7), Inches(1.9), SLIDE_WIDTH - Inches(1.2), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_font(run, 28, bold=True, color=COLOR_WHITE)

    # サブタイトル
    if subtitle:
        add_textbox(slide, subtitle,
                    Inches(0.7), Inches(4.5), SLIDE_WIDTH - Inches(1.2), Inches(0.8),
                    16, color=COLOR_TEXT)

    # 下部アクセントライン
    add_rect(slide, Inches(0.35), SLIDE_HEIGHT - Inches(0.15),
             SLIDE_WIDTH - Inches(0.35), Inches(0.15), COLOR_ACCENT)

    add_logo(slide, logo_path)
    return slide


def make_section_slide(prs, title, logo_path):
    """セクション区切りスライド"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_PRIMARY)
    add_rect(slide, 0, 0, Inches(0.35), SLIDE_HEIGHT, COLOR_ACCENT)

    tb = slide.shapes.add_textbox(Inches(0.7), Inches(2.8), SLIDE_WIDTH - Inches(1.2), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_font(run, 32, bold=True, color=COLOR_WHITE)

    add_logo(slide, logo_path)
    return slide


def make_content_slide(prs, title, items, logo_path):
    """コンテンツスライド（箇条書き）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 背景
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_BG)

    # 左カラーバー
    add_rect(slide, 0, 0, Inches(0.35), SLIDE_HEIGHT, COLOR_PRIMARY)

    # タイトル帯
    add_rect(slide, Inches(0.35), 0, SLIDE_WIDTH - Inches(0.35), Inches(1.1), COLOR_PRIMARY)

    # タイトルテキスト
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.15), SLIDE_WIDTH - Inches(1.0), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_font(run, 22, bold=True, color=COLOR_WHITE)

    # コンテンツエリア
    content_x = Inches(0.6)
    content_y = Inches(1.25)
    content_w = SLIDE_WIDTH - Inches(0.9)
    content_h = SLIDE_HEIGHT - Inches(1.5)

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

        p.space_before = Pt(4)
        p.space_after = Pt(2)

        if item_type == 'numbered':
            prefix = f"{num}. "
            p.level = 0
        elif item_type == 'bullet':
            prefix = "・"
            p.level = 0
        elif item_type == 'sub_bullet':
            prefix = "  − "
            p.level = 1
        else:
            prefix = ""

        run = p.add_run()
        run.text = prefix + text
        size = 14 if item_type == 'sub_bullet' else 16
        set_font(run, size, color=COLOR_TEXT)

    # 下部ライン
    add_rect(slide, Inches(0.35), SLIDE_HEIGHT - Inches(0.12),
             SLIDE_WIDTH - Inches(0.35), Inches(0.12), COLOR_ACCENT)

    add_logo(slide, logo_path)
    return slide


# ==================== Markdownパーサー ====================

def parse_markdown(md_text):
    """
    Markdownを解析してスライド構造のリストを返す。
    各要素: {'type': 'title'|'section'|'content', 'title': str, 'subtitle': str, 'items': [...]}
    """
    slides = []
    current = None

    for raw_line in md_text.splitlines():
        line = raw_line.strip()

        # H1 → タイトルスライド
        if line.startswith('# '):
            if current:
                slides.append(current)
            current = {'type': 'title', 'title': line[2:].strip(), 'subtitle': '', 'items': []}

        # H2 → セクションスライド（コンテンツがなければ区切り、あればコンテンツ）
        elif line.startswith('## '):
            if current:
                slides.append(current)
            current = {'type': 'content', 'title': line[3:].strip(), 'subtitle': '', 'items': []}

        # 箇条書き（* または -）
        elif re.match(r'^[\*\-] ', line):
            text = line[2:].strip()
            if current:
                current['items'].append({'type': 'bullet', 'text': text})

        # 字下げ箇条書き
        elif re.match(r'^\s{2,}[\*\-] ', raw_line):
            text = re.sub(r'^\s+[\*\-] ', '', raw_line).strip()
            if current:
                current['items'].append({'type': 'sub_bullet', 'text': text})

        # 番号付きリスト
        elif re.match(r'^\d+\. ', line):
            text = re.sub(r'^\d+\. ', '', line).strip()
            num = len([i for i in (current['items'] if current else []) if i['type'] == 'numbered']) + 1
            if current:
                current['items'].append({'type': 'numbered', 'text': text, 'num': num})

        # メタ情報行（* 作成：xxx や * 対象：xxx）→ タイトルスライドのサブタイトル化
        elif line.startswith('* ') and current and current['type'] == 'title':
            current['subtitle'] += line[2:].strip() + '  '

        # 空でないその他のテキスト → 概要テキストとして追加
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

    for s in slides_data:
        if s['type'] == 'title':
            make_title_slide(prs, s['title'], s.get('subtitle', '').strip(), logo_path)
        elif s['type'] == 'section':
            make_section_slide(prs, s['title'], logo_path)
        else:  # content
            items = s.get('items', [])
            if items:
                make_content_slide(prs, s['title'], items, logo_path)
            else:
                make_section_slide(prs, s['title'], logo_path)

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
