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
  > リード文                           → タイトル直下の概要メッセージ
  - 箇条書き                           → 標準箇条書き（●）
    - サブ項目                         → 字下げ箇条書き（–）
  1. 番号付きリスト                    → 番号付きリスト
  ---                                  → コンテンツ内の視覚的区切り
  | ヘッダー1 | ヘッダー2 |           → テーブル（Markdown標準記法）
  |---|---|                            → テーブルヘッダー区切り行
  ::left [タイトル]                   → 2カラムレイアウト（左カラム開始）
  ::right [タイトル]                  → 2カラムレイアウト（右カラム開始）
  :::                                 → カラムレイアウト終了
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
COLOR_STRIPE     = RGBColor(0xF5, 0xF7, 0xFA)   # テーブル縞模様
COLOR_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# スライドサイズ（16:9）
SLIDE_WIDTH  = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# ロゴ（右上・縦横比維持）
LOGO_W = Inches(2.2)
LOGO_X = SLIDE_WIDTH - LOGO_W - Inches(0.15)
LOGO_Y = Inches(0.08)

# コンテンツ左余白
CONTENT_X  = Inches(0.50)
CONTENT_RW = SLIDE_WIDTH - Inches(0.50) - Inches(0.35)

# タイトルエリア
TITLE_AREA_X     = Inches(0.50)
TITLE_AREA_MAX_W = SLIDE_WIDTH - Inches(2.95)

# ヘッダーライン Y（タイトル下）
HEADER_LINE_Y = Inches(0.83)
HEADER_LINE_H = Pt(1.8)

# フッターライン Y
FOOTER_LINE_Y = SLIDE_HEIGHT - Inches(0.42)
FOOTER_LINE_H = Pt(1.2)

FOOTER_TEXT = "Confidential All Rights Reserved SALESCORE Inc."

# テーブル行高さ
TABLE_HEADER_ROW_H = Inches(0.40)
TABLE_DATA_ROW_H   = Inches(0.35)


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


def set_cell_bg(cell, color):
    """テーブルセルの背景色をXMLで直接設定"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for tag in (qn('a:solidFill'), qn('a:noFill'), qn('a:gradFill'), qn('a:pattFill')):
        for el in tcPr.findall(tag):
            tcPr.remove(el)
    sf = etree.SubElement(tcPr, qn('a:solidFill'))
    clr = etree.SubElement(sf, qn('a:srgbClr'))
    clr.set('val', str(color))


def add_logo(slide, logo_path):
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, LOGO_X, LOGO_Y, width=LOGO_W)


def add_footer(slide, page_num=None):
    """フッターライン + テキスト + ページ番号"""
    add_rect(slide, 0, FOOTER_LINE_Y, SLIDE_WIDTH, FOOTER_LINE_H, COLOR_BLACK)

    tb = slide.shapes.add_textbox(
        CONTENT_X, FOOTER_LINE_Y + Pt(3), Inches(8), Inches(0.35)
    )
    tf = tb.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = FOOTER_TEXT
    set_font(r, 8, color=COLOR_GRAY)

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
#  コンテンツレンダラー（共通ヘルパー）
# ─────────────────────────────────────────────

def render_text_items(slide, items, x, y, w, h):
    """
    bullet / sub_bullet / numbered / text / divider の混在リストを描画。
    divider で分割し、各テキストブロックに高さを等分配する。
    """
    blocks = []
    current_block = []
    for item in items:
        if item.get('type') == 'divider':
            if current_block:
                blocks.append(current_block)
                current_block = []
            blocks.append(None)
        else:
            current_block.append(item)
    if current_block:
        blocks.append(current_block)

    DIVIDER_H = Pt(12)
    n_dividers = blocks.count(None)
    n_text_blocks = len([b for b in blocks if b is not None])
    block_h = (h - n_dividers * DIVIDER_H) / max(n_text_blocks, 1) if n_text_blocks else h

    block_y = y
    for block in blocks:
        if block is None:
            add_rect(slide, x, block_y + Pt(4), w, Pt(1.0), COLOR_GRAY_LIGHT)
            block_y += DIVIDER_H
        else:
            tb = slide.shapes.add_textbox(x, block_y, w, block_h)
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

                else:  # plain text
                    p.space_before = Pt(4)
                    p.space_after  = Pt(2)
                    r_txt = p.add_run()
                    r_txt.text = text
                    set_font(r_txt, 13, color=COLOR_TEXT)

            block_y += block_h


def render_table(slide, item, x, y, w):
    """
    テーブルアイテムをpptxテーブルとして描画。
    ヘッダー行: ダークネイビー背景・白太字・中央揃え
    データ行: ストライプ（薄グレー / 白）
    戻り値: テーブルの高さ
    """
    headers = item.get('headers', [])
    rows    = item.get('rows', [])
    n_cols  = len(headers)
    if n_cols == 0:
        return Inches(0)

    n_rows    = len(rows) + 1
    table_h   = TABLE_HEADER_ROW_H + TABLE_DATA_ROW_H * len(rows)

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, table_h)
    tbl = tbl_shape.table

    # 列幅を均等に
    col_w = w // n_cols
    for j in range(n_cols):
        tbl.columns[j].width = col_w

    # ヘッダー行
    tbl.rows[0].height = TABLE_HEADER_ROW_H
    for j, hdr in enumerate(headers):
        cell = tbl.cell(0, j)
        set_cell_bg(cell, COLOR_DARK_NAVY)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = hdr
        set_font(r, 11, bold=True, color=COLOR_WHITE)

    # データ行（ストライプ）
    for i, row_data in enumerate(rows):
        tbl.rows[i + 1].height = TABLE_DATA_ROW_H
        bg = COLOR_STRIPE if i % 2 == 0 else COLOR_WHITE
        for j in range(n_cols):
            cell = tbl.cell(i + 1, j)
            set_cell_bg(cell, bg)
            cell_text = row_data[j] if j < len(row_data) else ''
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = cell_text
            set_font(r, 11, color=COLOR_TEXT)

    return table_h


def render_columns(slide, item, x, y, w, h):
    """
    2カラムレイアウトを描画。
    左右のカラムタイトル（任意）＋各カラムのコンテンツ＋中央縦区切り線。
    戻り値: 使用した高さ
    """
    GAP   = Inches(0.35)
    col_w = (w - GAP) / 2
    left_x  = x
    right_x = x + col_w + GAP

    title_h = Inches(0.38) if (item.get('left_title') or item.get('right_title')) else Inches(0)

    # カラムタイトル
    for col_x, title_text in ((left_x, item.get('left_title')), (right_x, item.get('right_title'))):
        if title_text:
            tb = slide.shapes.add_textbox(col_x, y, col_w, Inches(0.35))
            tf = tb.text_frame
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = title_text
            set_font(r, 13, bold=True, color=COLOR_DARK_NAVY)

    # 中央の縦区切り線
    div_x = x + col_w + GAP / 2 - Pt(0.5)
    add_rect(slide, div_x, y, Pt(1.0), h, COLOR_GRAY_LIGHT)

    content_y = y + title_h
    content_h = h - title_h

    if item.get('left'):
        render_text_items(slide, item['left'], left_x, content_y, col_w, content_h)
    if item.get('right'):
        render_text_items(slide, item['right'], right_x, content_y, col_w, content_h)

    return h


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

    add_rect(slide, 0, 0, SLIDE_WIDTH, split_y, COLOR_DARK_NAVY)
    add_rect(slide, 0, split_y, SLIDE_WIDTH, SLIDE_HEIGHT - split_y, COLOR_WHITE)
    add_rect(slide, 0, split_y - Pt(3), SLIDE_WIDTH, Pt(6), COLOR_NAVY)

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
        tb_num = slide.shapes.add_textbox(
            Inches(1.2), center_y - Inches(0.55),
            SLIDE_WIDTH - Inches(2.4), Inches(0.35)
        )
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        r_num = p_num.add_run()
        r_num.text = section_num
        set_font(r_num, 11, bold=True, color=COLOR_BLUE)

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

    line_y = center_y + Inches(0.85) if not section_num else center_y + Inches(1.05)
    add_rect(slide, Inches(1.2), line_y, Inches(2.5), Pt(3), COLOR_BLUE)

    add_logo(slide, logo_path)
    add_footer(slide, page_num)
    return slide


def make_content_slide(prs, title, section_num, items, logo_path, page_num=None):
    """
    コンテンツスライド
    ・純白背景
    ・節番号（小グレー、任意）＋ メインタイトル（大・黒）
    ・ブラックヘッダーライン
    ・リード文 / 箇条書き / 番号付き / テーブル / 2カラム / 区切り線
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, COLOR_WHITE)

    # ── タイトルエリア ──
    if section_num:
        tb_sec = slide.shapes.add_textbox(
            TITLE_AREA_X, Inches(0.07), TITLE_AREA_MAX_W, Inches(0.24)
        )
        tf_sec = tb_sec.text_frame
        p_sec = tf_sec.paragraphs[0]
        r_sec = p_sec.add_run()
        r_sec.text = section_num
        set_font(r_sec, 9, bold=True, color=COLOR_SECTION_NUM)
        main_title_y = Inches(0.27)
    else:
        main_title_y = Inches(0.10)

    tb_title = slide.shapes.add_textbox(
        TITLE_AREA_X, main_title_y, TITLE_AREA_MAX_W, Inches(0.62)
    )
    tf_title = tb_title.text_frame
    tf_title.word_wrap = False
    p_title = tf_title.paragraphs[0]
    r_title = p_title.add_run()
    r_title.text = title
    set_font(r_title, 26, bold=True, color=COLOR_TEXT)

    # ── ヘッダーライン ──
    add_rect(slide, 0, HEADER_LINE_Y, SLIDE_WIDTH, HEADER_LINE_H, COLOR_BLACK)

    # ── コンテンツエリア ──
    lead_items    = [i for i in items if i.get('type') == 'lead']
    content_items = [i for i in items if i.get('type') != 'lead']

    current_y = HEADER_LINE_Y + Inches(0.18)

    # リード文
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

    # ── セグメント分割 ──
    # table / columns は独立セグメント、その他はテキストグループにまとめる
    segments = []
    text_buffer = []
    for item in content_items:
        if item.get('type') in ('table', 'columns'):
            if text_buffer:
                segments.append(('text', list(text_buffer)))
                text_buffer = []
            segments.append((item['type'], item))
        else:
            text_buffer.append(item)
    if text_buffer:
        segments.append(('text', text_buffer))

    available_h = FOOTER_LINE_Y - current_y - Inches(0.1)

    # テーブルの固定高さを先に計算
    fixed_h = sum(
        TABLE_HEADER_ROW_H + TABLE_DATA_ROW_H * len(data.get('rows', [])) + Inches(0.12)
        for seg_type, data in segments
        if seg_type == 'table'
    )

    # 残り高さをテキスト/カラムセグメントで等分
    n_var = sum(1 for t, _ in segments if t in ('text', 'columns'))
    var_h = (available_h - fixed_h) / max(n_var, 1) if n_var else available_h

    # ── セグメント描画 ──
    y = current_y
    for seg_type, seg_data in segments:
        if seg_type == 'text':
            render_text_items(slide, seg_data, CONTENT_X, y, CONTENT_RW, var_h)
            y += var_h
        elif seg_type == 'table':
            h = render_table(slide, seg_data, CONTENT_X, y, CONTENT_RW)
            y += h + Inches(0.12)
        elif seg_type == 'columns':
            render_columns(slide, seg_data, CONTENT_X, y, CONTENT_RW, var_h)
            y += var_h

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
    slides        = []
    current       = None
    col_side      = None   # None | 'left' | 'right'
    table_header  = None   # 保留中のテーブルヘッダー行
    table_active  = None   # 処理中のテーブルアイテム

    def flush_table():
        nonlocal table_header, table_active
        if table_active and current and current.get('type') == 'content':
            _append_item(table_active)
        table_header = None
        table_active = None

    def _append_item(item):
        """col_side に応じてアイテムを適切な場所に追加"""
        if (col_side
                and current.get('items')
                and current['items'][-1].get('type') == 'columns'):
            current['items'][-1][col_side].append(item)
        else:
            current['items'].append(item)

    for raw_line in md_text.splitlines():
        line = raw_line.strip()

        # ── テーブル行の検出（| で始まる行）──
        if line.startswith('|'):
            if re.match(r'^[\|\s\-:]+$', line):
                # 区切り行（|---|---|）→ ヘッダー確定
                if table_header is not None and current and current.get('type') == 'content':
                    table_active = {
                        'type': 'table',
                        'headers': table_header,
                        'rows': []
                    }
                    table_header = None
            else:
                # データ行
                cells = [c.strip() for c in line.strip('|').split('|')]
                if table_active is not None:
                    table_active['rows'].append(cells)
                else:
                    flush_table()
                    table_header = cells
            continue

        # テーブル以外の行が来たらテーブルを確定
        flush_table()

        # ── H3 → セクション区切りスライド ──
        if line.startswith('### '):
            if current:
                slides.append(current)
            raw_title = line[4:].strip()
            if ' | ' in raw_title:
                sec, ttl = raw_title.split(' | ', 1)
                current = {'type': 'section', 'title': ttl.strip(),
                           'section_num': sec.strip(), 'items': []}
            else:
                current = {'type': 'section', 'title': raw_title,
                           'section_num': None, 'items': []}
            col_side = None

        # ── H2 → コンテンツスライド ──
        elif line.startswith('## '):
            if current:
                slides.append(current)
            raw_title = line[3:].strip()
            if ' | ' in raw_title:
                sec, ttl = raw_title.split(' | ', 1)
                current = {'type': 'content', 'title': ttl.strip(),
                           'section_num': sec.strip(), 'items': []}
            else:
                current = {'type': 'content', 'title': raw_title,
                           'section_num': None, 'items': []}
            col_side = None

        # ── H1 → タイトルスライド ──
        elif line.startswith('# '):
            if current:
                slides.append(current)
            current = {
                'type': 'title',
                'title': line[2:].strip(),
                'subtitle': '',
                'items': [],
            }
            col_side = None

        # ── 2カラムレイアウト ──
        elif line.startswith('::left'):
            col_side  = 'left'
            col_title = line[6:].strip() or None
            if current and current.get('type') == 'content':
                if not current['items'] or current['items'][-1].get('type') != 'columns':
                    current['items'].append({
                        'type': 'columns',
                        'left': [], 'right': [],
                        'left_title': col_title, 'right_title': None
                    })
                elif col_title:
                    current['items'][-1]['left_title'] = col_title

        elif line.startswith('::right'):
            col_side  = 'right'
            col_title = line[7:].strip() or None
            if current and current.get('type') == 'content':
                if not current['items'] or current['items'][-1].get('type') != 'columns':
                    current['items'].append({
                        'type': 'columns',
                        'left': [], 'right': [],
                        'left_title': None, 'right_title': col_title
                    })
                elif col_title:
                    current['items'][-1]['right_title'] = col_title

        elif line == ':::':
            col_side = None

        # ── リード文（> blockquote）──
        elif line.startswith('> '):
            text = line[2:].strip()
            if current and current.get('type') == 'content':
                current['items'].append({'type': 'lead', 'text': text})

        # ── 区切り線 --- ──
        elif line == '---':
            if current and current.get('type') == 'content':
                _append_item({'type': 'divider'})

        # ── 字下げ箇条書き（インデント付き - or *）──
        elif re.match(r'^\s{2,}[\*\-] ', raw_line):
            text = re.sub(r'^\s+[\*\-] ', '', raw_line).strip()
            if current:
                _append_item({'type': 'sub_bullet', 'text': text})

        # ── 標準箇条書き（* または -）──
        elif re.match(r'^[\*\-] ', line):
            text = line[2:].strip()
            if current:
                _append_item({'type': 'bullet', 'text': text})

        # ── 番号付きリスト ──
        elif re.match(r'^\d+\. ', line):
            text = re.sub(r'^\d+\. ', '', line).strip()
            # 同一コンテキスト内での連番
            if (col_side
                    and current.get('items')
                    and current['items'][-1].get('type') == 'columns'):
                col_items = current['items'][-1][col_side]
                num = len([i for i in col_items if i.get('type') == 'numbered']) + 1
            else:
                num = len([i for i in current.get('items', [])
                           if i.get('type') == 'numbered']) + 1
            if current:
                _append_item({'type': 'numbered', 'text': text, 'num': num})

        # ── タイトルスライドのサブタイトル行 ──
        elif line and current and current['type'] == 'title':
            current['subtitle'] += line + ' '

        # ── その他テキスト（コンテンツスライドの平文）──
        elif line and current and current.get('type') == 'content':
            _append_item({'type': 'text', 'text': line})

    flush_table()
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
