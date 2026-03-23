"""
SALESCOREロゴ v2 - 実際のロゴに近い再現
炎+水滴の複合アイコン（ブルー）+ SALESCORE ボールドテキスト
"""
from PIL import Image, ImageDraw, ImageFilter, ImageFont
import math, os


def flame_points(cx, cy, w, h, n=120):
    """炎の形状のポリゴンポイントを生成"""
    pts = []
    for i in range(n):
        t = i / n * 2 * math.pi
        # 基本の楕円
        base_x = math.sin(t)
        base_y = -math.cos(t)

        # 上部を尖らせる（cos が正→上側）
        top_factor = 1.0
        if base_y < 0:  # 上半分
            sharpness = (-base_y) ** 0.5  # 0→1 as we go up
            top_factor = 1.0 - sharpness * 0.45 * abs(base_x)

        rx = (w / 2) * base_x * top_factor
        ry = (h / 2) * base_y

        # 少し右に傾ける
        angle = math.radians(8)
        rx2 = rx * math.cos(angle) - ry * math.sin(angle)
        ry2 = rx * math.sin(angle) + ry * math.cos(angle)

        pts.append((cx + rx2, cy + ry2))
    return pts


def make_logo(output_path, dark_bg=False):
    LW, LH = 560, 130
    img = Image.new("RGBA", (LW, LH), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # ── アイコン描画 ──────────────────────────────────────
    ICW = 105  # アイコン領域幅
    ICH = 105
    ix0 = 5    # アイコン開始X
    iy0 = 12   # アイコン開始Y
    cx = ix0 + ICW * 0.5
    cy = iy0 + ICH * 0.53

    # 後ろの炎 (やや右下にオフセット, 明るい青)
    pts_back = flame_points(cx + 10, cy + 4, ICW * 0.72, ICH * 0.88)
    draw.polygon(pts_back, fill=(74, 144, 226, 200))

    # 前の炎 (メイン, 濃い青)
    pts_front = flame_points(cx - 3, cy, ICW * 0.68, ICH * 0.85)
    draw.polygon(pts_front, fill=(26, 79, 200, 255))

    # 内側の弧 (光の反射/水の要素) - lighter blue oval
    arc_w = ICW * 0.38
    arc_h = ICH * 0.38
    draw.ellipse([
        cx - arc_w/2 - 5, cy + 4 - arc_h/2,
        cx + arc_w/2 - 5, cy + 4 + arc_h/2,
    ], fill=(160, 200, 240, 200))

    # 内側の小さい白い弧
    arc2_w = ICW * 0.20
    arc2_h = ICH * 0.20
    draw.ellipse([
        cx - arc2_w/2 - 8, cy + 10 - arc2_h/2,
        cx + arc2_w/2 - 8, cy + 10 + arc2_h/2,
    ], fill=(220, 235, 255, 200))

    # 軽くスムーズ
    # img = img.filter(ImageFilter.SMOOTH)

    # ── テキスト描画 ─────────────────────────────────────
    draw2 = ImageDraw.Draw(img)
    text_color = (255, 255, 255, 255) if dark_bg else (12, 20, 50, 255)

    font_candidates = [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    ]
    font = None
    for fc in font_candidates:
        if os.path.exists(fc):
            try:
                font = ImageFont.truetype(fc, 62)
                break
            except:
                pass
    if font is None:
        font = ImageFont.load_default()

    draw2.text((122, 32), "SALESCORE", font=font, fill=text_color)

    img.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    make_logo("/home/user/-/salescore_logo.png", dark_bg=False)
    make_logo("/home/user/-/salescore_logo_white.png", dark_bg=True)
