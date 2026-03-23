"""
SALESCOREロゴをPILで生成する
実際のロゴ: 青いフレーム/水滴アイコン + SALESCORE テキスト
"""
from PIL import Image, ImageDraw, ImageFont
import math

W, H = 520, 120

img = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw = ImageDraw.Draw(img)

# --- アイコン部分 (左側) ---
# SALESCOREロゴは2つの水滴/炎が重なったデザイン
# 色: 濃い青 #1a4fd6 〜 明るい青 #4fa3f7

def draw_flame(draw, cx, cy, w, h, color):
    """水滴/炎の形を描く"""
    points = []
    # 上部は尖った炎の形、下部は丸い
    steps = 60
    for i in range(steps + 1):
        t = i / steps * 2 * math.pi
        # 炎っぽい形: 上が細く、下が丸い
        if t < math.pi:
            # 上半分: 尖らせる
            r_w = w * 0.5 * math.sin(t) * (1 - 0.3 * math.sin(t))
            r_h = -h * 0.5 * math.cos(t)
        else:
            # 下半分: 丸く
            r_w = w * 0.5 * math.sin(t)
            r_h = -h * 0.5 * math.cos(t)
        points.append((cx + r_w, cy + r_h))
    draw.polygon(points, fill=color)

# 後ろの水滴 (やや右・下にオフセット、明るい青)
cx1, cy1 = 42, 62
draw_flame(draw, cx1, cy1, 48, 80, (79, 163, 247, 220))   # 明るい青

# 前の水滴 (やや左・上, 濃い青)
cx2, cy2 = 30, 58
draw_flame(draw, cx2, cy2, 44, 76, (26, 79, 214, 255))    # 濃い青

# --- テキスト部分 ---
# フォントを探す
import os
font_candidates = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
]
font = None
for fc in font_candidates:
    if os.path.exists(fc):
        try:
            font = ImageFont.truetype(fc, 52)
            break
        except:
            pass
if font is None:
    font = ImageFont.load_default()

text = "SALESCORE"
text_color = (30, 30, 35, 255)  # ほぼ黒

# テキスト配置
text_x = 82
text_y = 30
draw.text((text_x, text_y), text, font=font, fill=text_color)

# 保存
out_path = "/home/user/-/salescore_logo.png"
img.save(out_path)
print(f"Saved: {out_path}  ({W}x{H})")
