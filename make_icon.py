# -*- coding: utf-8 -*-
"""生成应用图标 app.ico —— 几何绘制：蓝底 + 白色表格 + 图片意象

用 PIL 多倍超采样绘制后缩放，保证边缘平滑；输出多尺寸 ICO。
不使用 AI 生图，保证可复现、无外网依赖。
"""
import os
import sys
from PIL import Image, ImageDraw

S = 1024                      # 超采样画布
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
os.makedirs(OUT_DIR, exist_ok=True)

BLUE_D = (32, 62, 122)        # 深蓝（渐变底）
BLUE_L = (68, 114, 196)       # 亮蓝
WHITE = (255, 255, 255)
LINE = (150, 175, 220)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# ---- 圆角方形底 + 竖向渐变
R = int(S * 0.22)
mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=R, fill=255)

grad = Image.new("RGBA", (S, S))
gd = ImageDraw.Draw(grad)
for y in range(S):
    gd.line([(0, y), (S, y)], fill=lerp(BLUE_L, BLUE_D, y / S))
img.paste(grad, (0, 0), mask)

# ---- 表格卡片（白色圆角矩形，略微内缩）
pad = int(S * 0.17)
card = [pad, int(S * 0.22), S - pad, S - int(S * 0.18)]
cr = int(S * 0.05)
d.rounded_rectangle(card, radius=cr, fill=WHITE)

# ---- 表格外框
d.rounded_rectangle(card, radius=cr, outline=LINE, width=int(S * 0.012))

# ---- 表头横条
head_h = int((card[3] - card[1]) * 0.26)
d.rounded_rectangle([card[0], card[1], card[2], card[1] + head_h],
                    radius=cr, fill=BLUE_D)
# 把表头下缘直角化（只保留上方圆角）
d.rectangle([card[0], card[1] + head_h - cr, card[2], card[1] + head_h],
            fill=BLUE_D)

# ---- 表头内白色短横线（示意文字）
hy = card[1] + head_h // 2
lw = int(S * 0.014)
d.rounded_rectangle([card[0] + int(S * 0.045), hy - lw,
                     card[0] + int(S * 0.20), hy + lw],
                    radius=lw, fill=WHITE)

# ---- 表格网格：3 行 × 3 列
body_top = card[1] + head_h
body_bot = card[3]
rows, cols = 3, 3
cw = (card[2] - card[0]) / cols
ch = (body_bot - body_top) / rows
gw = int(S * 0.008)
# 竖线
for i in range(1, cols):
    x = card[0] + cw * i
    d.rectangle([x - gw, body_top, x + gw, body_bot], fill=LINE)
# 横线
for j in range(1, rows):
    y = body_top + ch * j
    d.rectangle([card[0], y - gw, card[2], y + gw], fill=LINE)

# ---- 单元格内的数据短线（灰色，示意内容）
dh = int(S * 0.012)
for r_i, frac in enumerate([0.60, 0.45, 0.66, 0.52, 0.72, 0.40, 0.58, 0.50, 0.62]):
    row = r_i // 3
    col = r_i % 3
    cx0 = card[0] + cw * col + cw * 0.16
    cx1 = cx0 + cw * frac * 0.7
    cy = body_top + ch * row + ch * 0.5
    d.rounded_rectangle([cx0, cy - dh, cx1, cy + dh],
                        radius=dh, fill=(190, 205, 230))

# ---- 右下角「图片」角标：小照片 + 山与太阳
bw = int(S * 0.30)
bx1, by1 = S - int(S * 0.10), S - int(S * 0.10)
bx0, by0 = bx1 - bw, by1 - bw
br = int(S * 0.045)
# 角标白底 + 蓝边
d.rounded_rectangle([bx0 - int(S * 0.018), by0 - int(S * 0.018),
                     bx1 + int(S * 0.018), by1 + int(S * 0.018)],
                    radius=br, fill=WHITE)
d.rounded_rectangle([bx0, by0, bx1, by1], radius=int(S * 0.03), fill=BLUE_L)
# 太阳
sr = int(bw * 0.13)
sx, sy = bx0 + int(bw * 0.30), by0 + int(bw * 0.30)
d.ellipse([sx - sr, sy - sr, sx + sr, sy + sr], fill=WHITE)
# 山（两个三角）
d.polygon([(bx0 + int(bw * 0.10), by1 - int(bw * 0.14)),
           (bx0 + int(bw * 0.42), by0 + int(bw * 0.42)),
           (bx0 + int(bw * 0.74), by1 - int(bw * 0.14))], fill=WHITE)
d.polygon([(bx0 + int(bw * 0.50), by1 - int(bw * 0.14)),
           (bx0 + int(bw * 0.72), by0 + int(bw * 0.55)),
           (bx0 + int(bw * 0.95), by1 - int(bw * 0.14))], fill=(225, 235, 250))

# ---- 输出多尺寸 ICO
ico_path = os.path.join(OUT_DIR, "app.ico")
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
img.resize((256, 256), Image.LANCZOS).save(
    ico_path, format="ICO", sizes=sizes)
# 另存一份 PNG 便于预览
png_path = os.path.join(OUT_DIR, "app_icon_preview.png")
img.resize((256, 256), Image.LANCZOS).save(png_path)

print("ICO:", ico_path, os.path.getsize(ico_path), "bytes")
print("PNG:", png_path, os.path.getsize(png_path), "bytes")
