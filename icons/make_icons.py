#!/usr/bin/env python3
"""まなびクエスト PWA アイコン生成（標準ライブラリのみ）。

デザイン: 青のグラデーション角丸背景 + 白い4方向のきらめき（bi-stars 風）。
3倍スーパーサンプリングでアンチエイリアスします。
"""
import math
import os
import struct
import zlib

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
SS = 3  # supersampling


def lerp(a, b, t):
    return a + (b - a) * t


def rounded_rect_alpha(x, y, w, h, r):
    """角丸長方形の内側なら 1、外なら 0（w×h の座標系、左上原点）。"""
    if x < 0 or y < 0 or x > w or y > h:
        return 0
    cx = min(max(x, r), w - r)
    cy = min(max(y, r), h - r)
    dx = x - cx
    dy = y - cy
    if dx == 0 or dy == 0:
        return 1
    return 1 if (dx * dx + dy * dy) <= r * r else 0


def sparkle_alpha(px, py, cx, cy, radius):
    """アステロイド曲線による4方向のきらめき。中心 (cx, cy)、半径 radius。"""
    dx = abs(px - cx) / radius
    dy = abs(py - cy) / radius
    if dx > 1 or dy > 1:
        return 0
    return 1 if (math.sqrt(dx) + math.sqrt(dy)) <= 1.0 else 0


def render(size, maskable=False, opaque_bg=False):
    """RGBA バイト列を返します。"""
    w = h = size
    big = size * SS
    # 背景の角丸半径（maskable は全面塗りなので角丸なし）
    radius = 0 if maskable else big * 0.225
    # 前景の配置（maskable は安全領域 = 中央80%に収める）
    scale = 0.62 if maskable else 0.80
    cx, cy = big / 2.0, big / 2.0

    main_r = big * 0.30 * scale
    sub1 = (cx + big * 0.26 * scale, cy - big * 0.26 * scale, big * 0.115 * scale)
    sub2 = (cx - big * 0.27 * scale, cy + big * 0.25 * scale, big * 0.085 * scale)

    top = (0x6F, 0xB1, 0xF7)
    bottom = (0x36, 0x74, 0xCE)
    fg = (0xFF, 0xFF, 0xFF)

    rows = []
    for y in range(h):
        row = bytearray()
        for x in range(w):
            ar = ag = ab = aa = 0.0
            for sy in range(SS):
                for sx in range(SS):
                    bx = x * SS + sx + 0.5
                    by = y * SS + sy + 0.5
                    inside = rounded_rect_alpha(bx, by, big, big, radius)
                    if not inside:
                        continue
                    t = by / big
                    r = lerp(top[0], bottom[0], t)
                    g = lerp(top[1], bottom[1], t)
                    b = lerp(top[2], bottom[2], t)
                    if (sparkle_alpha(bx, by, cx, cy, main_r)
                            or sparkle_alpha(bx, by, sub1[0], sub1[1], sub1[2])
                            or sparkle_alpha(bx, by, sub2[0], sub2[1], sub2[2])):
                        r, g, b = fg
                    ar += r
                    ag += g
                    ab += b
                    aa += 1.0
            n = SS * SS
            if aa == 0:
                if opaque_bg:
                    row += bytes((bottom[0], bottom[1], bottom[2], 255))
                else:
                    row += b"\x00\x00\x00\x00"
                continue
            # 被覆率で色を平均し、アルファは被覆率
            cov = aa / n
            px = (int(round(ar / aa)), int(round(ag / aa)), int(round(ab / aa)))
            if opaque_bg:
                # 不透明化: 背景色と合成（iOS のホーム画面アイコンは透過を黒く塗るため）
                a = cov
                px = tuple(int(round(lerp(bottom[i], px[i], a))) for i in range(3))
                row += bytes((px[0], px[1], px[2], 255))
            else:
                row += bytes((px[0], px[1], px[2], int(round(cov * 255))))
        rows.append(bytes(row))
    return rows


def write_png(path, rows, size):
    raw = b"".join(b"\x00" + r for r in rows)
    comp = zlib.compress(raw, 9)

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", comp)
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)
    print(f"{path} ({len(png)} bytes)")


def main():
    os.makedirs(OUT, exist_ok=True)
    targets = [
        ("icon-192.png", 192, False, False),
        ("icon-512.png", 512, False, False),
        ("icon-maskable-192.png", 192, True, True),
        ("icon-maskable-512.png", 512, True, True),
        ("apple-touch-icon.png", 180, False, True),
        ("favicon-32.png", 32, False, False),
    ]
    for name, size, maskable, opaque in targets:
        write_png(os.path.join(OUT, name), render(size, maskable, opaque), size)


if __name__ == "__main__":
    main()
