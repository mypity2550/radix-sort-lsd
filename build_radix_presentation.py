"""
Generate High-Resolution Diagrams/Charts + 5-Minute PowerPoint Presentation
for Radix Sort (LSD)
===========================================================================
Reads real benchmark numbers from `benchmark_results.json`, renders crisp
antialiased 1:1-aspect diagrams/charts with Pillow (2x supersampling), and
builds `RadixSort_Presentation.pptx` (8 slides, 16:9, Thai + Speaker Notes ~5 min).
"""

from __future__ import annotations

import json
import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE_DIR = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
ASSETS_DIR.mkdir(exist_ok=True)

FONT_TH_REG = r"C:\Windows\Fonts\LeelawUI.ttf"
FONT_TH_BOLD = r"C:\Windows\Fonts\LeelaUIb.ttf"
FONT_MONO = r"C:\Windows\Fonts\consola.ttf"
FONT_MONO_BOLD = r"C:\Windows\Fonts\consolab.ttf"

# Theme palette (Midnight Tech Slate — Sound of Sorting inspired)
C_BG = (11, 17, 32)         # #0B1120
C_CARD = (19, 28, 49)       # #131C31
C_CARD2 = (15, 23, 42)      # #0F172A
C_BORDER = (38, 53, 82)     # #263552
C_TEXT = (248, 250, 252)    # #F8FAFC
C_SUB = (203, 213, 225)     # #CBD5E1
C_MUTED = (148, 163, 184)   # #94A3B8
C_CYAN = (56, 189, 248)     # #38BDF8
C_INDIGO = (129, 140, 248)  # #818CF8
C_GREEN = (16, 185, 129)    # #10B981
C_RED = (255, 77, 109)      # #FF4D6D
C_AMBER = (245, 158, 11)    # #F59E0B
C_PURPLE = (192, 132, 252)  # #C084FC


def rgb(t: tuple[int, int, int]) -> RGBColor:
    return RGBColor(t[0], t[1], t[2])


def load_font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


def draw_centered_text(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.FreeTypeFont,
    fill: tuple[int, int, int],
) -> None:
    x0, y0, x1, y1 = box
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = x0 + (x1 - x0 - tw) / 2 - bbox[0]
    ty = y0 + (y1 - y0 - th) / 2 - bbox[1]
    draw.text((tx, ty), text, font=font, fill=fill)


def draw_arrow(
    draw: ImageDraw.ImageDraw,
    p1: tuple[float, float],
    p2: tuple[float, float],
    color: tuple[int, int, int],
    width: int = 4,
    head: int = 12,
) -> None:
    draw.line([p1, p2], fill=color, width=width)
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    ang = math.atan2(dy, dx)
    a1 = ang + math.pi * 0.82
    a2 = ang - math.pi * 0.82
    draw.polygon(
        [
            p2,
            (p2[0] + head * math.cos(a1), p2[1] + head * math.sin(a1)),
            (p2[0] + head * math.cos(a2), p2[1] + head * math.sin(a2)),
        ],
        fill=color,
    )


def draw_highlighted_3digit_box(
    draw: ImageDraw.ImageDraw,
    x0: int,
    y0: int,
    w: int,
    h: int,
    num_str: str,
    hl_idx: int,  # -1 none, 0 hundreds, 1 tens, 2 ones
    hl_color: tuple[int, int, int],
    box_fill: tuple[int, int, int],
    box_outline: tuple[int, int, int],
    font_mono: ImageFont.FreeTypeFont,
) -> None:
    draw.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=12, fill=box_fill, outline=box_outline, width=3)
    # Render 3 characters evenly spaced so the active digit gets a glowing pill highlight
    char_w = w / 3.4
    start_x = x0 + (w - 3 * char_w) / 2
    for idx, ch in enumerate(num_str):
        cx0 = int(start_x + idx * char_w)
        cx1 = int(cx0 + char_w)
        if idx == hl_idx:
            draw.rounded_rectangle([cx0 + 2, y0 + 10, cx1 - 2, y0 + h - 10], radius=8, fill=hl_color)
            draw_centered_text(draw, (cx0, y0, cx1, y0 + h), ch, font_mono, (8, 12, 22))
        else:
            col = C_MUTED if (hl_idx >= 0 and ch == "0" and idx == 0 and num_str.startswith("0")) else C_TEXT
            draw_centered_text(draw, (cx0, y0, cx1, y0 + h), ch, font_mono, col)


# =============================================================================
# Asset 1: Radix Sort (LSD) Step-by-Step 3-Pass Diagram (Slide 3)
# Exact 1:1 aspect matching (12.23 in x 5.72 in) -> 2446 x 1144
# =============================================================================
def create_radix_step_diagram() -> Path:
    W, H = 2446, 1144
    img = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(img)

    f_stage = load_font(FONT_TH_BOLD, 31)
    f_sub = load_font(FONT_TH_REG, 24)
    f_num = load_font(FONT_MONO_BOLD, 42)
    f_bkt = load_font(FONT_TH_BOLD, 22)
    f_note = load_font(FONT_TH_BOLD, 24)

    stages = [
        (
            "ข้อมูลตั้งต้น (Input)",
            "ตัวเลข 8 ตัว (เติม 0 ด้านหน้าให้เห็นครบ 3 หลัก)",
            ["170", "045", "075", "090", "802", "024", "002", "066"],
            -1,
            C_MUTED,
            (18, 26, 46),
            (58, 76, 110),
            None,
            "ยังไม่ได้เรียงลำดับ",
        ),
        (
            "รอบที่ 1: เรียงตามหลักหน่วย (x1)",
            "ดูเลขขวาสุด ใส่ถัง 0–9 แล้วดึงออกตามลำดับถัง",
            ["170", "090", "802", "002", "024", "045", "075", "066"],
            2,
            C_CYAN,
            (13, 28, 48),
            C_CYAN,
            ["ถัง 0", "ถัง 0", "ถัง 2", "ถัง 2", "ถัง 4", "ถัง 5", "ถัง 5", "ถัง 6"],
            "หลักหน่วยเรียง 0, 0, 2, 2, 4, 5, 5, 6 (170 อยู่หน้า 090 — Stable!)",
        ),
        (
            "รอบที่ 2: เรียงตามหลักสิบ (x10)",
            "ดูเลขตรงกลาง โดยคงลำดับเดิมจากรอบที่ 1 ไว้",
            ["802", "002", "024", "045", "066", "170", "075", "090"],
            1,
            C_AMBER,
            (32, 26, 16),
            C_AMBER,
            ["ถัง 0", "ถัง 0", "ถัง 2", "ถัง 4", "ถัง 6", "ถัง 7", "ถัง 7", "ถัง 9"],
            "หลักสิบเรียง 0, 0, 2, 4, 6, 7, 7, 9 (170 ยังคงอยู่หน้า 075!)",
        ),
        (
            "รอบที่ 3: เรียงตามหลักร้อย (x100)",
            "ดูเลขซ้ายสุด — เรียงเสร็จสมบูรณ์โดยไม่เทียบค่าเลย!",
            ["002", "024", "045", "066", "075", "090", "170", "802"],
            0,
            C_GREEN,
            (11, 38, 32),
            C_GREEN,
            ["ถัง 0", "ถัง 0", "ถัง 0", "ถัง 0", "ถัง 0", "ถัง 0", "ถัง 1", "ถัง 8"],
            "ผลลัพธ์สุดท้าย: [2, 24, 45, 66, 75, 90, 170, 802]  (0 Comparisons!)",
        ),
    ]

    row_h = 242
    gap_y = 44
    top_start = 24

    for s_idx, (title, sub, vals, hl_idx, col, bg_band, border_col, bkts, footer_txt) in enumerate(stages):
        y0 = top_start + s_idx * (row_h + gap_y)
        y1 = y0 + row_h
        d.rounded_rectangle([18, y0, W - 18, y1], radius=22, fill=bg_band, outline=border_col, width=3)

        # Left text panel
        d.text((48, y0 + 28), title, font=f_stage, fill=col if hl_idx >= 0 else C_TEXT)
        d.text((48, y0 + 78), sub, font=f_sub, fill=C_SUB)
        d.text((48, y0 + 184), footer_txt, font=f_note, fill=col if hl_idx >= 0 else C_MUTED)

        # Array boxes on right
        bx_start = 740
        bw, bh = 186, 102
        gap_x = 24
        by = y0 + 36 if bkts else y0 + 62
        for i, num_s in enumerate(vals):
            bx = bx_start + i * (bw + gap_x)
            draw_highlighted_3digit_box(
                d, bx, by, bw, bh, num_s, hl_idx, col,
                (16, 24, 42), border_col, f_num,
            )
            if bkts:
                pill_y0 = by + bh + 12
                d.rounded_rectangle([bx + 26, pill_y0, bx + bw - 26, pill_y0 + 36], radius=10, fill=(10, 16, 30), outline=col, width=2)
                draw_centered_text(d, (bx + 26, pill_y0, bx + bw - 26, pill_y0 + 36), bkts[i], f_bkt, col)

        # Down arrow to next row
        if s_idx < len(stages) - 1:
            ay0 = y1 + 4
            ay1 = y1 + gap_y - 4
            for ax in (1100, 1550, 2000):
                draw_arrow(d, (ax, ay0), (ax, ay1), stages[s_idx + 1][4], width=4, head=11)

    out = ASSETS_DIR / "radix_step_diagram.png"
    img.resize((1223, 572), Image.Resampling.LANCZOS).save(out)
    return out


# =============================================================================
# Asset 2: Radix Sort (LSD) Base-4 Simulator Preview (Slide 5)
# Exact 1:1 aspect matching (8.25 in x 4.86 in) -> 1800 x 1060
# =============================================================================
def create_radix_simulator_preview() -> Path:
    W, H = 1800, 1060
    img = Image.new("RGB", (W, H), (4, 6, 10))
    d = ImageDraw.Draw(img)

    f_hud = load_font(FONT_MONO_BOLD, 27)
    f_hud_sub = load_font(FONT_MONO, 23)
    f_lbl = load_font(FONT_TH_BOLD, 23)
    f_bkt = load_font(FONT_MONO_BOLD, 22)

    # Simulate real Base-4 LSD Radix Sort on N=96 bars, paused mid-Pass-3 (Gather Bucket 2)!
    rng = random.Random(11)
    N = 96
    arr = list(range(1, N + 1))
    rng.shuffle(arr)

    # Complete Pass 1 (exp=1) and Pass 2 (exp=4)
    for exp in (1, 4):
        b = [[] for _ in range(4)]
        for x in arr:
            b[(x // exp) % 4].append(x)
        arr = [x for sub in b for x in sub]

    # Now in Pass 3 (exp=16): buckets for 0, 1, 2, 3
    buckets = [[] for _ in range(4)]
    for x in arr:
        buckets[(x // 16) % 4].append(x)
    # Write buckets 0, 1, and half of bucket 2 into arr; rest shows pre-Pass-3 waves!
    gathered = buckets[0] + buckets[1] + buckets[2][:12]
    active_write_idx = len(gathered) - 1
    bars = gathered + arr[len(gathered):]

    b_colors = [C_CYAN, C_INDIGO, C_PURPLE, (244, 114, 182)]
    b_bg = [(10, 32, 52), (20, 24, 54), (34, 20, 56), (48, 18, 40)]

    # Frame
    d.rounded_rectangle([8, 8, W - 8, H - 8], radius=24, fill=(4, 6, 10), outline=C_BORDER, width=4)

    # Top HUD
    d.rounded_rectangle([24, 22, W - 24, 132], radius=16, fill=(13, 20, 36), outline=(35, 52, 82), width=2)
    d.text((46, 36), "Radix Sort (LSD) — Base 4  (The Sound of Sorting Mode)", font=f_hud, fill=C_TEXT)
    d.text(
        (46, 82),
        "N = 96   |   Pass = 3 / 4 (exp=16)   |   Comparisons = 0!   |   Writes = 252",
        font=f_hud_sub,
        fill=C_GREEN,
    )

    d.rounded_rectangle([W - 395, 42, W - 44, 108], radius=14, fill=(40, 18, 28), outline=C_RED, width=2)
    draw_centered_text(d, (W - 395, 42, W - 44, 108), "PASS 3: GATHER BUCKET 2", f_hud_sub, (255, 160, 180))

    # Draw 4 bucket region bands
    pad_x = 46
    top_y = 195
    bot_y = H - 210
    draw_w = W - 2 * pad_x
    draw_h = bot_y - top_y
    bar_w = draw_w / N

    cur = 0
    for b_idx in range(4):
        blen = len(buckets[b_idx])
        lx = pad_x + cur * bar_w
        rx = pad_x + (cur + blen) * bar_w
        d.rectangle([lx, top_y - 12, rx, bot_y], fill=b_bg[b_idx])
        d.rectangle([lx, top_y - 16, max(lx + 2, rx - 2), top_y - 10], fill=b_colors[b_idx])
        d.text((lx + 18, top_y - 50), f"ถัง {b_idx} ({blen} ตัว)", font=f_lbl, fill=b_colors[b_idx])
        cur += blen

    # Draw bars
    for idx, val in enumerate(bars):
        bh = max(6, (val / N) * draw_h)
        x0 = pad_x + idx * bar_w + 1
        x1 = pad_x + (idx + 1) * bar_w - 2
        y0 = bot_y - bh
        col = (241, 245, 249)
        if idx == active_write_idx:
            col = C_RED
        elif idx < active_write_idx:
            col = (218, 245, 255)
        d.rectangle([x0, y0, x1, bot_y], fill=col)

    # Bottom Live 4-Bucket Strip
    strip_y0 = H - 190
    d.rounded_rectangle([24, strip_y0, W - 24, H - 22], radius=16, fill=(13, 20, 36), outline=(35, 52, 82), width=2)
    d.text((46, strip_y0 + 16), "Live Digit Buckets (Base 4 — ตรงกับวิดีโอ 15 Sorting Algorithms in 6 Minutes)", font=f_lbl, fill=C_SUB)

    remaining_counts = [0, 0, len(buckets[2]) - 12, len(buckets[3])]
    card_w = (W - 92 - 3 * 22) // 4
    for b_idx in range(4):
        cx0 = 46 + b_idx * (card_w + 22)
        cy0 = strip_y0 + 56
        cx1 = cx0 + card_w
        cy1 = H - 38
        is_act = b_idx == 2
        d.rounded_rectangle(
            [cx0, cy0, cx1, cy1], radius=12,
            fill=(42, 18, 30) if is_act else (9, 14, 26),
            outline=C_RED if is_act else b_colors[b_idx],
            width=3 if is_act else 2,
        )
        status_txt = "ดึงเสร็จแล้ว" if b_idx < 2 else ("กำลังเขียนคืน!" if b_idx == 2 else f"รอเขียน {remaining_counts[b_idx]} ตัว")
        d.text((cx0 + 18, cy0 + 12), f"Bucket {b_idx}", font=f_bkt, fill=b_colors[b_idx])
        d.text((cx0 + 18, cy0 + 46), status_txt, font=f_lbl, fill=C_TEXT if is_act else C_SUB)

    out = ASSETS_DIR / "radix_simulator_preview.png"
    img.resize((900, 530), Image.Resampling.LANCZOS).save(out)
    return out


# =============================================================================
# Asset 3: Real Benchmark Charts (Slide 6)
# Exact 1:1 aspect matching (12.23 in x 3.48 in) -> 2446 x 696
# =============================================================================
def create_radix_benchmark_chart(data: dict) -> Path:
    W, H = 2446, 696
    img = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(img)

    f_title = load_font(FONT_TH_BOLD, 30)
    f_axis = load_font(FONT_MONO_BOLD, 21)
    f_lbl = load_font(FONT_TH_REG, 22)
    f_lbl_b = load_font(FONT_TH_BOLD, 22)
    f_val = load_font(FONT_MONO_BOLD, 21)

    # Panel 1 (Left): Log-Log Scaling Comparison
    p1 = (12, 10, 1440, H - 10)
    d.rounded_rectangle(p1, radius=20, fill=C_CARD, outline=C_BORDER, width=3)
    d.text((44, 26), "เปรียบเทียบเวลา (Log-Log): Radix LSD vs Merge O(n log n) vs Selection O(n²)", font=f_title, fill=C_TEXT)

    gx0, gy0, gx1, gy1 = 165, 96, 1385, H - 105
    lx_min, lx_max = 3.82, 7.22
    ly_min, ly_max = -2.50, 4.95

    def to_px(n_val: float, sec_val: float) -> tuple[float, float]:
        lx = math.log10(n_val)
        ly = math.log10(sec_val)
        x = gx0 + (lx - lx_min) / (lx_max - lx_min) * (gx1 - gx0)
        y = gy1 - (ly - ly_min) / (ly_max - ly_min) * (gy1 - gy0)
        return (x, y)

    # Y ticks
    y_ticks = [
        (0.005, "5ms"),
        (0.01, "0.01s"),
        (0.1, "0.1s"),
        (1.0, "1s"),
        (10.0, "10s"),
        (100.0, "100s"),
        (1000.0, "1,000s"),
        (10000.0, "10,000s"),
    ]
    for yv, label in y_ticks:
        _, py = to_px(10_000, yv)
        if py < gy0 - 5 or py > gy1 + 5:
            continue
        d.line([(gx0, py), (gx1, py)], fill=(30, 44, 72), width=2)
        bbox = d.textbbox((0, 0), label, font=f_axis)
        d.text((gx0 - (bbox[2] - bbox[0]) - 14, py - 11), label, font=f_axis, fill=C_MUTED)

    # X ticks
    for xv, label in [(10_000, "10K"), (100_000, "100K"), (1_000_000, "1M"), (10_000_000, "10M")]:
        px, _ = to_px(xv, 1.0)
        d.line([(px, gy0), (px, gy1)], fill=(30, 44, 72), width=2)
        bbox = d.textbbox((0, 0), label, font=f_axis)
        d.text((px - (bbox[2] - bbox[0]) / 2, gy1 + 14), label, font=f_axis, fill=C_SUB)

    d.line([(gx0, gy0), (gx0, gy1), (gx1, gy1)], fill=C_MUTED, width=3)
    d.text((gx0 + (gx1 - gx0) / 2 - 120, gy1 + 48), "จำนวนตัวเลขสุ่ม (N)", font=f_lbl_b, fill=C_SUB)

    def draw_dashed(pts, color, width=4):
        for k in range(len(pts) - 1):
            pa, pb = pts[k], pts[k + 1]
            steps = 12
            for s in range(0, steps, 2):
                f0, f1 = s / steps, min(1.0, (s + 1) / steps)
                d.line(
                    [
                        (pa[0] + (pb[0] - pa[0]) * f0, pa[1] + (pb[1] - pa[1]) * f0),
                        (pa[0] + (pb[0] - pa[0]) * f1, pa[1] + (pb[1] - pa[1]) * f1),
                    ],
                    fill=color,
                    width=width,
                )

    # 1. Selection Sort O(n^2) dashed red
    sel_10k = data["comparison_summary"]["selection_sort_10k_sec"]
    sel_pts = [to_px(10 ** e, sel_10k * ((10 ** e / 10_000) ** 2)) for e in (4.0, 5.0, 6.0)]
    draw_dashed(sel_pts, C_RED, 5)
    d.ellipse([sel_pts[0][0] - 8, sel_pts[0][1] - 8, sel_pts[0][0] + 8, sel_pts[0][1] + 8], fill=C_RED)
    d.text((sel_pts[0][0] + 16, sel_pts[0][1] + 8), f"Selection 10K: {sel_10k:.2f}s", font=f_val, fill=C_RED)
    d.ellipse([sel_pts[2][0] - 8, sel_pts[2][1] - 8, sel_pts[2][0] + 8, sel_pts[2][1] + 8], outline=C_RED, width=4)
    d.text((sel_pts[2][0] + 18, sel_pts[2][1] - 14), "Selection 1M ~7.3 ชม.!", font=f_lbl_b, fill=C_RED)

    # 2. Merge Sort O(n log n) dashed indigo
    ref_m = data["reference_merge_sort"]
    m_pts = [to_px(int(k), v) for k, v in ref_m.items()]
    draw_dashed(m_pts, C_INDIGO, 4)
    for (px, py), (k, v) in zip(m_pts, ref_m.items()):
        d.ellipse([px - 6, py - 6, px + 6, py + 6], fill=C_INDIGO)
    d.text((m_pts[2][0] - 205, m_pts[2][1] - 42), f"Merge 1M: {ref_m['1000000']:.2f}s", font=f_val, fill=C_INDIGO)

    # 3. Radix Sort LSD Base 10 (Cyan)
    b10_pts = [(r["n"], r["avg_sec"], to_px(r["n"], r["avg_sec"])) for r in data["scaling_base10"]]
    d.line([p[2] for p in b10_pts], fill=C_CYAN, width=5)
    for nv, tv, (px, py) in b10_pts:
        d.ellipse([px - 7, py - 7, px + 7, py + 7], fill=C_BG, outline=C_CYAN, width=3)
    d.text((b10_pts[3][2][0] - 125, b10_pts[3][2][1] - 40), f"Base10: {b10_pts[3][1]:.1f}s", font=f_val, fill=C_CYAN)

    # 4. Radix Sort LSD Base 256 (Emerald Green — Hero curve!)
    b256_pts = [(r["n"], r["avg_sec"], to_px(r["n"], r["avg_sec"])) for r in data["scaling_base256"]]
    d.line([p[2] for p in b256_pts], fill=C_GREEN, width=7)
    for nv, tv, (px, py) in b256_pts:
        d.ellipse([px - 10, py - 10, px + 10, py + 10], fill=C_BG, outline=C_GREEN, width=4)
        d.ellipse([px - 4, py - 4, px + 4, py + 4], fill=C_GREEN)
        lbl = f"{tv*1000:.1f}ms" if tv < 0.1 else f"{tv:.2f}s"
        if nv == 10_000:
            ox, oy = 18, -14
        elif nv == 10_000_000:
            ox, oy = -110, 16
        else:
            ox, oy = 14, 14
        d.text((px + ox, py + oy), lbl, font=f_val, fill=C_GREEN)

    # Legend box (top left, completely clear of all lines)
    d.rounded_rectangle([188, 106, 700, 250], radius=12, fill=(10, 16, 30), outline=C_BORDER, width=2)
    leg = [
        (C_GREEN, "Radix LSD Base 256 — O(d(n+b))", False),
        (C_CYAN, "Radix LSD Base 10  — O(d(n+b))", False),
        (C_INDIGO, "Merge Sort อ้างอิง — O(n log n)", True),
        (C_RED, "Selection Sort อ้างอิง — O(n²)", True),
    ]
    for idx, (col, txt, dashed) in enumerate(leg):
        ly = 128 + idx * 32
        if dashed:
            d.line([(208, ly), (226, ly)], fill=col, width=4)
            d.line([(236, ly), (254, ly)], fill=col, width=4)
        else:
            d.line([(208, ly), (254, ly)], fill=col, width=5)
        d.text((268, ly - 13), txt, font=f_lbl_b, fill=col)

    # Panel 2 (Right): Effect of Radix Base b at N = 1,000,000
    p2 = (1466, 10, W - 12, H - 10)
    d.rounded_rectangle(p2, radius=20, fill=C_CARD, outline=C_BORDER, width=3)
    d.text((1498, 26), "ผลของการเลือกฐาน Radix (b)", font=f_title, fill=C_TEXT)
    d.text((1498, 68), "(ทดสอบ N = 1,000,000 ตัว ช่วง 0–999,999,999)", font=f_lbl, fill=C_MUTED)

    bases = data["base_comparison_1m"]
    b_thai = {
        4: "Base 4 (แบบในวิดีโอ — 15 รอบ)",
        10: "Base 10 (ฐานสิบ — 9 รอบ)",
        16: "Base 16 (ฐานสิบหก — 8 รอบ)",
        256: "Base 256 (ทีละไบต์ — 4 รอบ)",
    }
    b_cols = [C_PURPLE, C_CYAN, C_INDIGO, C_GREEN]
    max_t = max(x["time_sec"] for x in bases) * 1.15

    bx0 = 1502
    bmax_w = 875
    start_y = 126
    step_y = 135

    for idx, item in enumerate(bases):
        y = start_y + idx * step_y
        b = item["base"]
        t_val = item["time_sec"]
        col = b_cols[idx]
        d.text((bx0, y), b_thai[b], font=f_lbl_b, fill=C_TEXT)
        vt = f"{t_val:.2f} s"
        vb = d.textbbox((0, 0), vt, font=f_val)
        d.text((bx0 + bmax_w - (vb[2] - vb[0]), y + 2), vt, font=f_val, fill=col)

        d.rounded_rectangle([bx0, y + 36, bx0 + bmax_w, y + 76], radius=10, fill=(11, 18, 34))
        fw = max(24, int((t_val / max_t) * bmax_w))
        d.rounded_rectangle([bx0, y + 36, bx0 + fw, y + 76], radius=10, fill=col)
        per_pass = t_val / item["passes"]
        d.text((bx0, y + 83), f"Writes: {item['writes']:,}   |   เฉลี่ย {per_pass:.2f} วินาที/รอบ", font=f_lbl, fill=C_MUTED)

    out = ASSETS_DIR / "radix_benchmark_chart.png"
    img.resize((1223, 348), Image.Resampling.LANCZOS).save(out)
    return out


# =============================================================================
# Asset 4: Big-O Pass Cost Model + Linear Proof (Slide 7)
# Exact 1:1 aspect matching (7.73 in x 5.50 in) -> 1546 x 1100
# =============================================================================
def create_radix_bigo_diagram(data: dict) -> Path:
    W, H = 1546, 1100
    img = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(img)

    f_title = load_font(FONT_TH_BOLD, 28)
    f_mono_b = load_font(FONT_MONO_BOLD, 24)
    f_mono_s = load_font(FONT_MONO_BOLD, 20)
    f_th = load_font(FONT_TH_REG, 22)
    f_th_b = load_font(FONT_TH_BOLD, 23)

    # Top Card: Pass-by-Pass Work Breakdown -> O(d * (n + b))
    d.rounded_rectangle([10, 10, W - 10, 560], radius=22, fill=C_CARD, outline=C_BORDER, width=3)
    d.text((38, 28), "โครงสร้างต้นทุนการทำงาน:  T(n) = d × O(n + b)", font=f_title, fill=C_CYAN)

    passes = [
        ("Pass 1 (หลักขวาสุด)", "Scan N ตัวลงถัง b ใบ  +  รวบรวมคืน N ตัว", "O(n + b)", C_CYAN),
        ("Pass 2 (หลักถัดไป)", "Scan N ตัวลงถัง b ใบ  +  รวบรวมคืน N ตัว", "O(n + b)", C_INDIGO),
        ("Pass d (หลักซ้ายสุด)", "Scan N ตัวลงถัง b ใบ  +  รวบรวมคืน N ตัว", "O(n + b)", C_GREEN),
    ]
    ys = [104, 214, 384]
    for (p_lbl, p_desc, p_cost, col), py in zip(passes, ys):
        d.rounded_rectangle([185, py, 1210, py + 84], radius=14, fill=(14, 24, 46), outline=col, width=3)
        d.text((214, py + 12), p_lbl, font=f_th_b, fill=col)
        d.text((214, py + 46), p_desc, font=f_th, fill=C_SUB)
        d.rounded_rectangle([1235, py, 1495, py + 84], radius=14, fill=(11, 20, 38), outline=col, width=2)
        draw_centered_text(d, (1235, py, 1495, py + 84), p_cost, f_mono_b, col)

    # 3 vertical dots between Pass 2 and Pass d
    for dx in (700, 1365):
        for dy in (322, 340, 358):
            d.ellipse([dx - 5, dy - 5, dx + 5, dy + 5], fill=C_MUTED)

    # Left bracket for d passes
    bx = 152
    d.line([(bx, 114), (bx, 458)], fill=C_AMBER, width=4)
    d.line([(bx, 114), (bx + 18, 114)], fill=C_AMBER, width=4)
    d.line([(bx, 458), (bx + 18, 458)], fill=C_AMBER, width=4)
    d.text((34, 232), "ทำซ้ำ", font=f_th_b, fill=C_AMBER)
    d.text((30, 270), "d รอบ", font=f_th_b, fill=C_AMBER)
    d.text((22, 308), "log_b(M)", font=f_mono_s, fill=C_MUTED)

    # Summary strip inside Top Card
    d.rounded_rectangle([38, 488, W - 38, 540], radius=12, fill=(12, 38, 34), outline=C_GREEN, width=2)
    draw_centered_text(
        d, (38, 488, W - 38, 540),
        "รวมเวลาทำงาน = d x O(n + b)   |   Comparisons = 0  (ทะลุขีดจำกัด O(n log n) ของ Comparison Sort!)",
        f_th_b, C_GREEN,
    )

    # Bottom Card: Empirical Proof from real benchmark (Time per Pass across Bases!)
    d.rounded_rectangle([10, 586, W - 10, H - 10], radius=22, fill=C_CARD, outline=C_BORDER, width=3)
    d.text((38, 608), "พิสูจน์จากผลการรันจริง: เวลาต่อ 1 รอบ (N = 1,000,000) คงที่ ~0.28–0.30 วินาที!", font=f_title, fill=C_GREEN)

    bases = data["base_comparison_1m"]
    per_passes = [(x["base"], x["passes"], x["time_sec"] / x["passes"]) for x in bases]
    max_pp = max(p[2] for p in per_passes) * 1.35

    base_y = H - 86
    top_bar_y = 700
    col_w = 212
    bar_x0 = 42
    cols = [C_PURPLE, C_CYAN, C_INDIGO, C_GREEN]

    for idx, (b, passes_n, pp) in enumerate(per_passes):
        bh = int((pp / max_pp) * (base_y - top_bar_y))
        x0 = bar_x0 + idx * col_w + 26
        x1 = x0 + 148
        d.rounded_rectangle([x0, base_y - bh, x1, base_y], radius=12, fill=cols[idx])
        draw_centered_text(d, (x0 - 10, base_y - bh - 38, x1 + 10, base_y - bh - 4), f"{pp:.2f} s", f_mono_s, C_TEXT)
        draw_centered_text(d, (x0 - 10, base_y + 10, x1 + 10, base_y + 42), f"Base {b}", f_mono_s, cols[idx])
        draw_centered_text(d, (x0 - 10, base_y + 42, x1 + 10, base_y + 70), f"({passes_n} passes)", f_mono_s, C_MUTED)

    # Right insight box
    d.rounded_rectangle([920, 682, W - 36, H - 36], radius=16, fill=(14, 30, 44), outline=C_CYAN, width=2)
    d.text((950, 704), "ยืนยันสมการ T(n) = O(d(n + b))", font=f_th_b, fill=C_CYAN)
    d.text((950, 754), "• ไม่ว่าจะใช้ Base 4, 10, 16 หรือ 256", font=f_th, fill=C_TEXT)
    d.text((950, 790), "  เวลาต่อ 1 รอบเกือบเท่ากันเป๊ะ (~0.29 วิ)", font=f_th, fill=C_GREEN)
    d.text((950, 840), "• ลดจำนวนรอบ d จาก 15 เหลือ 4 รอบ", font=f_th, fill=C_TEXT)
    d.text((950, 876), "  ความเร็วพุ่งขึ้น 4.0 เท่าทันที (4.5s -> 1.1s)", font=f_th, fill=C_SUB)
    d.text((950, 928), "• เพิ่ม N จาก 1M เป็น 10M (x10 เท่า):", font=f_th, fill=C_TEXT)
    d.text((950, 964), "  เวลาเพิ่ม 1.22s -> 13.13s (x10.76 เชิงเส้น!)", font=f_th, fill=C_AMBER)

    out = ASSETS_DIR / "radix_bigo_diagram.png"
    img.resize((773, 550), Image.Resampling.LANCZOS).save(out)
    return out


# =============================================================================
# PowerPoint Helper Utilities
# =============================================================================
def set_slide_bg(slide, color_tuple=C_BG) -> None:
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color_tuple)


def add_rect(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: tuple[int, int, int],
    border_color: tuple[int, int, int] | None = None,
    border_pt: float = 1.0,
    rounded: bool = True,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    if rounded:
        shp.adjustments[0] = 0.035
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill_color)
    if border_color is not None:
        shp.line.color.rgb = rgb(border_color)
        shp.line.width = Pt(border_pt)
    else:
        shp.line.fill.background()
    return shp


def add_text(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    size: int = 16,
    color: tuple[int, int, int] = C_TEXT,
    bold: bool = False,
    font_name: str = "Leelawadee UI",
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
    line_spacing: float = 1.2,
):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)

    lines = text.split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
    return tb


def add_header_and_footer(
    slide,
    title: str,
    subtitle: str,
    slide_num: int,
    total_slides: int,
    time_range: str,
) -> None:
    add_rect(slide, 0, 0, 8.8, 0.07, C_GREEN, rounded=False)
    add_rect(slide, 8.8, 0, 4.533, 0.07, C_CYAN, rounded=False)

    add_text(slide, 0.55, 0.24, 10.5, 0.52, title, size=25, color=C_TEXT, bold=True)
    add_text(slide, 0.55, 0.76, 11.4, 0.36, subtitle, size=13, color=C_MUTED)

    add_rect(slide, 11.15, 0.30, 1.65, 0.40, C_CARD, C_BORDER, 1.0, rounded=True)
    add_text(
        slide, 11.15, 0.31, 1.65, 0.38,
        f"⏱ {time_range}", size=11, color=C_CYAN, bold=True, font_name="Consolas",
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE,
    )

    add_rect(slide, 0.55, 7.06, 12.23, 0.015, C_BORDER, rounded=False)
    add_text(
        slide, 0.55, 7.12, 8.5, 0.28,
        "Radix Sort (LSD — การจัดเรียงตามหลักตัวเลข)  •  อ้างอิง 15 Sorting Algorithms in 6 Minutes",
        size=10, color=C_MUTED,
    )
    add_text(
        slide, 11.4, 7.12, 1.38, 0.28,
        f"{slide_num} / {total_slides}", size=10, color=C_SUB, bold=True, font_name="Consolas",
        align=PP_ALIGN.RIGHT,
    )


def set_notes(slide, notes_text: str) -> None:
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = notes_text


def style_cell(
    cell,
    text: str,
    size: int = 12,
    color: tuple[int, int, int] = C_TEXT,
    bg: tuple[int, int, int] = C_CARD,
    bold: bool = False,
    font_name: str = "Leelawadee UI",
    align=PP_ALIGN.LEFT,
) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb(bg)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font_name
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = rgb(color)


# =============================================================================
# Build 8-Slide Presentation (~5:00 Minutes Total)
# =============================================================================
def build_presentation() -> Path:
    data = json.loads((BASE_DIR / "benchmark_results.json").read_text(encoding="utf-8"))

    img_step = create_radix_step_diagram()
    img_sim = create_radix_simulator_preview()
    img_bench = create_radix_benchmark_chart(data)
    img_bigo = create_radix_bigo_diagram(data)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    TOTAL = 8

    # =========================================================================
    # SLIDE 1: Title Slide (0:00 - 0:25)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1)
    add_rect(s1, 0, 0, 8.5, 0.10, C_GREEN, rounded=False)
    add_rect(s1, 8.5, 0, 4.833, 0.10, C_CYAN, rounded=False)

    # Decorative Radix Sort 4-Wave Bucket skyline on right side!
    wave_bars = [
        (0.65, C_CYAN), (1.05, C_CYAN), (1.45, C_CYAN), (1.85, C_CYAN),
        (1.15, C_INDIGO), (1.65, C_INDIGO), (2.20, C_INDIGO), (2.75, C_INDIGO),
        (1.80, C_PURPLE), (2.40, C_RED), (3.00, C_PURPLE), (3.55, C_PURPLE),
        (3.25, C_GREEN), (3.95, C_GREEN), (4.45, C_GREEN),
    ]
    for i, (bh, col) in enumerate(wave_bars):
        bx = 8.20 + i * 0.30
        by = 6.20 - bh
        add_rect(s1, bx, by, 0.21, bh, col, rounded=False)

    add_rect(s1, 0.75, 1.08, 4.95, 0.44, (12, 42, 36), C_GREEN, 1.2, rounded=True)
    add_text(
        s1, 0.75, 1.10, 4.95, 0.40,
        "NON-COMPARISON SORT  •  0 COMPARISONS  •  5 MIN",
        size=11, color=C_GREEN, bold=True, font_name="Consolas",
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE,
    )

    add_text(
        s1, 0.75, 1.78, 7.5, 1.35,
        "Radix Sort (LSD)\nการจัดเรียงตามหลักตัวเลข",
        size=42, color=C_TEXT, bold=True, line_spacing=1.15,
    )
    add_text(
        s1, 0.75, 3.55, 7.2, 1.00,
        "อัลกอริทึมจัดเรียงที่ “ไม่เคยเปรียบเทียบตัวเลขสองตัวเลยสักครั้ง”\nทะลุเพดานความเร็ว O(n log n) จำลองภาพ-เสียงแบบ Base 4 ตามวิดีโอ\nและทดลองจัดเรียงจริงถึง 10 ล้านตัวใน 13 วินาที!",
        size=15.5, color=C_SUB, line_spacing=1.35,
    )

    pills = [
        ("Time Complexity", "O(d(n + b))", C_GREEN),
        ("Space Complexity", "O(n + b)", C_CYAN),
        ("Comparisons", "0 ครั้ง!", C_AMBER),
    ]
    for idx, (lbl, val, col) in enumerate(pills):
        px = 0.75 + idx * 2.40
        add_rect(s1, px, 4.95, 2.22, 1.12, C_CARD, col, 1.4, rounded=True)
        add_text(s1, px + 0.14, 5.08, 1.94, 0.30, lbl, size=11, color=C_MUTED)
        add_text(s1, px + 0.14, 5.42, 1.94, 0.50, val, size=19, color=col, bold=True, font_name="Consolas")

    add_text(
        s1, 0.75, 6.62, 11.8, 0.35,
        "อ้างอิงวิดีโอตัวอย่าง: Timo Bingmann — 15 Sorting Algorithms in 6 Minutes (Algorithm #6: Radix Sort LSD)",
        size=11, color=C_MUTED,
    )
    set_notes(
        s1,
        "[เวลาแนะนำ: 0:00 – 0:25 | 25 วินาที]\n"
        "สวัสดีครับ วันนี้กลุ่มของผมขอนำเสนออัลกอริทึมจัดเรียงตัวเลขลำดับที่ 6 ในวิดีโอ "
        "'15 Sorting Algorithms in 6 Minutes' นั่นคือ 'Radix Sort (LSD)' หรือการจัดเรียงตามหลักตัวเลขจากขวาไปซ้ายครับ\n"
        "จุดเด่นที่น่าทึ่งที่สุดของอัลกอริทึมนี้คือ มันไม่เคยเอาตัวเลขสองตัวมาเปรียบเทียบกันว่าใครมากกว่าเลยสักครั้ง (0 Comparisons) "
        "ทำให้ทะลุขีดจำกัด O(n log n) ของ Comparison Sort และเรียงเลขสุ่ม 10 ล้านตัวด้วย Python ได้ภายใน 13 วินาทีครับ!",
    )

    # =========================================================================
    # SLIDE 2: Sorting Overview & Why Radix Sort (LSD) (0:25 - 1:05)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s2)
    add_header_and_footer(
        s2,
        "ภาพรวมอัลกอริทึมจัดเรียงตัวเลข & จุดเด่นของ Radix Sort (LSD)",
        "เทียบกลุ่มอัลกอริทึมในวิดีโอตัวอย่าง — ทำไม Radix Sort ถึงเร็วกว่า O(n log n) และต่างจาก Counting Sort?",
        2, TOTAL, "0:25–1:05",
    )

    # Left: Comparison Table
    t2_shape = s2.shapes.add_table(7, 5, Inches(0.55), Inches(1.28), Inches(7.70), Inches(5.40))
    t2 = t2_shape.table
    for idx, w in enumerate([2.15, 1.45, 1.45, 1.35, 1.30]):
        t2.columns[idx].width = Inches(w)

    for c, h in enumerate(["อัลกอริทึม", "Average", "Worst", "Comparisons", "Stable"]):
        style_cell(t2.cell(0, c), h, size=11.5, color=C_CYAN, bg=(17, 30, 54), bold=True, align=PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT)

    rows_cmp = [
        ("Bubble / Selection / Insertion", "O(n²)", "O(n²)", "O(n²)", "Yes/No", False),
        ("Shell Sort", "O(n¹.²⁵–n¹.⁵)", "O(n²)", "Comparison", "No", False),
        ("Quick Sort", "O(n log n)", "O(n²)", "O(n log n)", "No", False),
        ("Merge Sort / Heap Sort", "O(n log n)", "O(n log n)", "O(n log n)", "Yes / No", False),
        ("Counting Sort", "O(n + k)", "O(n + k)", "0 (จำกัด k เล็ก)", "Yes", False),
        ("Radix Sort (LSD) ★", "O(d(n + b))", "O(d(n + b))", "0 ครั้ง!", "Yes", True),
    ]
    for r_idx, (name, avg, worst, comp, stb, is_hl) in enumerate(rows_cmp, start=1):
        bg_c = (12, 48, 44) if is_hl else (C_CARD if r_idx % 2 == 1 else C_CARD2)
        tc = (255, 255, 255) if is_hl else C_SUB
        style_cell(t2.cell(r_idx, 0), name, size=12, color=C_GREEN if is_hl else C_TEXT, bg=bg_c, bold=is_hl)
        style_cell(t2.cell(r_idx, 1), avg, size=11, color=C_GREEN if is_hl else tc, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)
        w_col = C_GREEN if is_hl else (C_RED if "n²" in worst else tc)
        style_cell(t2.cell(r_idx, 2), worst, size=11, color=w_col, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)
        style_cell(t2.cell(r_idx, 3), comp, size=11, color=C_AMBER if is_hl else tc, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)
        style_cell(t2.cell(r_idx, 4), stb, size=11, color=C_CYAN if is_hl else tc, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)

    # Right: Why Radix Sort LSD (including clear contrast vs Counting Sort!)
    add_rect(s2, 8.50, 1.28, 4.28, 5.40, C_CARD, C_GREEN, 1.3, rounded=True)
    add_text(s2, 8.74, 1.46, 3.80, 0.42, "ทำไมต้อง Radix Sort (LSD)?", size=17.5, color=C_GREEN, bold=True)

    reasons = [
        ("1. ทะลุเพดาน O(n log n)", "Comparison Sort ทุกตัวติดขอบล่าง Ω(n log n) แต่ Radix Sort ไม่เทียบค่าเลย จึงทำเวลาเชิงเส้น O(n) ได้", C_GREEN),
        ("2. แก้จุดอ่อนของ Counting Sort", "Counting Sort ต้องจองช่องเท่าค่าสูงสุด k (ถ้าเลขถึง 1 พันล้าน = พัง!) แต่ Radix Sort แยกเรียงทีละหลัก ใช้ถังแค่ 10 หรือ 256 ช่อง", C_CYAN),
        ("3. เสถียรทุกกรณี + Stable", "Best = Average = Worst เท่ากันหมด ไม่มี Worst Case O(n²) และรักษาลำดับค่าซ้ำ 100%", C_INDIGO),
        ("4. ภาพและเสียงโดดเด่นในวิดีโอ", "ในวิดีโอใช้ฐาน 4 (Base 4) เกิดคลื่นการจัดกลุ่ม 4 ถังที่สวยงามและจดจำได้ทันที", C_AMBER),
    ]
    for idx, (head, desc, col) in enumerate(reasons):
        ry = 2.02 + idx * 1.14
        add_rect(s2, 8.74, ry, 0.08, 0.92, col, rounded=False)
        add_text(s2, 8.94, ry - 0.04, 3.64, 0.30, head, size=12.5, color=col, bold=True)
        add_text(s2, 8.94, ry + 0.27, 3.64, 0.72, desc, size=11, color=C_SUB, line_spacing=1.18)

    set_notes(
        s2,
        "[เวลาแนะนำ: 0:25 – 1:05 | 40 วินาที]\n"
        "จากตารางเปรียบเทียบจะเห็นว่าอัลกอริทึมส่วนใหญ่ตั้งแต่ Bubble, Selection, Insertion, Shell ไปจนถึง Quick, Merge และ Heap Sort "
        "ล้วนเป็น Comparison Sort ซึ่งทางทฤษฎีพิสูจน์แล้วว่าเร็วสุดได้ไม่เกิน O(n log n)\n"
        "ส่วน Counting Sort แม้จะเร็วแบบ O(n+k) แต่ถ้าค่าตัวเลขกว้างถึง 1,000 ล้าน จะต้องจองหน่วยความจำมหาศาลจนใช้งานไม่ได้\n"
        "Radix Sort (LSD) จึงแก้ปัญหานี้ด้วยการแยกเรียง 'ทีละหลัก' จากขวาไปซ้าย ใช้ถังแค่ 10 ถัง (ฐาน 10) หรือ 256 ถัง (ฐาน 256) "
        "แต่เรียงตัวเลขหลักพันล้านได้ในเวลาเชิงเส้น O(d(n+b)) โดยไม่ต้องเปรียบเทียบตัวเลขเลยครับ",
    )

    # =========================================================================
    # SLIDE 3: How Radix Sort (LSD) Works — Step-by-Step Visual (1:05 - 1:55)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s3)
    add_header_and_footer(
        s3,
        "หลักการทำงานของ Radix Sort (LSD): เรียงทีละหลักจากขวาไปซ้าย",
        "ตัวอย่างอาร์เรย์ [170, 45, 75, 90, 802, 24, 2, 66] — แบ่งใส่ถัง 0–9 ทีละหลัก (หน่วย → สิบ → ร้อย)",
        3, TOTAL, "1:05–1:55",
    )

    s3.shapes.add_picture(str(img_step), Inches(0.55), Inches(1.22), width=Inches(12.23), height=Inches(5.72))

    set_notes(
        s3,
        "[เวลาแนะนำ: 1:05 – 1:55 | 50 วินาที]\n"
        "มาดูการทำงานจริงทีละขั้นกับตัวเลข 8 ตัวคือ [170, 45, 75, 90, 802, 24, 2, 66] ครับ\n"
        "• รอบที่ 1 (สีฟ้า — หลักหน่วย): ดูแค่เลขหลักสุดท้ายแล้วหย่อนลงถัง 0–9 จากนั้นหยิบออกจากถัง 0 ไล่ไปถัง 9 จะได้ 170, 090, 802, 002, 024, 045, 075, 066\n"
        "• รอบที่ 2 (สีส้ม — หลักสิบ): ดูเลขหลักตรงกลาง หย่อนลงถังอีกครั้ง สังเกตจุดสำคัญมากคือ 170 กับ 075 มีหลักสิบเป็นเลข 7 เหมือนกัน แต่ 170 เข้าถังก่อนจึงออกก่อน 075 เพราะเราเรียงแบบ Stable!\n"
        "• รอบที่ 3 (สีเขียว — หลักร้อย): ดูเลขซ้ายสุด แล้วดึงออกจากถัง 0, 1, 8 จะได้ผลลัพธ์ [2, 24, 45, 66, 75, 90, 170, 802] เรียงสมบูรณ์ครบทุกตัวโดยไม่มีการเปรียบเทียบกันเลยแม้แต่ครั้งเดียวครับ!",
    )

    # =========================================================================
    # SLIDE 4: Deep Dive — Bucket Code & Why Stability Matters (1:55 - 2:35)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s4)
    add_header_and_footer(
        s4,
        "เจาะลึกโค้ด Python (Base 10 & Base 256) และหัวใจของความ Stable",
        "โค้ดสั้นกระชับเพียง ~12 บรรทัด — รองรับทั้งฐานสิบ (อ่านง่าย) และฐาน 256 Bitwise (เร็วสูงสุด)",
        4, TOTAL, "1:55–2:35",
    )

    # Left: Code Card
    add_rect(s4, 0.55, 1.28, 6.85, 5.50, (9, 14, 26), C_BORDER, 1.3, rounded=True)
    add_rect(s4, 0.55, 1.28, 6.85, 0.44, (18, 28, 48), C_BORDER, 1.0, rounded=False)
    add_text(s4, 0.75, 1.34, 6.4, 0.34, "radix_sort_lsd.py  (Pure Python — 0 Comparisons)", size=12, color=C_GREEN, bold=True, font_name="Consolas")

    code_text = (
        "# แบบที่ 1: ฐาน 10 (เข้าใจง่าย — เรียงทีละหลัก หน่วย/สิบ/ร้อย)\n"
        "def radix_sort_base10(arr):\n"
        "    max_val, exp = max(arr), 1\n"
        "    while max_val // exp > 0:          # วน d รอบ (จำนวนหลัก)\n"
        "        buckets = [[] for _ in range(10)]\n"
        "        for x in arr:                  # 1. แยกใส่ถัง 0..9\n"
        "            buckets[(x // exp) % 10].append(x)\n"
        "        arr = [x for b in buckets for x in b]  # 2. รวบรวมคืน\n"
        "        exp *= 10\n"
        "    return arr\n"
        "\n"
        "# แบบที่ 2: ฐาน 256 (ใช้งานจริง — ใช้ Bit-Shift ทีละ 8 บิต!)\n"
        "def radix_sort_base256(arr):\n"
        "    max_val, shift = max(arr), 0\n"
        "    while (max_val >> shift) > 0:      # เลข 32-bit ใช้แค่ 4 รอบ!\n"
        "        buckets = [[] for _ in range(256)]\n"
        "        for x in arr:\n"
        "            buckets[(x >> shift) & 255].append(x)\n"
        "        arr = [x for b in buckets for x in b]\n"
        "        shift += 8\n"
        "    return arr"
    )
    add_text(s4, 0.75, 1.82, 6.50, 4.82, code_text, size=11.0, color=(226, 232, 240), font_name="Consolas", line_spacing=1.18)

    # Right: Key Engineering Concepts
    add_rect(s4, 7.62, 1.28, 5.16, 5.50, C_CARD, C_BORDER, 1.3, rounded=True)
    add_text(s4, 7.86, 1.46, 4.70, 0.40, "กลไกสำคัญที่ทำให้ Radix Sort ถูกต้องและเร็ว", size=16.5, color=C_CYAN, bold=True)

    cards_s4 = [
        ("1. กฎเหล็ก: ต้องเป็น Stable ในทุกรอบ!", "เมื่อใส่ค่าลงถังด้วย .append(x) (FIFO) ค่าที่เรียงถูกมาจากหลักขวาแล้วจะไม่สลับที่กันเมื่อหลักซ้ายมีค่าเท่ากัน (เช่น 045 กับ 075)", C_GREEN),
        ("2. สูตรดึงค่าหลักแบบ O(1)", "• ฐาน 10: digit = (x // exp) % 10\n• ฐาน 256: digit = (x >> shift) & 255  (เร็วกว่าหารและ mod มาก!)", C_CYAN),
        ("3. เลือกฐาน b = 256 ลดรอบจาก 9 → 4 รอบ", "เลข 0–999,999,999 (30 บิต) ใช้ฐาน 10 ต้องกวาด 9 รอบ แต่ใช้ฐาน 256 กวาดทีละ 8 บิต เหลือแค่ 4 รอบ!", C_AMBER),
        ("4. รองรับเลขติดลบได้อย่างไร?", "แยกกลุ่มเลขลบกับเลขบวก หรือแปลง Sign Bit (XOR 0x80000000) ก่อนเรียงแล้วแปลงกลับ", C_INDIGO),
    ]
    for idx, (ch, cb, col) in enumerate(cards_s4):
        sy = 2.00 + idx * 1.15
        add_rect(s4, 7.86, sy, 4.68, 1.02, C_CARD2, col, 1.0, rounded=True)
        add_text(s4, 8.02, sy + 0.08, 4.38, 0.30, ch, size=12, color=col, bold=True)
        add_text(s4, 8.02, sy + 0.38, 4.38, 0.58, cb, size=10.8, color=C_SUB, line_spacing=1.16)

    set_notes(
        s4,
        "[เวลาแนะนำ: 1:55 – 2:35 | 40 วินาที]\n"
        "ดูโค้ด Python จะเห็นว่าสั้นและสะอาดมากครับ มีแค่ 2 ขั้นตอนในแต่ละรอบคือ แยกใส่ถัง (Distribute) แล้วรวบรวมคืน (Gather)\n"
        "หัวใจสำคัญที่สุดคือ 'ความ Stable' การใช้ list.append เป็นคิวแบบเข้าก่อนออกก่อน ทำให้ลำดับที่เรียงถูกต้องจากหลักขวายังคงอยู่ครบเมื่อย้ายไปดูหลักซ้าย\n"
        "และในงานจริง แทนที่จะหารเอาเศษด้วยฐาน 10 (ต้องทำถึง 9 รอบสำหรับเลข 9 หลัก) เราสามารถใช้ฐาน 256 ด้วยตัวดำเนินการ Bit-Shift `(x >> shift) & 255` "
        "ซึ่งประมวลผลทีละ 1 ไบต์ (8 บิต) ลดจำนวนรอบจาก 9 รอบเหลือแค่ 4 รอบ เร็วขึ้นกว่า 2 เท่าทันทีครับ!",
    )

    # =========================================================================
    # SLIDE 5: Visual & Audio Simulator Demo (2:35 - 3:10)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s5)
    add_header_and_footer(
        s5,
        "โปรแกรมจำลองการจัดเรียง (Simulation ตามวิดีโอตัวอย่าง)",
        "ถอดแบบ Algorithm #6: Radix Sort (LSD) Base 4 จากวิดีโอ พร้อมแถบถังสดและเสียง Web Audio API",
        5, TOTAL, "2:35–3:10",
    )

    s5.shapes.add_picture(str(img_sim), Inches(0.55), Inches(1.28), width=Inches(8.25), height=Inches(4.86))

    sim_cards = [
        ("🎬 ตรงตามวิดีโอต้นฉบับ (Base 4)", "Timo Bingmann ใช้ RADIX = 4 (ทีละ 2 บิต) ทำให้เห็นคลื่น 4 ทางลาดซ้อนกันค่อย ๆ รวมตัวเป็นเส้นตรงในรอบสุดท้าย พร้อม Comparisons = 0!", C_CYAN),
        ("🔊 เสียงสังเคราะห์ Web Audio API", "สร้างคลื่น Triangle Wave ตามความสูงแท่ง (140–1,250 Hz) ได้ยินจังหวะ Scan (สีส้ม) สลับกับจังหวะ Gather จากถัง 0→3 (สีแดง)", C_AMBER),
        ("🎛 โหมดสาธิต & ไฟล์โปรแกรม", "• สลับฐานได้ทันที: Base 4 / 10 / 2 / 8\n• sound_of_sorting_radix_web.html (มีเสียงสด)\n• radix_sort_visualizer.py (Python Tkinter)", C_GREEN),
    ]
    for idx, (ch, cb, col) in enumerate(sim_cards):
        cy = 1.28 + idx * 1.70
        add_rect(s5, 9.02, cy, 3.76, 1.54, C_CARD, col, 1.2, rounded=True)
        add_text(s5, 9.20, cy + 0.12, 3.42, 0.34, ch, size=13, color=col, bold=True)
        add_text(s5, 9.20, cy + 0.48, 3.42, 0.98, cb, size=11.2, color=C_SUB, line_spacing=1.20)

    add_rect(s5, 0.55, 6.26, 8.25, 0.56, C_CARD, C_BORDER, 1.0, rounded=True)
    add_text(
        s5, 0.72, 6.34, 7.90, 0.40,
        "💡 สังเกตเอกลักษณ์ของ Radix Sort: ระหว่างจัดเรียงจะมีคลื่นย่อย 4 ช่วง (ถัง 0–3) แล้วรอบสุดท้ายจะ 'ดีด' เรียงสมบูรณ์พร้อมกันทั้งกราฟ!",
        size=11.2, color=C_SUB, valign=MSO_ANCHOR.MIDDLE,
    )

    set_notes(
        s5,
        "[เวลาแนะนำ: 2:35 – 3:10 | 35 วินาที]\n"
        "โปรแกรมจำลองของเราสร้างไว้ทั้งเวอร์ชันเว็บ HTML5 ที่มีเสียงสังเคราะห์สด และเวอร์ชัน Python Tkinter ครับ\n"
        "จุดที่ตั้งใจทำให้เหมือนวิดีโอ '15 Sorting Algorithms in 6 Minutes' เป๊ะ ๆ คือตั้งค่าเริ่มต้นเป็น Base 4 เหมือนโค้ดต้นฉบับของ Timo Bingmann "
        "ในแต่ละรอบจะมี 2 จังหวะคือ สแกนแยกเข้าถัง 0–3 (แท่งสีส้ม) และดึงคืนจากถังทีละใบ (แท่งสีแดง) จะเห็นตัวเลขด้านล่างแสดงจำนวนของในแต่ละถังแบบเรียลไทม์ "
        "และค่า Comparisons จะเป็น 0 ตลอดการรันครับ\n"
        "(สามารถสลับหน้าจอไปเปิดไฟล์ sound_of_sorting_radix_web.html โชว์สด 10–15 วินาทีตรงนี้ได้เลยครับ!)",
    )

    # =========================================================================
    # SLIDE 6: Benchmark Results (10K, 100K, 1M, 10M) (3:10 - 3:55)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s6)
    add_header_and_footer(
        s6,
        "ผลการทดลองวัดเวลา (Benchmark Analysis)",
        "สุ่มตัวเลข 9 หลัก (0–999,999,999) บน Pure Python 3.12 / AMD Ryzen 5 220 / RAM 16 GB — ตรวจสอบความถูกต้องครบทุกรอบ",
        6, TOTAL, "3:10–3:55",
    )

    t6_shape = s6.shapes.add_table(5, 6, Inches(0.55), Inches(1.22), Inches(12.23), Inches(2.12))
    t6 = t6_shape.table
    for idx, w in enumerate([1.95, 2.10, 2.05, 1.95, 2.08, 2.10]):
        t6.columns[idx].width = Inches(w)

    h6 = ["จำนวนตัวเลขสุ่ม (N)", "Radix Base 256 (4 รอบ)", "Radix Base 10 (9 รอบ)", "เทียบ Merge Sort", "อัตราโต (Base 256)", "Comparisons / Writes"]
    for c, h in enumerate(h6):
        style_cell(t6.cell(0, c), h, size=11, color=C_CYAN, bg=(17, 30, 54), bold=True, align=PP_ALIGN.CENTER)

    b256_rows = data["scaling_base256"]
    b10_rows = data["scaling_base10"]
    ref_m = data["reference_merge_sort"]
    prev_t = None

    for idx, (r256, r10) in enumerate(zip(b256_rows, b10_rows), start=1):
        n = r256["n"]
        t256 = r256["avg_sec"]
        t10 = r10["avg_sec"]
        ratio_txt = "—" if prev_t is None else f"×{t256 / prev_t:.2f} (N ×10)"
        prev_t = t256

        bg_c = (12, 46, 42) if n == 1_000_000 else (C_CARD if idx % 2 == 1 else C_CARD2)
        f256 = f"{t256*1000:.2f} ms" if t256 < 0.1 else f"{t256:.4f} s"
        f10 = f"{t10*1000:.2f} ms" if t10 < 0.1 else f"{t10:.4f} s"
        fm = f"{ref_m[str(n)]:.4f} s" if str(n) in ref_m else "— (ช้ากว่า ~4×)"
        cw = f"0 / {r256['writes']:,}"

        style_cell(t6.cell(idx, 0), f"{n:,}", size=12, color=C_TEXT, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.RIGHT)
        style_cell(t6.cell(idx, 1), f256, size=12, color=C_GREEN, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.CENTER)
        style_cell(t6.cell(idx, 2), f10, size=11.5, color=C_CYAN, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.CENTER)
        style_cell(t6.cell(idx, 3), fm, size=11, color=C_INDIGO, bg=bg_c, font_name="Consolas", align=PP_ALIGN.CENTER)
        style_cell(t6.cell(idx, 4), ratio_txt, size=11, color=C_AMBER, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.CENTER)
        style_cell(t6.cell(idx, 5), cw, size=10.5, color=C_SUB, bg=bg_c, font_name="Consolas", align=PP_ALIGN.RIGHT)

    s6.shapes.add_picture(str(img_bench), Inches(0.55), Inches(3.46), width=Inches(12.23), height=Inches(3.48))

    set_notes(
        s6,
        "[เวลาแนะนำ: 3:10 – 3:55 | 45 วินาที]\n"
        "มาดูผลการทดลองจับเวลาจริงด้วย Pure Python บนเครื่องจริงกันครับ เราทดสอบครบทั้ง 10,000 / 100,000 / 1,000,000 และ 10,000,000 ตัว:\n"
        "• 10,000 ตัว: Base 256 ใช้เวลาแค่ 5.33 มิลลิวินาที! เร็วกว่า Selection Sort เกือบ 500 เท่า!\n"
        "• 100,000 ตัว: ใช้เวลาเพียง 0.062 วินาที (62 มิลลิวินาที)\n"
        "• 1,000,000 ตัว: Base 256 จบใน 1.22 วินาที (และ Base 10 จบใน 2.73 วินาที) เร็วกว่า Merge Sort O(n log n) ถึง 3.26 เท่า!\n"
        "• 10,000,000 ตัว: Base 256 เรียงเลข 10 ล้านตัวจบในแค่ 13.13 วินาที! สังเกตว่าพอเพิ่มข้อมูล 10 เท่าจาก 1M เป็น 10M เวลาเพิ่มแค่ 10.76 เท่า เป็นเส้นตรงชัดเจน!\n"
        "ส่วนกราฟด้านขวาแสดงให้เห็นว่ายิ่งเพิ่มฐานจาก Base 4 → 10 → 16 → 256 จำนวนรอบลดลงจาก 15 เหลือ 4 รอบ เวลาจึงลดจาก 4.50 วิ เหลือ 1.13 วิครับ",
    )

    # =========================================================================
    # SLIDE 7: Runtime Big-O Mathematical Analysis (3:55 - 4:35)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s7)
    add_header_and_footer(
        s7,
        "การวิเคราะห์ Runtime Big-O (Time & Space Complexity)",
        "ที่มาของเวลาเชิงเส้น O(d(n + b)) และหลักฐานยืนยันจากการทดลองจริงทั้งมิติของ N และจำนวนรอบ d",
        7, TOTAL, "3:55–4:35",
    )

    # Left: Time & Space Complexity Cards
    add_rect(s7, 0.55, 1.28, 4.30, 2.86, C_CARD, C_GREEN, 1.4, rounded=True)
    add_text(s7, 0.78, 1.42, 3.85, 0.36, "⏱ Time Complexity (เวลาทำงาน)", size=15, color=C_GREEN, bold=True)
    add_text(
        s7, 0.78, 1.84, 3.85, 0.46,
        "Best / Avg / Worst = Θ(d(n + b))",
        size=14.5, color=C_CYAN, bold=True, font_name="Consolas",
    )
    add_text(
        s7, 0.78, 2.34, 3.85, 1.66,
        "• n = จำนวนข้อมูลทั้งหมด\n"
        "• b = ฐานเลข / จำนวนถัง (เช่น 10 หรือ 256)\n"
        "• d = จำนวนหลัก = ⌈log_b(max + 1)⌉\n"
        "• เมื่อค่าคงที่ (เช่น 32-bit int, b=256 → d=4)\n"
        "  จะได้เวลาทำงานเชิงเส้น O(n)!",
        size=11.8, color=C_SUB, line_spacing=1.24,
    )

    add_rect(s7, 0.55, 4.30, 4.30, 2.48, C_CARD, C_INDIGO, 1.4, rounded=True)
    add_text(s7, 0.78, 4.44, 3.85, 0.36, "💾 Space Complexity (หน่วยความจำ)", size=15, color=C_INDIGO, bold=True)
    add_text(
        s7, 0.78, 4.86, 3.85, 0.46,
        "Auxiliary Space = O(n + b)",
        size=16, color=C_AMBER, bold=True, font_name="Consolas",
    )
    add_text(
        s7, 0.78, 5.36, 3.85, 1.32,
        "• ถังเก็บรอแยกข้อมูล (buckets) รวม n ตัว = O(n)\n"
        "• โครงสร้างถัง/ตารางนับความถี่ b ช่อง = O(b)\n"
        "• ไม่อาศัย Recursion → Stack Space = O(1)",
        size=11.8, color=C_SUB, line_spacing=1.24,
    )

    s7.shapes.add_picture(str(img_bigo), Inches(5.05), Inches(1.28), width=Inches(7.73), height=Inches(5.50))

    set_notes(
        s7,
        "[เวลาแนะนำ: 3:55 – 4:35 | 40 วินาที]\n"
        "มาดูการวิเคราะห์ Runtime Big-O กันครับ:\n"
        "1. Time Complexity คือ O(d × (n + b)) เท่ากันทุกกรณี ทั้ง Best, Average และ Worst Case! "
        "เพราะใน 1 รอบเราต้องสแกนข้อมูล n ตัวลงถัง b ใบแล้วดึงคืน n ตัว ใช้เวลา O(n + b) และทำซ้ำทั้งหมด d รอบตามจำนวนหลัก\n"
        "2. จุดสำคัญคือสำหรับเลขจำนวนเต็ม 32-bit เมื่อเลือกฐาน b = 256 จำนวนรอบ d จะคงที่แค่ 4 รอบเสมอ ดังนั้น O(4(n + 256)) จึงลดรูปเป็น O(n) เชิงเส้น!\n"
        "3. กราฟด้านขวาล่างยืนยันจากผลรันจริงที่ 1 ล้านตัว: ไม่ว่าจะใช้ Base 4, 10, 16 หรือ 256 เวลาเฉลี่ยต่อ 1 รอบคงที่ประมาณ 0.28–0.30 วินาทีเป๊ะ "
        "และ Space Complexity อยู่ที่ O(n + b) โดยไม่ต้องใช้ Recursion Stack ครับ",
    )

    # =========================================================================
    # SLIDE 8: Summary, Trade-offs & Real-World Uses (4:35 - 5:00)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s8)
    add_header_and_footer(
        s8,
        "สรุป ข้อดี-ข้อจำกัด และการประยุกต์ใช้งานจริง",
        "บทสรุปของ Radix Sort (LSD) — อัลกอริทึมความเร็วเชิงเส้นที่อยู่เบื้องหลังงานประมวลผลข้อมูลขนาดใหญ่",
        8, TOTAL, "4:35–5:00",
    )

    cols_s8 = [
        (
            "✅ จุดเด่น (Strengths)",
            C_GREEN,
            [
                ("เร็วระดับเชิงเส้น O(d(n+b))", "เรียงตัวเลข 1 ล้านตัวใน 1.22 วินาที (เร็วกว่า Merge Sort 3.26× และเร็วกว่า O(n²) กว่า 20,000×)"),
                ("0 Comparisons & Stable", "ไม่เปรียบเทียบค่าเลย ไม่มีจุดอ่อน Worst Case จากข้อมูลเรียงกลับด้าน และรักษาลำดับค่าซ้ำ"),
                ("เหมาะกับ Parallel & GPU", "โครงสร้างแยกถังทำงานขนานได้ยอดเยี่ยม เป็นอัลกอริทึมเรียงลำดับที่เร็วที่สุดบน GPU"),
            ],
        ),
        (
            "⚠️ ข้อจำกัด (Trade-offs)",
            C_AMBER,
            [
                ("จำกัดชนิดของคีย์ข้อมูล", "ใช้ได้ดีกับจำนวนเต็ม, สตริงความยาวคงที่ หรือคีย์ที่แยกเป็นหลัก/ไบต์ได้ ไม่ยืดหยุ่นเท่า Comparison Sort"),
                ("ใช้หน่วยความจำเพิ่ม O(n+b)", "ไม่เป็น In-place Sort ต้องมีพื้นที่สำหรับถังหรืออาร์เรย์ผลลัพธ์เพิ่มอีกเท่าตัว"),
                ("ช้าลงถ้าจำนวนหลัก d สูงมาก", "หากคีย์ยาวมาก (เช่น สตริงยาวหลายร้อยตัวอักษร) ค่า d จะสูงจนอาจแพ้ Quick/Merge Sort"),
            ],
        ),
        (
            "🚀 การใช้งานในโลกจริง",
            C_CYAN,
            [
                ("NVIDIA CUDA / GPU Sorting", "ไลบรารี Thrust & CUB ใช้ Radix Sort เป็นตัวหลักในการเรียงข้อมูลหลายร้อยล้านแถวบน GPU"),
                ("ฐานข้อมูล, IP & วันที่ (Multi-Key)", "เรียงหมายเลขบัตร, IPv4, Timestamp หรือรหัสไปรษณีย์ได้อย่างรวดเร็วและเสถียร"),
                ("Bioinformatics & Suffix Array", "ใช้สร้าง Suffix Array (อัลกอริทึม DC3) สำหรับค้นหาลำดับ DNA และงาน Text Indexing"),
            ],
        ),
    ]

    for idx, (heading, col, items) in enumerate(cols_s8):
        cx = 0.55 + idx * 4.16
        add_rect(s8, cx, 1.32, 3.92, 4.62, C_CARD, col, 1.3, rounded=True)
        add_text(s8, cx + 0.22, 1.50, 3.48, 0.42, heading, size=16.5, color=col, bold=True)
        for j, (ih, ib) in enumerate(items):
            iy = 2.08 + j * 1.24
            add_rect(s8, cx + 0.22, iy, 3.48, 1.08, C_CARD2, C_BORDER, 0.8, rounded=True)
            add_text(s8, cx + 0.36, iy + 0.08, 3.22, 0.30, ih, size=12, color=C_TEXT, bold=True)
            add_text(s8, cx + 0.36, iy + 0.38, 3.22, 0.66, ib, size=10.8, color=C_SUB, line_spacing=1.16)

    # Bottom Closing Banner
    add_rect(s8, 0.55, 6.14, 12.23, 0.76, (12, 42, 38), C_GREEN, 1.2, rounded=True)
    add_text(
        s8, 0.78, 6.24, 8.45, 0.54,
        "สรุป: เมื่อข้อมูลเป็นตัวเลขหรือคีย์ที่มีหลักชัดเจน Radix Sort (LSD) ชนะ O(n log n) ได้จริง — ขอบคุณครับ / Q&A",
        size=13.0, color=C_TEXT, bold=True, valign=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        s8, 9.30, 6.24, 3.25, 0.54,
        "1M = 1.22s  |  10M = 13.13s!",
        size=13.5, color=C_GREEN, bold=True, font_name="Consolas",
        align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE,
    )

    set_notes(
        s8,
        "[เวลาแนะนำ: 4:35 – 5:00 | 25 วินาที]\n"
        "โดยสรุปครับ Radix Sort (LSD) เป็นอัลกอริทึมที่แสดงให้เห็นว่าเราไม่จำเป็นต้องเปรียบเทียบตัวเลขสองตัวเสมอไป "
        "ด้วยเวลา O(d(n+b)) ที่เสถียรทุกกรณีและเป็น Stable Sort ทำให้มันเร็วกว่าแม้กระทั่ง Merge Sort ถึงกว่า 3 เท่า "
        "และถูกใช้งานจริงอย่างแพร่หลายในระบบ GPU, ฐานข้อมูล และงานชีวสารสนเทศครับ\n"
        "ขอบคุณอาจารย์และเพื่อน ๆ ที่รับฟังครับ มีคำถามเพิ่มเติมสอบถามได้เลยครับ!",
    )

    out_pptx = BASE_DIR / "RadixSort_Presentation.pptx"
    prs.save(str(out_pptx))
    print(f"Presentation generated successfully: {out_pptx}")
    return out_pptx


if __name__ == "__main__":
    build_presentation()
