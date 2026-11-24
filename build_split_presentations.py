"""
Split Radix Sort (LSD) 8-Slide Presentation into 4 Independent PPTX Decks
=========================================================================
Generates 4 clean, standalone PowerPoint files (2 slides each) inside
`slides_4_parts/` so 4 team members can present distinct parts smoothly,
plus `Speaker_Script_4_People.md` with exact timing and handoff lines.
"""

from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

import build_radix_presentation as brp

BASE_DIR = Path(__file__).parent
PARTS_DIR = BASE_DIR / "slides_4_parts"
PARTS_DIR.mkdir(exist_ok=True)


def _new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def _add_part_footer(
    slide,
    part_label: str,
    title: str,
    subtitle: str,
    slide_idx: int,
    total_in_part: int,
    time_range: str,
) -> None:
    brp.add_rect(slide, 0, 0, 8.8, 0.07, brp.C_GREEN, rounded=False)
    brp.add_rect(slide, 8.8, 0, 4.533, 0.07, brp.C_CYAN, rounded=False)

    brp.add_text(slide, 0.55, 0.24, 10.4, 0.52, title, size=25, color=brp.C_TEXT, bold=True)
    brp.add_text(slide, 0.55, 0.76, 11.4, 0.36, subtitle, size=13, color=brp.C_MUTED)

    brp.add_rect(slide, 11.00, 0.30, 1.80, 0.40, brp.C_CARD, brp.C_BORDER, 1.0, rounded=True)
    brp.add_text(
        slide, 11.00, 0.31, 1.80, 0.38,
        f"⏱ {time_range}", size=11, color=brp.C_CYAN, bold=True, font_name="Consolas",
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE,
    )

    brp.add_rect(slide, 0.55, 7.06, 12.23, 0.015, brp.C_BORDER, rounded=False)
    brp.add_text(
        slide, 0.55, 7.12, 9.2, 0.28,
        f"Radix Sort (LSD)  •  {part_label}",
        size=10, color=brp.C_MUTED,
    )
    brp.add_text(
        slide, 11.4, 7.12, 1.38, 0.28,
        f"{slide_idx} / {total_in_part}", size=10, color=brp.C_SUB, bold=True, font_name="Consolas",
        align=PP_ALIGN.RIGHT,
    )


def build_part_1() -> Path:
    """Part 1 (คนที่ 1): บทนำ ภาพรวมอัลกอริทึม และเหตุผลที่เลือก Radix Sort (LSD) [0:00–1:05]"""
    prs = _new_prs()
    blank = prs.slide_layouts[6]

    # Slide 1: Cover
    s1 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s1)
    brp.add_rect(s1, 0, 0, 8.5, 0.10, brp.C_GREEN, rounded=False)
    brp.add_rect(s1, 8.5, 0, 4.833, 0.10, brp.C_CYAN, rounded=False)

    wave_bars = [
        (0.65, brp.C_CYAN), (1.05, brp.C_CYAN), (1.45, brp.C_CYAN), (1.85, brp.C_CYAN),
        (1.15, brp.C_INDIGO), (1.65, brp.C_INDIGO), (2.20, brp.C_INDIGO), (2.75, brp.C_INDIGO),
        (1.80, brp.C_PURPLE), (2.40, brp.C_RED), (3.00, brp.C_PURPLE), (3.55, brp.C_PURPLE),
        (3.25, brp.C_GREEN), (3.95, brp.C_GREEN), (4.45, brp.C_GREEN),
    ]
    for i, (bh, col) in enumerate(wave_bars):
        brp.add_rect(s1, 8.20 + i * 0.30, 6.20 - bh, 0.21, bh, col, rounded=False)

    brp.add_rect(s1, 0.75, 1.08, 5.30, 0.44, (12, 42, 36), brp.C_GREEN, 1.2, rounded=True)
    brp.add_text(
        s1, 0.75, 1.10, 5.30, 0.40,
        "PART 1 / 4 (คนที่ 1)  •  INTRO & OVERVIEW  •  0:00–1:05",
        size=10.5, color=brp.C_GREEN, bold=True, font_name="Consolas",
        align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE,
    )
    brp.add_text(
        s1, 0.75, 1.78, 7.5, 1.35,
        "Radix Sort (LSD)\nการจัดเรียงตามหลักตัวเลข",
        size=42, color=brp.C_TEXT, bold=True, line_spacing=1.15,
    )
    brp.add_text(
        s1, 0.75, 3.55, 7.2, 1.00,
        "อัลกอริทึมจัดเรียงที่ “ไม่เคยเปรียบเทียบตัวเลขสองตัวเลยสักครั้ง”\nทะลุเพดานความเร็ว O(n log n) จำลองภาพ-เสียงแบบ Base 4 ตามวิดีโอ\nและทดลองจัดเรียงจริงถึง 10 ล้านตัวใน 13 วินาที!",
        size=15.5, color=brp.C_SUB, line_spacing=1.35,
    )
    for idx, (lbl, val, col) in enumerate([
        ("Time Complexity", "O(d(n + b))", brp.C_GREEN),
        ("Space Complexity", "O(n + b)", brp.C_CYAN),
        ("Comparisons", "0 ครั้ง!", brp.C_AMBER),
    ]):
        px = 0.75 + idx * 2.40
        brp.add_rect(s1, px, 4.95, 2.22, 1.12, brp.C_CARD, col, 1.4, rounded=True)
        brp.add_text(s1, px + 0.14, 5.08, 1.94, 0.30, lbl, size=11, color=brp.C_MUTED)
        brp.add_text(s1, px + 0.14, 5.42, 1.94, 0.50, val, size=19, color=col, bold=True, font_name="Consolas")

    brp.add_text(
        s1, 0.75, 6.62, 11.8, 0.35,
        "อ้างอิงวิดีโอตัวอย่าง: Timo Bingmann — 15 Sorting Algorithms in 6 Minutes (Algorithm #6: Radix Sort LSD)",
        size=11, color=brp.C_MUTED,
    )
    brp.set_notes(
        s1,
        "[คนที่ 1 — สไลด์ 1/2 | เวลา 0:00–0:25 (25 วินาที)]\n"
        "สวัสดีอาจารย์และเพื่อน ๆ ครับ วันนี้กลุ่มของพวกเราขอนำเสนออัลกอริทึมจัดเรียงตัวเลขลำดับที่ 6 "
        "จากวิดีโอ '15 Sorting Algorithms in 6 Minutes' นั่นคือ 'Radix Sort (LSD)' หรือการจัดเรียงตามหลักตัวเลขครับ\n"
        "จุดที่น่าทึ่งที่สุดคือมันไม่เคยเปรียบเทียบตัวเลขสองตัวว่าใครมากกว่าเลยสักครั้ง (0 Comparisons) "
        "ทำให้ทะลุขีดจำกัด O(n log n) ได้ครับ",
    )

    # Slide 2: Overview & Why Radix Sort
    s2 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s2)
    _add_part_footer(
        s2, "ส่วนที่ 1/4 (คนที่ 1): บทนำและเหตุผลที่เลือก",
        "ภาพรวมอัลกอริทึมจัดเรียงตัวเลข & จุดเด่นของ Radix Sort (LSD)",
        "เทียบกลุ่มอัลกอริทึมในวิดีโอตัวอย่าง — ทำไม Radix Sort ถึงเร็วกว่า O(n log n) และต่างจาก Counting Sort?",
        2, 2, "0:25–1:05",
    )
    t2 = s2.shapes.add_table(7, 5, Inches(0.55), Inches(1.28), Inches(7.70), Inches(5.40)).table
    for idx, w in enumerate([2.15, 1.45, 1.45, 1.35, 1.30]):
        t2.columns[idx].width = Inches(w)
    for c, h in enumerate(["อัลกอริทึม", "Average", "Worst", "Comparisons", "Stable"]):
        brp.style_cell(t2.cell(0, c), h, size=11.5, color=brp.C_CYAN, bg=(17, 30, 54), bold=True, align=PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT)
    for r_idx, (name, avg, worst, comp, stb, is_hl) in enumerate([
        ("Bubble / Selection / Insertion", "O(n²)", "O(n²)", "O(n²)", "Yes/No", False),
        ("Shell Sort", "O(n¹.²⁵–n¹.⁵)", "O(n²)", "Comparison", "No", False),
        ("Quick Sort", "O(n log n)", "O(n²)", "O(n log n)", "No", False),
        ("Merge Sort / Heap Sort", "O(n log n)", "O(n log n)", "O(n log n)", "Yes / No", False),
        ("Counting Sort", "O(n + k)", "O(n + k)", "0 (จำกัด k เล็ก)", "Yes", False),
        ("Radix Sort (LSD) ★", "O(d(n + b))", "O(d(n + b))", "0 ครั้ง!", "Yes", True),
    ], start=1):
        bg_c = (12, 48, 44) if is_hl else (brp.C_CARD if r_idx % 2 == 1 else brp.C_CARD2)
        tc = (255, 255, 255) if is_hl else brp.C_SUB
        brp.style_cell(t2.cell(r_idx, 0), name, size=12, color=brp.C_GREEN if is_hl else brp.C_TEXT, bg=bg_c, bold=is_hl)
        brp.style_cell(t2.cell(r_idx, 1), avg, size=11, color=brp.C_GREEN if is_hl else tc, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)
        w_col = brp.C_GREEN if is_hl else (brp.C_RED if "n²" in worst else tc)
        brp.style_cell(t2.cell(r_idx, 2), worst, size=11, color=w_col, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)
        brp.style_cell(t2.cell(r_idx, 3), comp, size=11, color=brp.C_AMBER if is_hl else tc, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)
        brp.style_cell(t2.cell(r_idx, 4), stb, size=11, color=brp.C_CYAN if is_hl else tc, bg=bg_c, bold=is_hl, font_name="Consolas", align=PP_ALIGN.CENTER)

    brp.add_rect(s2, 8.50, 1.28, 4.28, 5.40, brp.C_CARD, brp.C_GREEN, 1.3, rounded=True)
    brp.add_text(s2, 8.74, 1.46, 3.80, 0.42, "ทำไมต้อง Radix Sort (LSD)?", size=17.5, color=brp.C_GREEN, bold=True)
    for idx, (head, desc, col) in enumerate([
        ("1. ทะลุเพดาน O(n log n)", "Comparison Sort ทุกตัวติดขอบล่าง Ω(n log n) แต่ Radix Sort ไม่เทียบค่าเลย จึงทำเวลาเชิงเส้น O(n) ได้", brp.C_GREEN),
        ("2. แก้จุดอ่อนของ Counting Sort", "Counting Sort ต้องจองช่องเท่าค่าสูงสุด k (ถ้าเลขถึง 1 พันล้าน = พัง!) แต่ Radix Sort แยกเรียงทีละหลัก ใช้ถังแค่ 10 หรือ 256 ช่อง", brp.C_CYAN),
        ("3. เสถียรทุกกรณี + Stable", "Best = Average = Worst เท่ากันหมด ไม่มี Worst Case O(n²) และรักษาลำดับค่าซ้ำ 100%", brp.C_INDIGO),
        ("4. ภาพและเสียงโดดเด่นในวิดีโอ", "ในวิดีโอใช้ฐาน 4 (Base 4) เกิดคลื่นการจัดกลุ่ม 4 ถังที่สวยงามและจดจำได้ทันที", brp.C_AMBER),
    ]):
        ry = 2.02 + idx * 1.14
        brp.add_rect(s2, 8.74, ry, 0.08, 0.92, col, rounded=False)
        brp.add_text(s2, 8.94, ry - 0.04, 3.64, 0.30, head, size=12.5, color=col, bold=True)
        brp.add_text(s2, 8.94, ry + 0.27, 3.64, 0.72, desc, size=11, color=brp.C_SUB, line_spacing=1.18)

    brp.set_notes(
        s2,
        "[คนที่ 1 — สไลด์ 2/2 | เวลา 0:25–1:05 (40 วินาที)]\n"
        "อัลกอริทึมส่วนใหญ่ตั้งแต่ Bubble, Selection, Insertion, Shell ไปจนถึง Quick, Merge และ Heap Sort "
        "ล้วนอาศัยการเปรียบเทียบ จึงติดขอบล่างทางคณิตศาสตร์ที่ O(n log n)\n"
        "ส่วน Counting Sort แม้ไม่เปรียบเทียบ แต่ถ้าค่าเลขสูงถึง 1 พันล้านจะต้องจองหน่วยความจำมหาศาลจนใช้งานไม่ได้\n"
        "Radix Sort (LSD) แก้ปัญหานี้ด้วยการแยกเรียง 'ทีละหลัก' ใช้ถังแค่ 10 หรือ 256 ช่อง และได้ความเร็วเชิงเส้นครับ\n"
        "▶ ส่งต่อคนที่ 2: 'ต่อไปขอเชิญเพื่อนคนที่ 2 อธิบายหลักการทำงานทีละหลักและโค้ด Python ครับ'",
    )

    out = PARTS_DIR / "Part1_Intro_Overview.pptx"
    prs.save(str(out))
    return out


def build_part_2(img_step: Path) -> Path:
    """Part 2 (คนที่ 2): หลักการทำงานทีละหลัก และโค้ด Python [1:05–2:35]"""
    prs = _new_prs()
    blank = prs.slide_layouts[6]

    # Slide 1: Step-by-Step Diagram
    s1 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s1)
    _add_part_footer(
        s1, "ส่วนที่ 2/4 (คนที่ 2): หลักการทำงานและโค้ดโปรแกรม",
        "หลักการทำงานของ Radix Sort (LSD): เรียงทีละหลักจากขวาไปซ้าย",
        "ตัวอย่างอาร์เรย์ [170, 45, 75, 90, 802, 24, 2, 66] — แบ่งใส่ถัง 0–9 ทีละหลัก (หน่วย → สิบ → ร้อย)",
        1, 2, "1:05–1:55",
    )
    s1.shapes.add_picture(str(img_step), Inches(0.55), Inches(1.22), width=Inches(12.23), height=Inches(5.72))
    brp.set_notes(
        s1,
        "[คนที่ 2 — สไลด์ 1/2 | เวลา 1:05–1:55 (50 วินาที)]\n"
        "ขอบคุณครับ ต่อไปผมจะอธิบายว่าทำไมไม่เทียบตัวเลขเลยแต่ยังเรียงถูกต้องได้ ด้วยตัวอย่าง 8 ตัวคือ [170, 45, 75, 90, 802, 24, 2, 66] ครับ\n"
        "• รอบที่ 1 (สีฟ้า — หลักหน่วย): ดูแค่เลขขวาสุดแล้วหย่อนลงถัง 0–9 จากนั้นหยิบออกจากถัง 0 ไล่ไป 9\n"
        "• รอบที่ 2 (สีส้ม — หลักสิบ): ดูเลขตรงกลาง สังเกต 170 กับ 075 มีหลักสิบเป็น 7 เหมือนกัน แต่ 170 เข้าก่อนจึงออกก่อน 075 (Stable!)\n"
        "• รอบที่ 3 (สีเขียว — หลักร้อย): ดูเลขซ้ายสุด พอหยิบออกจากถัง 0, 1, 8 ก็จะได้ลำดับที่ถูกต้องสมบูรณ์ [2, 24, 45, 66, 75, 90, 170, 802] โดยไม่เปรียบเทียบเลยครับ!",
    )

    # Slide 2: Code & Stability
    s2 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s2)
    _add_part_footer(
        s2, "ส่วนที่ 2/4 (คนที่ 2): หลักการทำงานและโค้ดโปรแกรม",
        "เจาะลึกโค้ด Python (Base 10 & Base 256) และหัวใจของความ Stable",
        "โค้ดสั้นกระชับเพียง ~12 บรรทัด — รองรับทั้งฐานสิบ (อ่านง่าย) และฐาน 256 Bitwise (เร็วสูงสุด)",
        2, 2, "1:55–2:35",
    )
    brp.add_rect(s2, 0.55, 1.28, 6.85, 5.50, (9, 14, 26), brp.C_BORDER, 1.3, rounded=True)
    brp.add_rect(s2, 0.55, 1.28, 6.85, 0.44, (18, 28, 48), brp.C_BORDER, 1.0, rounded=False)
    brp.add_text(s2, 0.75, 1.34, 6.4, 0.34, "radix_sort_lsd.py  (Pure Python — 0 Comparisons)", size=12, color=brp.C_GREEN, bold=True, font_name="Consolas")
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
    brp.add_text(s2, 0.75, 1.82, 6.50, 4.82, code_text, size=11.0, color=(226, 232, 240), font_name="Consolas", line_spacing=1.18)

    brp.add_rect(s2, 7.62, 1.28, 5.16, 5.50, brp.C_CARD, brp.C_BORDER, 1.3, rounded=True)
    brp.add_text(s2, 7.86, 1.46, 4.70, 0.40, "กลไกสำคัญที่ทำให้ Radix Sort ถูกต้องและเร็ว", size=16.5, color=brp.C_CYAN, bold=True)
    for idx, (ch, cb, col) in enumerate([
        ("1. กฎเหล็ก: ต้องเป็น Stable ในทุกรอบ!", "เมื่อใส่ค่าลงถังด้วย .append(x) (FIFO) ค่าที่เรียงถูกมาจากหลักขวาแล้วจะไม่สลับที่กันเมื่อหลักซ้ายมีค่าเท่ากัน (เช่น 045 กับ 075)", brp.C_GREEN),
        ("2. สูตรดึงค่าหลักแบบ O(1)", "• ฐาน 10: digit = (x // exp) % 10\n• ฐาน 256: digit = (x >> shift) & 255  (เร็วกว่าหารและ mod มาก!)", brp.C_CYAN),
        ("3. เลือกฐาน b = 256 ลดรอบจาก 9 → 4 รอบ", "เลข 0–999,999,999 (30 บิต) ใช้ฐาน 10 ต้องกวาด 9 รอบ แต่ใช้ฐาน 256 กวาดทีละ 8 บิต เหลือแค่ 4 รอบ!", brp.C_AMBER),
        ("4. รองรับเลขติดลบได้อย่างไร?", "แยกกลุ่มเลขลบกับเลขบวก หรือแปลง Sign Bit (XOR 0x80000000) ก่อนเรียงแล้วแปลงกลับ", brp.C_INDIGO),
    ]):
        sy = 2.00 + idx * 1.15
        brp.add_rect(s2, 7.86, sy, 4.68, 1.02, brp.C_CARD2, col, 1.0, rounded=True)
        brp.add_text(s2, 8.02, sy + 0.08, 4.38, 0.30, ch, size=12, color=col, bold=True)
        brp.add_text(s2, 8.02, sy + 0.38, 4.38, 0.58, cb, size=10.8, color=brp.C_SUB, line_spacing=1.16)

    brp.set_notes(
        s2,
        "[คนที่ 2 — สไลด์ 2/2 | เวลา 1:55–2:35 (40 วินาที)]\n"
        "ในโค้ด Python จะเห็นว่าสั้นมากครับ มีแค่ 2 ขั้นคือแยกใส่ถัง (Distribute) แล้วรวบรวมคืน (Gather)\n"
        "กฎเหล็กคือต้อง Stable ด้วยการใช้ .append() แบบเข้าก่อนออกก่อน เพื่อให้ลำดับจากหลักขวาไม่พัง\n"
        "และเมื่อเปลี่ยนจากฐาน 10 มาใช้ฐาน 256 ด้วย Bit-Shift `(x >> shift) & 255` จะลดรอบการกวาดจาก 9 รอบเหลือแค่ 4 รอบ เร็วขึ้นกว่า 2 เท่าครับ\n"
        "▶ ส่งต่อคนที่ 3: 'ต่อไปเชิญเพื่อนคนที่ 3 สาธิตโปรแกรมจำลองภาพ-เสียง และโชว์ผลการจับเวลาจริงถึง 10 ล้านตัวครับ'",
    )

    out = PARTS_DIR / "Part2_HowItWorks_Code.pptx"
    prs.save(str(out))
    return out


def build_part_3(data: dict, img_sim: Path, img_bench: Path) -> Path:
    """Part 3 (คนที่ 3): โปรแกรมจำลองตามวิดีโอ และผลการทดลอง Benchmark [2:35–3:55]"""
    prs = _new_prs()
    blank = prs.slide_layouts[6]

    # Slide 1: Simulator
    s1 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s1)
    _add_part_footer(
        s1, "ส่วนที่ 3/4 (คนที่ 3): โปรแกรมจำลองและผลการทดลองจับเวลา",
        "โปรแกรมจำลองการจัดเรียง (Simulation ตามวิดีโอตัวอย่าง)",
        "ถอดแบบ Algorithm #6: Radix Sort (LSD) Base 4 จากวิดีโอ พร้อมแถบถังสดและเสียง Web Audio API",
        1, 2, "2:35–3:10",
    )
    s1.shapes.add_picture(str(img_sim), Inches(0.55), Inches(1.28), width=Inches(8.25), height=Inches(4.86))
    for idx, (ch, cb, col) in enumerate([
        ("🎬 ตรงตามวิดีโอต้นฉบับ (Base 4)", "Timo Bingmann ใช้ RADIX = 4 (ทีละ 2 บิต) ทำให้เห็นคลื่น 4 ทางลาดซ้อนกันค่อย ๆ รวมตัวเป็นเส้นตรงในรอบสุดท้าย พร้อม Comparisons = 0!", brp.C_CYAN),
        ("🔊 เสียงสังเคราะห์ Web Audio API", "สร้างคลื่น Triangle Wave ตามความสูงแท่ง (140–1,250 Hz) ได้ยินจังหวะ Scan (สีส้ม) สลับกับจังหวะ Gather จากถัง 0→3 (สีแดง)", brp.C_AMBER),
        ("🎛 โหมดสาธิต & ไฟล์โปรแกรม", "• สลับฐานได้ทันที: Base 4 / 10 / 2 / 8\n• index.html / sound_of_sorting_radix_web.html\n• radix_sort_visualizer.py (Python Tkinter)", brp.C_GREEN),
    ]):
        cy = 1.28 + idx * 1.70
        brp.add_rect(s1, 9.02, cy, 3.76, 1.54, brp.C_CARD, col, 1.2, rounded=True)
        brp.add_text(s1, 9.20, cy + 0.12, 3.42, 0.34, ch, size=13, color=col, bold=True)
        brp.add_text(s1, 9.20, cy + 0.48, 3.42, 0.98, cb, size=11.2, color=brp.C_SUB, line_spacing=1.20)

    brp.add_rect(s1, 0.55, 6.26, 8.25, 0.56, brp.C_CARD, brp.C_BORDER, 1.0, rounded=True)
    brp.add_text(
        s1, 0.72, 6.34, 7.90, 0.40,
        "💡 สังเกตเอกลักษณ์ของ Radix Sort: ระหว่างจัดเรียงจะมีคลื่นย่อย 4 ช่วง (ถัง 0–3) แล้วรอบสุดท้ายจะ 'ดีด' เรียงสมบูรณ์พร้อมกันทั้งกราฟ!",
        size=11.2, color=brp.C_SUB, valign=MSO_ANCHOR.MIDDLE,
    )
    brp.set_notes(
        s1,
        "[คนที่ 3 — สไลด์ 1/2 | เวลา 2:35–3:10 (35 วินาที)]\n"
        "ขอบคุณครับ สำหรับส่วนของผมจะพาไปชมโปรแกรมจำลองที่เราสร้างตามวิดีโอ 'The Sound of Sorting' ครับ\n"
        "เราสร้างทั้งเวอร์ชันเว็บ HTML5 ที่มีเสียงสังเคราะห์สด และ Python Tkinter โดยใช้ Base 4 เหมือนโค้ดของ Timo Bingmann\n"
        "แต่ละรอบจะมีจังหวะสแกนลงถัง 0–3 (สีส้ม) สลับกับดึงคืนจากถังทีละใบ (สีแดง) จะเห็นคลื่นย่อย 4 ช่วงก่อนดีดเรียงสมบูรณ์ในรอบสุดท้าย "
        "และตัวเลข Comparisons เป็น 0 ตลอดเวลาครับ (เปิดเดโมสดในเบราว์เซอร์ 10–15 วิตรงนี้ได้ครับ!)",
    )

    # Slide 2: Benchmark
    s2 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s2)
    _add_part_footer(
        s2, "ส่วนที่ 3/4 (คนที่ 3): โปรแกรมจำลองและผลการทดลองจับเวลา",
        "ผลการทดลองวัดเวลา (Benchmark Analysis)",
        "สุ่มตัวเลข 9 หลัก (0–999,999,999) บน Pure Python 3.12 / AMD Ryzen 5 220 / RAM 16 GB — ตรวจสอบความถูกต้องครบทุกรอบ",
        2, 2, "3:10–3:55",
    )
    t6 = s2.shapes.add_table(5, 6, Inches(0.55), Inches(1.22), Inches(12.23), Inches(2.12)).table
    for idx, w in enumerate([1.95, 2.10, 2.05, 1.95, 2.08, 2.10]):
        t6.columns[idx].width = Inches(w)
    for c, h in enumerate(["จำนวนตัวเลขสุ่ม (N)", "Radix Base 256 (4 รอบ)", "Radix Base 10 (9 รอบ)", "เทียบ Merge Sort", "อัตราโต (Base 256)", "Comparisons / Writes"]):
        brp.style_cell(t6.cell(0, c), h, size=11, color=brp.C_CYAN, bg=(17, 30, 54), bold=True, align=PP_ALIGN.CENTER)

    ref_m = data["reference_merge_sort"]
    prev_t = None
    for idx, (r256, r10) in enumerate(zip(data["scaling_base256"], data["scaling_base10"]), start=1):
        n = r256["n"]
        t256, t10 = r256["avg_sec"], r10["avg_sec"]
        ratio_txt = "—" if prev_t is None else f"×{t256 / prev_t:.2f} (N ×10)"
        prev_t = t256
        bg_c = (12, 46, 42) if n == 1_000_000 else (brp.C_CARD if idx % 2 == 1 else brp.C_CARD2)
        f256 = f"{t256*1000:.2f} ms" if t256 < 0.1 else f"{t256:.4f} s"
        f10 = f"{t10*1000:.2f} ms" if t10 < 0.1 else f"{t10:.4f} s"
        fm = f"{ref_m[str(n)]:.4f} s" if str(n) in ref_m else "— (ช้ากว่า ~4×)"
        brp.style_cell(t6.cell(idx, 0), f"{n:,}", size=12, color=brp.C_TEXT, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.RIGHT)
        brp.style_cell(t6.cell(idx, 1), f256, size=12, color=brp.C_GREEN, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.CENTER)
        brp.style_cell(t6.cell(idx, 2), f10, size=11.5, color=brp.C_CYAN, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.CENTER)
        brp.style_cell(t6.cell(idx, 3), fm, size=11, color=brp.C_INDIGO, bg=bg_c, font_name="Consolas", align=PP_ALIGN.CENTER)
        brp.style_cell(t6.cell(idx, 4), ratio_txt, size=11, color=brp.C_AMBER, bg=bg_c, bold=True, font_name="Consolas", align=PP_ALIGN.CENTER)
        brp.style_cell(t6.cell(idx, 5), f"0 / {r256['writes']:,}", size=10.5, color=brp.C_SUB, bg=bg_c, font_name="Consolas", align=PP_ALIGN.RIGHT)

    s2.shapes.add_picture(str(img_bench), Inches(0.55), Inches(3.46), width=Inches(12.23), height=Inches(3.48))
    brp.set_notes(
        s2,
        "[คนที่ 3 — สไลด์ 2/2 | เวลา 3:10–3:55 (45 วินาที)]\n"
        "นี่คือผลการรันจับเวลาจริงบน Pure Python ครบทั้ง 10,000 / 100,000 / 1,000,000 และ 10,000,000 ตัวครับ:\n"
        "• 10,000 ตัว: Base 256 ใช้แค่ 5.33 มิลลิวินาที (เร็วกว่า Selection Sort ถึง 490 เท่า!)\n"
        "• 100,000 ตัว: ใช้เพียง 62.45 มิลลิวินาที\n"
        "• 1,000,000 ตัว: จบใน 1.22 วินาที เร็วกว่า Merge Sort O(n log n) ถึง 3.26 เท่า!\n"
        "• 10,000,000 ตัว: จบใน 13.13 วินาที เพิ่มขึ้นเพียง 10.76 เท่าเมื่อข้อมูลเพิ่ม 10 เท่า — โตเป็นเส้นตรงชัดเจนครับ\n"
        "▶ ส่งต่อคนที่ 4: 'ต่อไปเชิญเพื่อนคนที่ 4 วิเคราะห์ Runtime Big-O และสรุปการใช้งานจริงครับ'",
    )

    out = PARTS_DIR / "Part3_Simulation_Benchmark.pptx"
    prs.save(str(out))
    return out


def build_part_4(img_bigo: Path) -> Path:
    """Part 4 (คนที่ 4): การวิเคราะห์ Runtime Big-O และสรุปผล [3:55–5:00]"""
    prs = _new_prs()
    blank = prs.slide_layouts[6]

    # Slide 1: Big-O
    s1 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s1)
    _add_part_footer(
        s1, "ส่วนที่ 4/4 (คนที่ 4): การวิเคราะห์ Runtime Big-O และสรุป",
        "การวิเคราะห์ Runtime Big-O (Time & Space Complexity)",
        "ที่มาของเวลาเชิงเส้น O(d(n + b)) และหลักฐานยืนยันจากการทดลองจริงทั้งมิติของ N และจำนวนรอบ d",
        1, 2, "3:55–4:35",
    )
    brp.add_rect(s1, 0.55, 1.28, 4.30, 2.86, brp.C_CARD, brp.C_GREEN, 1.4, rounded=True)
    brp.add_text(s1, 0.78, 1.42, 3.85, 0.36, "⏱ Time Complexity (เวลาทำงาน)", size=15, color=brp.C_GREEN, bold=True)
    brp.add_text(s1, 0.78, 1.84, 3.85, 0.46, "Best / Avg / Worst = Θ(d(n + b))", size=14.5, color=brp.C_CYAN, bold=True, font_name="Consolas")
    brp.add_text(
        s1, 0.78, 2.34, 3.85, 1.66,
        "• n = จำนวนข้อมูลทั้งหมด\n"
        "• b = ฐานเลข / จำนวนถัง (เช่น 10 หรือ 256)\n"
        "• d = จำนวนหลัก = ⌈log_b(max + 1)⌉\n"
        "• เมื่อค่าคงที่ (เช่น 32-bit int, b=256 → d=4)\n"
        "  จะได้เวลาทำงานเชิงเส้น O(n)!",
        size=11.8, color=brp.C_SUB, line_spacing=1.24,
    )

    brp.add_rect(s1, 0.55, 4.30, 4.30, 2.48, brp.C_CARD, brp.C_INDIGO, 1.4, rounded=True)
    brp.add_text(s1, 0.78, 4.44, 3.85, 0.36, "💾 Space Complexity (หน่วยความจำ)", size=15, color=brp.C_INDIGO, bold=True)
    brp.add_text(s1, 0.78, 4.86, 3.85, 0.46, "Auxiliary Space = O(n + b)", size=16, color=brp.C_AMBER, bold=True, font_name="Consolas")
    brp.add_text(
        s1, 0.78, 5.36, 3.85, 1.32,
        "• ถังเก็บรอแยกข้อมูล (buckets) รวม n ตัว = O(n)\n"
        "• โครงสร้างถัง/ตารางนับความถี่ b ช่อง = O(b)\n"
        "• ไม่อาศัย Recursion → Stack Space = O(1)",
        size=11.8, color=brp.C_SUB, line_spacing=1.24,
    )

    s1.shapes.add_picture(str(img_bigo), Inches(5.05), Inches(1.28), width=Inches(7.73), height=Inches(5.50))
    brp.set_notes(
        s1,
        "[คนที่ 4 — สไลด์ 1/2 | เวลา 3:55–4:35 (40 วินาที)]\n"
        "ขอบคุณครับ ต่อไปผมขออธิบาย Runtime Big-O ที่อยู่เบื้องหลังความเร็วนี้ครับ\n"
        "• Time Complexity คือ O(d × (n + b)) เท่ากันทั้ง Best, Average และ Worst Case เพราะทุกรอบต้องสแกนข้อมูล n ตัวลงถัง b ใบแล้วดึงคืน ทำซ้ำ d รอบ\n"
        "• สำหรับเลข 32-bit เมื่อใช้ฐาน b = 256 ค่า d คงที่แค่ 4 รอบ จึงลดรูปเป็น O(n) เชิงเส้น!\n"
        "• กราฟขวาล่างพิสูจน์จากการรันจริง: ไม่ว่า Base 4, 10, 16 หรือ 256 เวลาเฉลี่ยต่อ 1 รอบคงที่ที่ ~0.28–0.30 วินาทีเป๊ะ "
        "ส่วน Space Complexity คือ O(n + b) และไม่ใช้ Recursion Stack ครับ",
    )

    # Slide 2: Summary & Closing
    s2 = prs.slides.add_slide(blank)
    brp.set_slide_bg(s2)
    _add_part_footer(
        s2, "ส่วนที่ 4/4 (คนที่ 4): การวิเคราะห์ Runtime Big-O และสรุป",
        "สรุป ข้อดี-ข้อจำกัด และการประยุกต์ใช้งานจริง",
        "บทสรุปของ Radix Sort (LSD) — อัลกอริทึมความเร็วเชิงเส้นที่อยู่เบื้องหลังงานประมวลผลข้อมูลขนาดใหญ่",
        2, 2, "4:35–5:00",
    )
    for idx, (heading, col, items) in enumerate([
        (
            "✅ จุดเด่น (Strengths)", brp.C_GREEN,
            [
                ("เร็วระดับเชิงเส้น O(d(n+b))", "เรียงตัวเลข 1 ล้านตัวใน 1.22 วินาที (เร็วกว่า Merge Sort 3.26× และเร็วกว่า O(n²) กว่า 20,000×)"),
                ("0 Comparisons & Stable", "ไม่เปรียบเทียบค่าเลย ไม่มีจุดอ่อน Worst Case จากข้อมูลเรียงกลับด้าน และรักษาลำดับค่าซ้ำ"),
                ("เหมาะกับ Parallel & GPU", "โครงสร้างแยกถังทำงานขนานได้ยอดเยี่ยม เป็นอัลกอริทึมเรียงลำดับที่เร็วที่สุดบน GPU"),
            ],
        ),
        (
            "⚠️ ข้อจำกัด (Trade-offs)", brp.C_AMBER,
            [
                ("จำกัดชนิดของคีย์ข้อมูล", "ใช้ได้ดีกับจำนวนเต็ม, สตริงความยาวคงที่ หรือคีย์ที่แยกเป็นหลัก/ไบต์ได้ ไม่ยืดหยุ่นเท่า Comparison Sort"),
                ("ใช้หน่วยความจำเพิ่ม O(n+b)", "ไม่เป็น In-place Sort ต้องมีพื้นที่สำหรับถังหรืออาร์เรย์ผลลัพธ์เพิ่มอีกเท่าตัว"),
                ("ช้าลงถ้าจำนวนหลัก d สูงมาก", "หากคีย์ยาวมาก (เช่น สตริงยาวหลายร้อยตัวอักษร) ค่า d จะสูงจนอาจแพ้ Quick/Merge Sort"),
            ],
        ),
        (
            "🚀 การใช้งานในโลกจริง", brp.C_CYAN,
            [
                ("NVIDIA CUDA / GPU Sorting", "ไลบรารี Thrust & CUB ใช้ Radix Sort เป็นตัวหลักในการเรียงข้อมูลหลายร้อยล้านแถวบน GPU"),
                ("ฐานข้อมูล, IP & วันที่ (Multi-Key)", "เรียงหมายเลขบัตร, IPv4, Timestamp หรือรหัสไปรษณีย์ได้อย่างรวดเร็วและเสถียร"),
                ("Bioinformatics & Suffix Array", "ใช้สร้าง Suffix Array (อัลกอริทึม DC3) สำหรับค้นหาลำดับ DNA และงาน Text Indexing"),
            ],
        ),
    ]):
        cx = 0.55 + idx * 4.16
        brp.add_rect(s2, cx, 1.32, 3.92, 4.62, brp.C_CARD, col, 1.3, rounded=True)
        brp.add_text(s2, cx + 0.22, 1.50, 3.48, 0.42, heading, size=16.5, color=col, bold=True)
        for j, (ih, ib) in enumerate(items):
            iy = 2.08 + j * 1.24
            brp.add_rect(s2, cx + 0.22, iy, 3.48, 1.08, brp.C_CARD2, brp.C_BORDER, 0.8, rounded=True)
            brp.add_text(s2, cx + 0.36, iy + 0.08, 3.22, 0.30, ih, size=12, color=brp.C_TEXT, bold=True)
            brp.add_text(s2, cx + 0.36, iy + 0.38, 3.22, 0.66, ib, size=10.8, color=brp.C_SUB, line_spacing=1.16)

    brp.add_rect(s2, 0.55, 6.14, 12.23, 0.76, (12, 42, 38), brp.C_GREEN, 1.2, rounded=True)
    brp.add_text(
        s2, 0.78, 6.24, 8.45, 0.54,
        "สรุป: เมื่อข้อมูลเป็นตัวเลขหรือคีย์ที่มีหลักชัดเจน Radix Sort (LSD) ชนะ O(n log n) ได้จริง — ขอบคุณครับ / Q&A",
        size=13.0, color=brp.C_TEXT, bold=True, valign=MSO_ANCHOR.MIDDLE,
    )
    brp.add_text(
        s2, 9.30, 6.24, 3.25, 0.54,
        "1M = 1.22s  |  10M = 13.13s!",
        size=13.5, color=brp.C_GREEN, bold=True, font_name="Consolas",
        align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE,
    )
    brp.set_notes(
        s2,
        "[คนที่ 4 — สไลด์ 2/2 | เวลา 4:35–5:00 (25 วินาที)]\n"
        "โดยสรุปครับ Radix Sort (LSD) แสดงให้เห็นว่าเราไม่จำเป็นต้องเปรียบเทียบตัวเลขสองตัวเสมอไป "
        "แม้มีข้อจำกัดเรื่องชนิดคีย์และหน่วยความจำเพิ่ม O(n+b) แต่ด้วยความเร็วเชิงเส้นและทำขนานได้ง่าย "
        "จึงถูกใช้งานจริงอย่างแพร่หลายทั้งบน GPU, ฐานข้อมูล และงาน DNA ครับ\n"
        "กลุ่มของพวกเราขอจบการนำเสนอเพียงเท่านี้ ขอบคุณอาจารย์และเพื่อน ๆ ครับ มีคำถามสอบถามได้เลยครับ!",
    )

    out = PARTS_DIR / "Part4_BigO_Summary.pptx"
    prs.save(str(out))
    return out


def write_speaker_script() -> Path:
    script = r"""# บทพูดนำเสนอ 5 นาที (แบ่ง 4 คน — คนละ 2 สไลด์)
> **หัวข้อ:** Radix Sort (LSD) — การจัดเรียงตามหลักตัวเลข (0 Comparisons!)  
> **เวลารวม:** 5:00 นาที (เฉลี่ยคนละ 1:05 – 1:30 นาที)

---

## ตารางแบ่งหน้าที่ 4 คน

| ผู้พูด | ไฟล์สไลด์ | หัวข้อที่รับผิดชอบ | ช่วงเวลา | ระยะเวลา |
| :---: | :--- | :--- | :---: | :---: |
| **คนที่ 1** | [`Part1_Intro_Overview.pptx`](./Part1_Intro_Overview.pptx) | บทนำ, ภาพรวมอัลกอริทึมจัดเรียง และเหตุผลที่เลือก Radix Sort (LSD) | `0:00 – 1:05` | 1:05 นาที |
| **คนที่ 2** | [`Part2_HowItWorks_Code.pptx`](./Part2_HowItWorks_Code.pptx) | หลักการทำงานทีละหลัก (หน่วย→สิบ→ร้อย), โค้ด Python และความ Stable | `1:05 – 2:35` | 1:30 นาที |
| **คนที่ 3** | [`Part3_Simulation_Benchmark.pptx`](./Part3_Simulation_Benchmark.pptx) | สาธิตโปรแกรมจำลองตามวิดีโอ (Base 4 + เสียงสด) และผลจับเวลาถึง 10 ล้านตัว | `2:35 – 3:55` | 1:20 นาที |
| **คนที่ 4** | [`Part4_BigO_Summary.pptx`](./Part4_BigO_Summary.pptx) | วิเคราะห์ Runtime Big-O ($O(d(n+b))$ / $O(n+b)$), ข้อดี-ข้อจำกัด และสรุปจบ | `3:55 – 5:00` | 1:05 นาที |

---

## บทพูดและจังหวะส่งต่อ (Handoff Script)

### คนที่ 1 — บทนำ & ทำไมต้อง Radix Sort (`0:00 – 1:05`)
- **สไลด์ 1 (0:00–0:25):** สวัสดีอาจารย์และเพื่อน ๆ ครับ วันนี้กลุ่มของพวกเรานำเสนออัลกอริทึมลำดับที่ 6 จากวิดีโอ *15 Sorting Algorithms in 6 Minutes* นั่นคือ **Radix Sort (LSD)** จุดที่น่าทึ่งที่สุดคือมัน **ไม่เคยเปรียบเทียบตัวเลขสองตัวว่าใครมากกว่าเลยสักครั้ง (`0 Comparisons`)** ทำให้ทะลุขีดจำกัด $O(n \log n)$ ได้ครับ
- **สไลด์ 2 (0:25–1:05):** จากตารางจะเห็นว่า Comparison Sort ทุกตัวติดเพดานเร็วสุดที่ $O(n \log n)$ ส่วน Counting Sort แม้ไม่เปรียบเทียบแต่ถ้าเลขสูงถึงพันล้านจะจอง RAM ไม่ไหว — Radix Sort (LSD) จึงแก้ด้วยวิธีแยกเรียงทีละหลัก ใช้ถังแค่ 10 หรือ 256 ใบ
- **ประโยคส่งต่อ:** *“ต่อไปขอเชิญเพื่อนคนที่ 2 อธิบายหลักการทำงานทีละหลักและโค้ด Python ครับ”*

### คนที่ 2 — หลักการทำงาน & โค้ด Python (`1:05 – 2:35`)
- **สไลด์ 1 (1:05–1:55):** ยกตัวอย่าง 8 ตัว `[170, 45, 75, 90, 802, 24, 2, 66]` ไล่เรียงจากขวาไปซ้าย 3 รอบ:
  1. รอบที่ 1 (หลักหน่วย): หย่อนลงถัง `0–9` ตามเลขขวาสุดแล้วหยิบออกตามลำดับ
  2. รอบที่ 2 (หลักสิบ): หย่อนตามเลขกลาง — เน้นว่า `170` กับ `075` หลักสิบเป็น `7` เหมือนกัน แต่ `170` เข้าก่อนออกก่อน (**Stable**)
  3. รอบที่ 3 (หลักร้อย): หยิบออกจากถัง `0, 1, 8` → ได้ `[2, 24, 45, 66, 75, 90, 170, 802]` เรียงเสร็จโดยไม่เทียบค่าเลย!
- **สไลด์ 2 (1:55–2:35):** โค้ด Python มีแค่ 2 ขั้นคือแยกใส่ถัง (`.append` รักษา Stable) แล้วรวบรวมคืน และเมื่อเปลี่ยนจากฐาน 10 (9 รอบ) เป็น **ฐาน 256 Bitwise** `(x >> shift) & 255` จะเหลือแค่ **4 รอบ** เร็วขึ้นกว่า 2 เท่า
- **ประโยคส่งต่อ:** *“ต่อไปเชิญเพื่อนคนที่ 3 สาธิตโปรแกรมจำลองภาพ-เสียง และโชว์ผลการจับเวลาจริงครับ”*

### คนที่ 3 — โปรแกรมจำลอง & ผล Benchmark (`2:35 – 3:55`)
- **สไลด์ 1 (2:35–3:10):** แสดงโปรแกรมจำลองที่ถอดแบบ Base 4 จากวิดีโอต้นฉบับ (เปิดไฟล์เว็บ `index.html` โชว์สด 10–15 วิ พร้อมเสียง Web Audio API!) ชี้ให้ดูแถบถัง 0–3 ด้านล่าง และค่า `Comparisons = 0`
- **สไลด์ 2 (3:10–3:55):** สรุปผลจับเวลาจริง Pure Python 3.12:
  - `10,000` ตัว = **5.33 ms** (เร็วกว่า Selection Sort 490 เท่า)
  - `100,000` ตัว = **62.45 ms**
  - `1,000,000` ตัว = **1.22 วินาที** (เร็วกว่า Merge Sort 3.26 เท่า!)
  - `10,000,000` ตัว = **13.13 วินาที** (เพิ่มแค่ ×10.76 เท่าเมื่อข้อมูลเพิ่ม ×10 เท่า — โตเป็นเส้นตรง!)
- **ประโยคส่งต่อ:** *“ต่อไปเชิญเพื่อนคนที่ 4 วิเคราะห์ Runtime Big-O และสรุปปิดท้ายครับ”*

### คนที่ 4 — Runtime Big-O & สรุปปิดท้าย (`3:55 – 5:00`)
- **สไลด์ 1 (3:55–4:35):** อธิบาย **Time Complexity = $\Theta(d(n + b))$** เท่ากันทุกกรณี (Best/Avg/Worst) โดยเมื่อใช้เลข 32-bit และฐาน 256 ($d = 4$) จะลดรูปเป็น **$O(n)$ เชิงเส้น** พร้อมชี้กราฟพิสูจน์ว่าเวลาต่อ 1 รอบคงที่เป๊ะที่ `0.28–0.30` วินาทีทุกฐาน ส่วน **Space Complexity = $O(n + b)$** และไม่ใช้ Recursion Stack
- **สไลด์ 2 (4:35–5:00):** สรุปจุดเด่น (เร็วเชิงเส้น, 0 Comparisons, Stable, ทำขนานบน GPU ได้ดี) ข้อจำกัด (จำกัดชนิดคีย์, ใช้พื้นที่เพิ่ม $O(n+b)$) และการใช้งานจริง (NVIDIA CUDA Thrust, ฐานข้อมูล, DNA Suffix Array) — *“ขอบคุณอาจารย์และเพื่อน ๆ ครับ มีคำถามสอบถามได้เลยครับ!”*
"""
    out = PARTS_DIR / "Speaker_Script_4_People.md"
    out.write_text(script, encoding="utf-8")
    return out


def main() -> None:
    data = json.loads((BASE_DIR / "benchmark_results.json").read_text(encoding="utf-8"))
    assets = BASE_DIR / "assets"
    img_step = assets / "radix_step_diagram.png"
    img_sim = assets / "radix_simulator_preview.png"
    img_bench = assets / "radix_benchmark_chart.png"
    img_bigo = assets / "radix_bigo_diagram.png"

    p1 = build_part_1()
    p2 = build_part_2(img_step)
    p3 = build_part_3(data, img_sim, img_bench)
    p4 = build_part_4(img_bigo)
    sc = write_speaker_script()
    for p in (p1, p2, p3, p4, sc):
        print(f"Created: {p}")


if __name__ == "__main__":
    main()
