# Radix Sort (LSD) — การจัดเรียงตามหลักตัวเลข (0 Comparisons!)

[![Algorithm](https://img.shields.io/badge/Algorithm-Radix%20Sort%20(LSD)-10b981?style=flat-square)](#)
[![Time Complexity](https://img.shields.io/badge/Time-O(d(n%2Bb))-38bdf8?style=flat-square)](#runtime-big-o)
[![Comparisons](https://img.shields.io/badge/Comparisons-0%20%E0%B8%84%E0%B8%A3%E0%B8%B1%E0%B9%89%E0%B8%87-f59e0b?style=flat-square)](#)
[![Python](https://img.shields.io/badge/Python-3.12%20Pure%20Python-818cf8?style=flat-square)](#)

โปรเจกต์ศึกษา จำลองภาพ-เสียง ทดลองวัดเวลา และวิเคราะห์ Runtime Big-O ของอัลกอริทึม **Radix Sort (Least Significant Digit)** อ้างอิงจากวิดีโอ [15 Sorting Algorithms in 6 Minutes (The Sound of Sorting)](https://www.youtube.com/watch?v=kPRA0W1kECg) ของ Timo Bingmann

---

## 📂 ไฟล์ทั้งหมดในโปรเจกต์

| ไฟล์ | คำอธิบาย |
| :--- | :--- |
| [`index.html`](./index.html) / [`sound_of_sorting_radix_web.html`](./sound_of_sorting_radix_web.html) | โปรแกรมจำลองแบบเว็บ (HTML5 Canvas + เสียงสังเคราะห์ **Web Audio API**, Base 4 ตามวิดีโอ / Base 10 / Base 2 / Base 8) |
| [`radix_sort_visualizer.py`](./radix_sort_visualizer.py) | โปรแกรมจำลอง Python Desktop GUI (`tkinter`, ไม่ต้องติดตั้งไลบรารีเพิ่ม) |
| [`benchmark_radix.py`](./benchmark_radix.py) | สคริปต์ทดสอบจับเวลาจริง (`10,000` / `100,000` / `1,000,000` / `10,000,000` ตัว) |
| [`benchmark_results.json`](./benchmark_results.json) | ผลการวัดเวลาดิบจากเครื่องทดสอบ |
| [`RadixSort_Presentation.pptx`](./RadixSort_Presentation.pptx) | สไลด์นำเสนอฉบับรวม 5 นาที (8 สไลด์ 16:9 พร้อม Speaker Notes) |
| [`slides_4_parts/`](./slides_4_parts/) | **สไลด์แยก 4 ไฟล์ สำหรับแบ่งกันนำเสนอ 4 คน (คนละ 2 สไลด์)** + [บทพูดและจังหวะส่งต่อ 4 คน](./slides_4_parts/Speaker_Script_4_People.md) |
| [`report_summary.md`](./report_summary.md) | รายงานสรุปการศึกษาและผลวิเคราะห์ฉบับเต็ม |

### 👥 สไลด์แยก 4 ส่วนสำหรับนำเสนอ 4 คน (รวม 5 นาที)

| ผู้พูด | ไฟล์ PowerPoint | เนื้อหาที่รับผิดชอบ | เวลา |
| :---: | :--- | :--- | :---: |
| **คนที่ 1** | [`Part1_Intro_Overview.pptx`](./slides_4_parts/Part1_Intro_Overview.pptx) | บทนำ, ภาพรวมอัลกอริทึม และเหตุผลที่เลือก Radix Sort (LSD) | `0:00 – 1:05` |
| **คนที่ 2** | [`Part2_HowItWorks_Code.pptx`](./slides_4_parts/Part2_HowItWorks_Code.pptx) | หลักการทำงานทีละหลัก (หน่วย→สิบ→ร้อย), โค้ด Python และความ Stable | `1:05 – 2:35` |
| **คนที่ 3** | [`Part3_Simulation_Benchmark.pptx`](./slides_4_parts/Part3_Simulation_Benchmark.pptx) | สาธิตโปรแกรมจำลอง (Base 4 + เสียงสด) และผลจับเวลาถึง 10 ล้านตัว | `2:35 – 3:55` |
| **คนที่ 4** | [`Part4_BigO_Summary.pptx`](./slides_4_parts/Part4_BigO_Summary.pptx) | วิเคราะห์ Runtime Big-O ($O(d(n+b))$ / $O(n+b)$), ข้อดี-ข้อจำกัด และสรุป | `3:55 – 5:00` |
| **บทพูดรวม** | [`Speaker_Script_4_People.md`](./slides_4_parts/Speaker_Script_4_People.md) | บทพูดละเอียด + ประโยคส่งต่อระหว่างทั้ง 4 คน | รวม `5:00` |

---

## 1. หลักการทำงาน (Step-by-Step)

Radix Sort (LSD) แยกจัดเรียงข้อมูล **ทีละหลักจากขวาไปซ้าย** (เช่น หลักหน่วย → สิบ → ร้อย หรือทีละ 2 บิตใน Base 4) โดยหย่อนลงถังตามเลขโดดแล้วดึงกลับแบบ **Stable** จึงไม่ต้องเปรียบเทียบตัวเลขสองตัวเลยสักครั้ง (`Comparisons = 0`)

![Radix Sort LSD Step-by-Step](assets/radix_step_diagram.png)

---

## 2. โปรแกรมจำลองตามวิดีโอตัวอย่าง

![Radix Sort Base 4 Simulator Preview](assets/radix_simulator_preview.png)

### วิธีรันในเครื่อง
```bash
# 1. เปิดเวอร์ชันเว็บ (มีเสียงสังเคราะห์สดตามความสูงของค่า)
# ดับเบิลคลิกไฟล์ index.html หรือ sound_of_sorting_radix_web.html ในเบราว์เซอร์

# 2. รันเวอร์ชัน Python Desktop GUI (Tkinter)
python radix_sort_visualizer.py

# 3. รันทดสอบ Benchmark จับเวลาจริง
python benchmark_radix.py
```

---

## 3. ผลการทดลองวัดเวลา (Benchmark Analysis)

ทดสอบด้วย Pure Python `3.12.10` บน **AMD Ryzen 5 220 / RAM 16 GB** สุ่มจำนวนเต็ม 9 หลัก (`0 – 999,999,999`):

| จำนวนตัวเลขสุ่ม ($N$) | Radix Base 256 (4 รอบ) | Radix Base 10 (9 รอบ) | Merge Sort $O(n \log n)$ | อัตราโต (Base 256) | Comparisons |
| :---: | :---: | :---: | :---: | :---: | :---: |
| **10,000** | **5.33 ms** | 11.32 ms | 22.70 ms | — | 0 |
| **100,000** | **62.45 ms** | 0.1507 s | 0.2988 s | ×11.72 ($N \times 10$) | 0 |
| **1,000,000** | **1.2203 s** | 2.7289 s | 3.9739 s | ×19.54 *(L3 Cache)* | 0 |
| **10,000,000** | **13.1253 s** | 30.9526 s | — | **×10.76** ($N \times 10$) | 0 |

![Benchmark Comparison Chart](assets/radix_benchmark_chart.png)

---

## 4. Runtime Big-O

- **Time Complexity:** Best = Average = Worst = **$\Theta(d \cdot (n + b))$** โดย $d = \lceil \log_b(\max + 1) \rceil$
  - สำหรับจำนวนเต็ม 32-bit เมื่อใช้ฐาน $b = 256$ จะทำคงที่เพียง $d = 4$ รอบ → เวลาเติบโตแบบเชิงเส้น **$O(n)$**
- **Space Complexity:** **$O(n + b)$** (พื้นที่ถังชั่วคราว, Call Stack = $O(1)$)

![Runtime Big-O Diagram](assets/radix_bigo_diagram.png)
