"""
Radix Sort (LSD) Visual Simulator (Python Tkinter — "The Sound of Sorting" Style)
=================================================================================
Zero-dependency native Python desktop simulator modeled after Timo Bingmann's
"15 Sorting Algorithms in 6 Minutes" (Algorithm #6: Radix Sort LSD, Base 4).

Run:
    python radix_sort_visualizer.py
"""

from __future__ import annotations

import queue
import random
import threading
import tkinter as tk
from tkinter import ttk

try:
    import winsound

    HAS_WINSOUND = True
except ImportError:
    HAS_WINSOUND = False


class AsyncBeepPlayer:
    """Plays short non-blocking pitch tones mapped to bar values on Windows."""

    def __init__(self) -> None:
        self._q: queue.Queue[tuple[int, int] | None] = queue.Queue(maxsize=1)
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    def play(self, value: int, max_val: int) -> None:
        if not HAS_WINSOUND:
            return
        norm = max(0.0, min(1.0, value / max(1, max_val)))
        freq = int(220 + 1100 * (norm ** 1.3))
        try:
            if self._q.full():
                try:
                    self._q.get_nowait()
                except queue.Empty:
                    pass
            self._q.put_nowait((freq, 20))
        except queue.Full:
            pass

    def _worker(self) -> None:
        while True:
            item = self._q.get()
            if item is None:
                break
            freq, dur = item
            try:
                winsound.Beep(freq, dur)
            except Exception:
                pass


class RadixSortVisualizerApp:
    BG_DARK = "#080B12"
    BG_CANVAS = "#04060A"
    BG_PANEL = "#101624"
    FG_TEXT = "#F1F5F9"
    FG_MUTED = "#94A3B8"
    COLOR_BAR = "#F1F5F9"
    COLOR_RED = "#FF3B5C"
    COLOR_AMBER = "#F59E0B"
    COLOR_GREEN = "#10B981"
    COLOR_ACCENT = "#38BDF8"

    BUCKET_BANDS = [
        ("#0B2438", "#38BDF8"),
        ("#181B3A", "#818CF8"),
        ("#27163B", "#C084FC"),
        ("#36142A", "#F472B6"),
        ("#35181C", "#FB7185"),
        ("#352212", "#FB923C"),
        ("#332C10", "#FACC15"),
        ("#102E22", "#4ADE80"),
        ("#0F2D2C", "#2DD4BF"),
        ("#13233D", "#60A5FA"),
    ]

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Radix Sort (LSD) Simulator — The Sound of Sorting (Python)")
        self.root.configure(bg=self.BG_DARK)
        self.root.geometry("1120x690")
        self.root.minsize(840, 560)

        self.beeper = AsyncBeepPlayer()

        self.n_var = tk.IntVar(value=96)
        self.base_var = tk.StringVar(value="Base 4 (Video)")
        self.delay_var = tk.IntVar(value=10)
        self.pattern_var = tk.StringVar(value="Random")
        self.sound_var = tk.BooleanVar(value=False)

        self.arr: list[int] = []
        self.reads = 0
        self.writes = 0
        self.current_pass = 0
        self.total_passes = 1
        self.current_exp = 1

        self.highlight_amber = -1
        self.highlight_red = -1
        self.bucket_bounds: list[tuple[int, int, int]] = []
        self.verified_count = 0
        self.phase_text = "READY"

        self.gen = None
        self.running = False
        self.paused = False
        self.after_id = None

        self._build_ui()
        self.reset_array()

    def _get_base(self) -> int:
        txt = self.base_var.get()
        if "Base 10" in txt:
            return 10
        if "Base 2" in txt:
            return 2
        if "Base 8" in txt:
            return 8
        return 4

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, bg=self.BG_PANEL, padx=16, pady=10)
        header.pack(fill=tk.X, padx=14, pady=(12, 8))

        tk.Label(
            header,
            text="Radix Sort (LSD) Simulator  |  การจัดเรียงตามหลักตัวเลข (0 Comparisons!)",
            font=("Leelawadee UI", 14, "bold"),
            bg=self.BG_PANEL,
            fg=self.FG_TEXT,
        ).pack(side=tk.LEFT)

        tk.Label(
            header,
            text="Time: O(d(n+b))  |  Space: O(n+b)  |  Stable",
            font=("Consolas", 10, "bold"),
            bg=self.BG_PANEL,
            fg=self.COLOR_GREEN,
        ).pack(side=tk.RIGHT)

        stage_frame = tk.Frame(self.root, bg="#1E293B", padx=1, pady=1)
        stage_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=4)

        self.canvas = tk.Canvas(stage_frame, bg=self.BG_CANVAS, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", lambda _e: self.draw())

        ctrl = tk.Frame(self.root, bg=self.BG_PANEL, padx=16, pady=12)
        ctrl.pack(fill=tk.X, padx=14, pady=(8, 12))

        btn_frame = tk.Frame(ctrl, bg=self.BG_PANEL)
        btn_frame.pack(side=tk.LEFT)

        self.btn_start = tk.Button(
            btn_frame,
            text="▶ เริ่มจัดเรียง",
            command=self.start_sort,
            bg="#0284C7",
            fg="#FFFFFF",
            activebackground="#0EA5E9",
            activeforeground="#FFFFFF",
            font=("Leelawadee UI", 10, "bold"),
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_pause = tk.Button(
            btn_frame,
            text="⏸ พัก",
            command=self.toggle_pause,
            bg="#334155",
            fg="#FCD34D",
            font=("Leelawadee UI", 10, "bold"),
            relief=tk.FLAT,
            padx=12,
            pady=6,
            state=tk.DISABLED,
            cursor="hand2",
        )
        self.btn_pause.pack(side=tk.LEFT, padx=4)

        self.btn_step = tk.Button(
            btn_frame,
            text="⏭ ทีละขั้น",
            command=self.step_once_manual,
            bg="#1E293B",
            fg=self.FG_TEXT,
            font=("Leelawadee UI", 10),
            relief=tk.FLAT,
            padx=12,
            pady=6,
            cursor="hand2",
        )
        self.btn_step.pack(side=tk.LEFT, padx=4)

        self.btn_shuffle = tk.Button(
            btn_frame,
            text="🔀 สุ่มใหม่",
            command=self.reset_array,
            bg="#1E293B",
            fg=self.FG_TEXT,
            font=("Leelawadee UI", 10),
            relief=tk.FLAT,
            padx=12,
            pady=6,
            cursor="hand2",
        )
        self.btn_shuffle.pack(side=tk.LEFT, padx=4)

        opts = tk.Frame(ctrl, bg=self.BG_PANEL)
        opts.pack(side=tk.RIGHT)

        self.combo_base = ttk.Combobox(
            opts,
            textvariable=self.base_var,
            values=["Base 4 (Video)", "Base 10 (Decimal)", "Base 2 (Binary)", "Base 8 (Octal)"],
            state="readonly",
            width=16,
        )
        self.combo_base.pack(side=tk.LEFT, padx=(0, 8))
        self.combo_base.bind("<<ComboboxSelected>>", lambda _e: self.reset_array())

        tk.Label(opts, text="N:", bg=self.BG_PANEL, fg=self.FG_MUTED, font=("Consolas", 10)).pack(
            side=tk.LEFT, padx=(4, 2)
        )
        self.scale_n = tk.Scale(
            opts,
            from_=16,
            to=240,
            resolution=8,
            orient=tk.HORIZONTAL,
            variable=self.n_var,
            command=lambda _v: self.reset_array() if not self.running else None,
            bg=self.BG_PANEL,
            fg=self.FG_TEXT,
            troughcolor="#0B101B",
            highlightthickness=0,
            length=105,
        )
        self.scale_n.pack(side=tk.LEFT, padx=(0, 8))

        tk.Label(opts, text="ms:", bg=self.BG_PANEL, fg=self.FG_MUTED, font=("Consolas", 10)).pack(
            side=tk.LEFT, padx=(4, 2)
        )
        tk.Scale(
            opts,
            from_=1,
            to=80,
            orient=tk.HORIZONTAL,
            variable=self.delay_var,
            bg=self.BG_PANEL,
            fg=self.FG_TEXT,
            troughcolor="#0B101B",
            highlightthickness=0,
            length=95,
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.combo_pat = ttk.Combobox(
            opts,
            textvariable=self.pattern_var,
            values=["Random", "Reversed", "Nearly Sorted", "Pyramid"],
            state="readonly",
            width=12,
        )
        self.combo_pat.pack(side=tk.LEFT, padx=(0, 8))
        self.combo_pat.bind("<<ComboboxSelected>>", lambda _e: self.reset_array())

        if HAS_WINSOUND:
            tk.Checkbutton(
                opts,
                text="🔊 Sound",
                variable=self.sound_var,
                bg=self.BG_PANEL,
                fg=self.FG_TEXT,
                selectcolor="#0B101B",
                activebackground=self.BG_PANEL,
                activeforeground=self.FG_TEXT,
                font=("Leelawadee UI", 9),
            ).pack(side=tk.LEFT)

    def _compute_passes(self, max_val: int, base: int) -> int:
        p, e = 0, 1
        while max_val // e > 0:
            p += 1
            e *= base
        return max(1, p)

    def _generate_data(self, n: int, pattern: str) -> list[int]:
        arr = list(range(1, n + 1))
        if pattern == "Random":
            random.shuffle(arr)
        elif pattern == "Reversed":
            arr.reverse()
        elif pattern == "Nearly Sorted":
            for _ in range(max(2, n // 12)):
                i = random.randrange(n)
                j = min(n - 1, max(0, i + random.randint(-4, 4)))
                arr[i], arr[j] = arr[j], arr[i]
        elif pattern == "Pyramid":
            out = [0] * n
            l, r = 0, n - 1
            for idx, val in enumerate(arr):
                if idx % 2 == 0:
                    out[l] = val
                    l += 1
                else:
                    out[r] = val
                    r -= 1
            return out
        return arr

    def reset_array(self) -> None:
        if self.after_id is not None:
            self.root.after_cancel(self.after_id)
            self.after_id = None

        self.running = False
        self.paused = False
        self.gen = None
        self.reads = 0
        self.writes = 0
        self.current_pass = 0
        self.current_exp = 1
        self.highlight_amber = -1
        self.highlight_red = -1
        self.bucket_bounds = []
        self.verified_count = 0
        self.phase_text = "READY"

        n = self.n_var.get()
        base = self._get_base()
        self.arr = self._generate_data(n, self.pattern_var.get())
        self.total_passes = self._compute_passes(n, base)

        self.btn_start.config(state=tk.NORMAL, text="▶ เริ่มจัดเรียง")
        self.btn_pause.config(state=tk.DISABLED, text="⏸ พัก")
        self.scale_n.config(state=tk.NORMAL)
        self.combo_base.config(state="readonly")
        self.draw()

    def _radix_sort_steps(self):
        n = len(self.arr)
        base = self._get_base()
        max_val = max(self.arr)
        self.total_passes = self._compute_passes(max_val, base)
        exp = 1

        for p in range(1, self.total_passes + 1):
            self.current_pass = p
            self.current_exp = exp
            self.bucket_bounds = []
            buckets: list[list[int]] = [[] for _ in range(base)]

            # Phase 1: Scan & distribute into digit buckets
            for i in range(n):
                val = self.arr[i]
                digit = (val // exp) % base
                self.reads += 1
                buckets[digit].append(val)
                self.highlight_amber = i
                self.highlight_red = -1
                self.phase_text = f"PASS {p}/{self.total_passes} SCAN  val={val} -> bucket {digit}"
                if self.sound_var.get():
                    self.beeper.play(val, n)
                yield "step"

            # Compute bucket visual bands
            cur = 0
            bounds = []
            for d in range(base):
                bounds.append((cur, cur + len(buckets[d]), d))
                cur += len(buckets[d])
            self.bucket_bounds = bounds

            # Phase 2: Gather back stably from bucket 0 -> base-1
            write_idx = 0
            for d in range(base):
                for val in buckets[d]:
                    self.arr[write_idx] = val
                    self.writes += 1
                    self.reads += 1
                    self.highlight_amber = -1
                    self.highlight_red = write_idx
                    self.phase_text = f"PASS {p}/{self.total_passes} GATHER bucket {d} -> idx {write_idx}"
                    if self.sound_var.get():
                        self.beeper.play(val, n)
                    write_idx += 1
                    yield "step"

            exp *= base

        # Green verification sweep
        self.highlight_amber = -1
        self.bucket_bounds = []
        for idx in range(n):
            self.verified_count = idx + 1
            self.highlight_red = idx
            self.phase_text = f"VERIFYING  {idx + 1} / {n}  (0 Comparisons)"
            if self.sound_var.get():
                self.beeper.play(self.arr[idx], n)
            yield "verify"

        self.highlight_red = -1
        self.phase_text = f"SORTED ✓  ({self.total_passes} passes, 0 comparisons)"

    def draw(self) -> None:
        c = self.canvas
        c.delete("all")
        w = max(100, c.winfo_width())
        h = max(100, c.winfo_height())
        n = len(self.arr)
        if n == 0:
            return

        pad_x = 16
        top_pad = 66
        bot_pad = 14
        draw_w = w - 2 * pad_x
        draw_h = h - top_pad - bot_pad
        bar_w = draw_w / n
        gap = 1.0 if bar_w > 5 else 0.0

        # Draw bucket bands during Gather phase
        for start, end, digit in self.bucket_bounds:
            if end <= start:
                continue
            bg_col, top_col = self.BUCKET_BANDS[digit % len(self.BUCKET_BANDS)]
            x0 = pad_x + start * bar_w
            x1 = pad_x + end * bar_w
            c.create_rectangle(x0, top_pad - 6, x1, h - bot_pad, fill=bg_col, outline="")
            c.create_rectangle(x0, top_pad - 9, max(x0 + 1, x1 - 1), top_pad - 6, fill=top_col, outline="")

        # Draw bars
        for i, val in enumerate(self.arr):
            bh = max(2.0, (val / n) * draw_h)
            x0 = pad_x + i * bar_w
            x1 = max(x0 + 1.0, x0 + bar_w - gap)
            y0 = h - bot_pad - bh
            y1 = h - bot_pad

            color = self.COLOR_BAR
            if i < self.verified_count:
                color = self.COLOR_GREEN
            if i == self.highlight_amber:
                color = self.COLOR_AMBER
            if i == self.highlight_red:
                color = self.COLOR_RED
            c.create_rectangle(x0, y0, x1, y1, fill=color, outline="")

        base = self._get_base()
        hud1 = (
            f"Radix Sort (LSD, Base {base})  |  N = {n}  |  Pass = {self.current_pass}/{self.total_passes}  "
            f"|  Comparisons = 0  |  Reads = {self.reads:,}  |  Writes = {self.writes:,}"
        )
        hud2 = f"exp = {self.current_exp:,}  |  Delay = {self.delay_var.get()} ms  |  {self.phase_text}"
        c.create_text(18, 16, anchor=tk.NW, text=hud1, fill="#FFFFFF", font=("Consolas", 10, "bold"))
        c.create_text(18, 37, anchor=tk.NW, text=hud2, fill="#38BDF8", font=("Consolas", 9))

    def start_sort(self) -> None:
        if self.verified_count == len(self.arr):
            self.reset_array()
        if self.gen is None:
            self.gen = self._radix_sort_steps()
        self.running = True
        self.paused = False
        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL, text="⏸ พัก")
        self.scale_n.config(state=tk.DISABLED)
        self.combo_base.config(state=tk.DISABLED)
        self._tick()

    def toggle_pause(self) -> None:
        if not self.running:
            return
        self.paused = not self.paused
        if self.paused:
            if self.after_id is not None:
                self.root.after_cancel(self.after_id)
                self.after_id = None
            self.btn_pause.config(text="▶ ทำต่อ")
            self.phase_text = "PAUSED"
            self.draw()
        else:
            self.btn_pause.config(text="⏸ พัก")
            self._tick()

    def step_once_manual(self) -> None:
        if self.running and not self.paused:
            self.toggle_pause()
        if self.verified_count == len(self.arr):
            self.reset_array()
        if self.gen is None:
            self.gen = self._radix_sort_steps()
            self.scale_n.config(state=tk.DISABLED)
            self.combo_base.config(state=tk.DISABLED)
        self._advance()

    def _advance(self) -> str | None:
        try:
            mode = next(self.gen)
            self.draw()
            return mode
        except StopIteration:
            self.running = False
            self.paused = False
            self.gen = None
            self.btn_start.config(state=tk.NORMAL, text="▶ เริ่มใหม่")
            self.btn_pause.config(state=tk.DISABLED, text="⏸ พัก")
            self.scale_n.config(state=tk.NORMAL)
            self.combo_base.config(state="readonly")
            self.draw()
            return None

    def _tick(self) -> None:
        if not self.running or self.paused:
            return
        mode = self._advance()
        if mode is None:
            return
        wait = 8 if mode == "verify" else max(1, self.delay_var.get())
        self.after_id = self.root.after(wait, self._tick)


def main() -> None:
    root = tk.Tk()
    RadixSortVisualizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
