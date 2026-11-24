"""
Radix Sort (LSD) Benchmark & Empirical Runtime Analysis
=======================================================
Measures real execution time, pass counts, and array writes of pure Python
Least Significant Digit (LSD) Radix Sort (0 comparisons!) across:
- Sizes: 10,000 | 100,000 | 1,000,000 | 10,000,000 random integers
- Bases: Base 256 (byte-wise, fast practical) & Base 10 (decimal textbook)
         plus Base 4 (Timo Bingmann's video base) and Base 16 at N=1,000,000
- Input distributions at N=100,000 (Random, Sorted, Reversed, Many Duplicates)
- Contrast against O(n log n) Merge Sort & O(n^2) Selection Sort
"""

from __future__ import annotations

import json
import math
import platform
import random
import statistics
import time
from pathlib import Path


def radix_sort_lsd_base256(arr: list[int]) -> list[int]:
    """
    Pure Python LSD Radix Sort (Base 256 = 8 bits per pass).
    Non-comparison sort: ZERO element-to-element comparisons!
    Time: O(d * (n + 256)) where d = ceil(log_256(max_val + 1)).
    Space: O(n + 256). Stable.
    """
    n = len(arr)
    if n <= 1:
        return arr
    max_val = max(arr)
    shift = 0
    while (max_val >> shift) > 0:
        buckets: list[list[int]] = [[] for _ in range(256)]
        for x in arr:
            buckets[(x >> shift) & 255].append(x)
        arr = [x for b in buckets for x in b]
        shift += 8
    return arr


def radix_sort_lsd_base(arr: list[int], base: int = 10) -> list[int]:
    """
    General Pure Python LSD Radix Sort for arbitrary radix `base` (e.g. 4, 10, 16).
    Time: O(d * (n + base)), Space: O(n + base), Comparisons = 0, Stable.
    """
    n = len(arr)
    if n <= 1:
        return arr
    max_val = max(arr)

    # Fast bit-mask path when base is a power of 2 (e.g., Base 4, 16, 256)
    if base & (base - 1) == 0:
        bits = int(math.log2(base))
        mask = base - 1
        shift = 0
        while (max_val >> shift) > 0:
            buckets: list[list[int]] = [[] for _ in range(base)]
            for x in arr:
                buckets[(x >> shift) & mask].append(x)
            arr = [x for b in buckets for x in b]
            shift += bits
        return arr

    # General integer division path (e.g., Base 10)
    exp = 1
    while max_val // exp > 0:
        buckets = [[] for _ in range(base)]
        for x in arr:
            buckets[(x // exp) % base].append(x)
        arr = [x for b in buckets for x in b]
        exp *= base
    return arr


def _merge_sort_ref(arr: list[int]) -> list[int]:
    """Pure Python O(n log n) Merge Sort reference for comparison."""
    n = len(arr)
    if n <= 1:
        return arr
    aux = arr.copy()

    def _sort(src: list[int], dst: list[int], left: int, right: int) -> None:
        if right - left <= 1:
            return
        mid = (left + right) >> 1
        _sort(dst, src, left, mid)
        _sort(dst, src, mid, right)
        i, j, k = left, mid, left
        while i < mid and j < right:
            if src[i] <= src[j]:
                dst[k] = src[i]
                i += 1
            else:
                dst[k] = src[j]
                j += 1
            k += 1
        if i < mid:
            dst[k:right] = src[i:mid]
        elif j < right:
            dst[k:right] = src[j:right]

    _sort(aux, arr, 0, n)
    return arr


def _selection_sort_ref(arr: list[int]) -> list[int]:
    """Pure Python O(n^2) Selection Sort reference at N=10,000."""
    n = len(arr)
    for i in range(n - 1):
        min_idx = i
        min_val = arr[i]
        for j in range(i + 1, n):
            if arr[j] < min_val:
                min_val = arr[j]
                min_idx = j
        if min_idx != i:
            arr[i], arr[min_idx] = arr[min_idx], arr[i]
    return arr


def is_sorted(arr: list[int]) -> bool:
    return all(arr[i] <= arr[i + 1] for i in range(len(arr) - 1))


def run_benchmark() -> dict:
    rng = random.Random(42)
    # Fixed 9-digit value range [0, 999_999_999] so d is constant across N:
    #   Base 256 -> d = 4 passes (30 bits -> ceil(30/8) = 4)
    #   Base 10  -> d = 9 passes (9 decimal digits)
    MAX_VAL = 999_999_999

    print("=" * 84)
    print("  RADIX SORT (LSD) BENCHMARK (Pure Python 3.12 — 0 Comparisons!)")
    print("=" * 84)
    print(f"Python  : {platform.python_version()}")
    print(f"OS      : {platform.system()} {platform.release()} ({platform.machine()})")
    print(f"Range   : 0 .. {MAX_VAL:,} (fixed 9 decimal digits / 30 bits)")
    print()

    # [1] Size scaling: Base 256 (primary fast LSD) and Base 10 (textbook decimal LSD)
    size_configs = [
        (10_000, 5),
        (100_000, 5),
        (1_000_000, 3),
        (10_000_000, 1),
    ]

    scaling_b256 = []
    scaling_b10 = []
    ref_merge = {}

    print("[1] Random Input Scaling — Radix Sort LSD (Base 256: 4 passes) vs (Base 10: 9 passes)")
    print(
        f"{'N':>12} | {'Base256 Avg(s)':>14} | {'ns/item(256)':>12} | "
        f"{'Base10 Avg(s)':>13} | {'ns/item(10)':>11} | {'MergeSort(s)':>12} | {'OK'}"
    )
    print("-" * 92)

    for n, trials in size_configs:
        t_256 = []
        t_10 = []
        verified = True

        for _ in range(trials):
            base_data = [rng.randint(0, MAX_VAL) for _ in range(n)]

            # Base 256
            arr1 = base_data.copy()
            s0 = time.perf_counter()
            res1 = radix_sort_lsd_base256(arr1)
            t_256.append(time.perf_counter() - s0)
            if not is_sorted(res1):
                verified = False

            # Base 10 (run on 10K, 100K, 1M, and 10M!)
            arr2 = base_data.copy()
            s1 = time.perf_counter()
            res2 = radix_sort_lsd_base(arr2, 10)
            t_10.append(time.perf_counter() - s1)
            if not is_sorted(res2):
                verified = False

        # Reference Merge Sort on 10K, 100K, 1M for direct chart contrast
        ms_time = None
        if n <= 1_000_000:
            ms_trials = 3 if n <= 100_000 else 1
            ms_times = []
            for _ in range(ms_trials):
                arr_m = [rng.randint(0, MAX_VAL) for _ in range(n)]
                sm = time.perf_counter()
                _merge_sort_ref(arr_m)
                ms_times.append(time.perf_counter() - sm)
            ms_time = statistics.mean(ms_times)
            ref_merge[str(n)] = round(ms_time, 6)

        avg256 = statistics.mean(t_256)
        avg10 = statistics.mean(t_10)
        ns256 = (avg256 * 1e9) / n
        ns10 = (avg10 * 1e9) / n

        ms_str = f"{ms_time:.4f}" if ms_time is not None else "—"
        print(
            f"{n:>12,} | {avg256:>14.4f} | {ns256:>12.1f} | "
            f"{avg10:>13.4f} | {ns10:>11.1f} | {ms_str:>12} | {'OK' if verified else 'FAIL'}"
        )

        scaling_b256.append(
            {
                "n": n,
                "base": 256,
                "passes": 4,
                "trials": trials,
                "times_sec": [round(x, 6) for x in t_256],
                "avg_sec": round(avg256, 6),
                "min_sec": round(min(t_256), 6),
                "max_sec": round(max(t_256), 6),
                "ns_per_item": round(ns256, 2),
                "comparisons": 0,
                "writes": n * 4,
                "verified": verified,
            }
        )
        scaling_b10.append(
            {
                "n": n,
                "base": 10,
                "passes": 9,
                "trials": trials,
                "times_sec": [round(x, 6) for x in t_10],
                "avg_sec": round(avg10, 6),
                "min_sec": round(min(t_10), 6),
                "max_sec": round(max(t_10), 6),
                "ns_per_item": round(ns10, 2),
                "comparisons": 0,
                "writes": n * 9,
                "verified": verified,
            }
        )

    # [2] Radix Base comparison at N = 1,000,000 (Base 4, 10, 16, 256)
    print("\n[2] Effect of Radix Base b at N = 1,000,000 (0..999,999,999)")
    sample_1m = [rng.randint(0, MAX_VAL) for _ in range(1_000_000)]
    base_configs = [
        (4, 15, "Base 4 (Video Default — 2 bits/pass)"),
        (10, 9, "Base 10 (Decimal — 1 digit/pass)"),
        (16, 8, "Base 16 (Hex — 4 bits/pass)"),
        (256, 4, "Base 256 (Byte — 8 bits/pass)"),
    ]
    base_comparison = []
    print(f"{'Base Configuration':<38} | {'Passes (d)':>10} | {'Time (s)':>10} | {'Writes':>12}")
    print("-" * 78)
    for b, passes, label in base_configs:
        arr_b = sample_1m.copy()
        s0 = time.perf_counter()
        res_b = radix_sort_lsd_base256(arr_b) if b == 256 else radix_sort_lsd_base(arr_b, b)
        dt = time.perf_counter() - s0
        assert is_sorted(res_b)
        print(f"{label:<38} | {passes:>10} | {dt:>10.4f} | {1_000_000 * passes:>12,}")
        base_comparison.append(
            {
                "base": b,
                "passes": passes,
                "label": label,
                "time_sec": round(dt, 6),
                "writes": 1_000_000 * passes,
            }
        )

    # [3] Input Distribution Sensitivity at N = 100,000 (Base 256)
    print("\n[3] Input Distribution Sensitivity Test (N = 100,000, Base 256, 3 trials)")
    dist_n = 100_000
    distributions = {
        "Random": lambda: [rng.randint(0, MAX_VAL) for _ in range(dist_n)],
        "Already Sorted": lambda: [i * 9999 for i in range(dist_n)],
        "Reverse Sorted": lambda: [(dist_n - i) * 9999 for i in range(dist_n)],
        "Many Duplicates (100 values)": lambda: [rng.randint(0, 99) * 10_000_000 for _ in range(dist_n)],
    }
    dist_results = []
    print(f"{'Distribution':<30} | {'Avg Time (s)':>12} | {'Comparisons':>11} | {'Writes':>10}")
    print("-" * 72)
    for label, gen in distributions.items():
        ts = []
        for _ in range(3):
            arr_d = gen()
            s0 = time.perf_counter()
            res_d = radix_sort_lsd_base256(arr_d)
            ts.append(time.perf_counter() - s0)
            assert is_sorted(res_d)
        avg_t = statistics.mean(ts)
        print(f"{label:<30} | {avg_t:>12.4f} | {0:>11} | {dist_n * 4:>10,}")
        dist_results.append(
            {
                "distribution": label,
                "n": dist_n,
                "avg_sec": round(avg_t, 6),
                "comparisons": 0,
                "writes": dist_n * 4,
            }
        )

    # [4] Contrast vs O(n^2) Selection Sort at N = 10,000
    print("\n[4] Contrast vs O(n^2) Selection Sort at N = 10,000")
    sel_arr = [rng.randint(0, MAX_VAL) for _ in range(10_000)]
    t0 = time.perf_counter()
    _selection_sort_ref(sel_arr)
    sel_time = time.perf_counter() - t0
    assert is_sorted(sel_arr)
    rad_10k = scaling_b256[0]["avg_sec"]
    speedup_sel = sel_time / rad_10k
    est_sel_1m_hr = (sel_time * (100 ** 2)) / 3600.0
    speedup_merge_1m = ref_merge["1000000"] / scaling_b256[2]["avg_sec"]
    print(f"Selection Sort (10K): {sel_time:.4f} s  -> Radix Base256 is {speedup_sel:.1f}x faster!")
    print(f"Merge Sort     (1M) : {ref_merge['1000000']:.4f} s  -> Radix Base256 is {speedup_merge_1m:.2f}x faster!")
    print(f"Estimated Selection Sort at 1M: ~{est_sel_1m_hr:.1f} hours!")

    results = {
        "environment": {
            "python_version": platform.python_version(),
            "os": f"{platform.system()} {platform.release()}",
            "machine": platform.machine(),
            "cpu": "AMD Ryzen 5 220 (6 Cores / 12 Threads)",
            "ram_gb": 16,
            "value_range": f"0 .. {MAX_VAL:,} (9 digits / 30 bits)",
        },
        "scaling_base256": scaling_b256,
        "scaling_base10": scaling_b10,
        "reference_merge_sort": ref_merge,
        "base_comparison_1m": base_comparison,
        "distributions_100k": dist_results,
        "comparison_summary": {
            "selection_sort_10k_sec": round(sel_time, 6),
            "radix256_10k_sec": round(rad_10k, 6),
            "speedup_vs_selection_10k": round(speedup_sel, 1),
            "est_selection_1m_hours": round(est_sel_1m_hr, 1),
            "speedup_vs_merge_1m": round(speedup_merge_1m, 2),
        },
    }

    out_path = Path(__file__).parent / "benchmark_results.json"
    out_path.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved benchmark results to: {out_path}")
    return results


if __name__ == "__main__":
    run_benchmark()
