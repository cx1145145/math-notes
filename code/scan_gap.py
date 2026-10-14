# -*- coding: utf-8 -*-
"""截断值 p_d(x) 的取值"缺口"扫描。

用法示例：
    python scan_gap.py --ds 1,2,3,6,12 --rs 100,1000,10000,100000 --cap 100000
    python scan_gap.py --ds 2 --rs 1000000,3000000,10000000 --cap 1000000

两个命题必须分开看，否则结论会错：
  * 「M_d = B」      —— 每个非 3 倍数都被某个 p_d(x) 取到（本程序的"最小缺口"就是它）；
  * 「M_d + 3Z = B」 —— 每类模 3 余 1、余 2 的非 3 倍数被取到（论文定理 5.3 已证）。
后者只要求两个剩余类各出现一次，前者要求逐个命中，难度完全不同。

判据（按用户要求）：r 必须跨 3~5 个数量级。若"最小缺口"一路上升 → 数值支持；
若长期停滞 → M_d = B 出问题，停滞点即"取不到"的候选反例。
"""
import argparse
import sys
import time


def p_d(x, d):
    """按定义求 p_d(x)：b 首次达到 d*n(x) 时停止（该步必为除以 2）。返回 (p, q)。"""
    thr = d * x.bit_length()
    b = 0
    q = 0
    z = x
    while b < thr:
        if z & 1:
            z = 3 * z + 1
            q += 1
        else:
            z >>= 1
            b += 1
    return z, q


def gap_report(attained, cap):
    """返回 (最小的未被取到的非 3 倍数, 缺口个数, 前 15 个缺口)。"""
    miss = [m for m in range(1, cap + 1) if m % 3 and not attained[m]]
    return (miss[0] if miss else None), len(miss), miss[:15]


def scan(ds, rs, cap, quiet=False):
    rs = sorted(set(rs))
    rmax = rs[-1]
    everything = {}
    for d in ds:
        attained = bytearray(cap + 1)
        res_min = {}                       # 剩余类 1/2 mod 3 -> 取到的最小值
        checkpoints = set(rs)
        t0 = time.time()
        rows = []
        c12 = 0                            # p_d(x) ∈ {1,2} 的个数（截断落在 1→4→2→1 循环内）
        for x in range(1, rmax + 1):
            p, _q = p_d(x, d)
            if p <= 2:
                c12 += 1
            if p <= cap and p % 3:
                attained[p] = 1
                r3 = p % 3
                if r3 not in res_min or p < res_min[r3]:
                    res_min[r3] = p
            if x in checkpoints:
                sm, n, head = gap_report(attained, cap)
                t = time.time() - t0
                rows.append({"r": x, "smallest": sm, "count": n, "head": head, "t": t,
                             "res": dict(res_min), "c12": c12})
                if not quiet:
                    print("  d=%-3d r=%-10d 最小缺口=%-9s 缺口数=%-7d 余1=%-5s 余2=%-5s "
                          "p∈{1,2}占比=%4.1f%% 用时%5.1fs"
                          % (d, x, "无(≤%d)" % cap if sm is None else sm, n,
                             res_min.get(1, "-"), res_min.get(2, "-"),
                             100.0 * c12 / x, t))
                    sys.stdout.flush()
        everything[d] = rows
    return everything


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ds", default="2", help="倍数参数 d，逗号分隔")
    ap.add_argument("--rs", default="100,1000,10000,100000", help="遍历上界 r，逗号分隔")
    ap.add_argument("--cap", type=int, default=100000, help="只统计不超过 cap 的取值")
    ap.add_argument("--head", type=int, default=3, help="每个 d 打印前几个 r 的缺口明细")
    a = ap.parse_args()
    ds = [int(s) for s in a.ds.split(",") if s.strip()]
    rs = [int(s) for s in a.rs.split(",") if s.strip()]
    print("扫描 p_d(x)（x ≤ r，只统计 p ≤ %d 的非 3 倍数）" % a.cap)
    print("缺失点 = 猜想「M_d = B = {非 3 倍数}」的反例候选")
    print("-" * 96)
    everything = scan(ds, rs, a.cap)
    print("-" * 96)
    for d, rows in everything.items():
        print("\n【d = %d】" % d)
        for row in rows:
            print("  r=%-10d 最小缺口=%-9s 缺口数=%-7d 前几个缺口: %s"
                  % (row["r"], "无(≤%d)" % a.cap if row["smallest"] is None else row["smallest"],
                     row["count"], row["head"][:10]))
        rows2 = rows[1:1 + a.head]
        for row in rows2:
            print("  明细(r=%d): %s" % (row["r"], row["head"]))


if __name__ == "__main__":
    main()
