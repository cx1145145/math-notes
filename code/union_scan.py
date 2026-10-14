# -*- coding: utf-8 -*-
"""核验新版论文中的可计算断言。

1) 交叉验证：p_2(11)=4、p_2(7)=10（论文问题 4 的断言）；
2) 检验猜想 conj:13（13 ∉ M_d 对一切 d）：搜索 13 ∈ M_d 的见证元；
3) 问题 1 的诚实数据：d ≤ D、x ≤ r 时，⋃ M_d 里仍缺的最小非 3 倍数。

用法：
    python union_scan.py --r 200000 --D 12 --m 13
"""
import argparse
import sys

sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from scan_gap import p_d          # noqa: E402


def witnesses_for(target, ds, r, limit=5):
    """返回每个 d 下使 p_d(x)=target 的若干 x。"""
    out = {}
    for d in ds:
        got = []
        for x in range(1, r + 1):
            if p_d(x, d)[0] == target:
                got.append(x)
                if len(got) >= limit:
                    break
        out[d] = got
    return out


def union_gaps(ds, r, cap):
    """d ∈ ds、x ≤ r 时，≤ cap 的非 3 倍数里从未被取到的（含最小者）。"""
    attained = bytearray(cap + 1)
    for d in ds:
        for x in range(1, r + 1):
            p, _ = p_d(x, d)
            if p <= cap and p % 3:
                attained[p] = 1
    miss = [m for m in range(1, cap + 1) if m % 3 and not attained[m]]
    return miss


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--r", type=int, default=200000)
    ap.add_argument("--D", type=int, default=12)
    ap.add_argument("--m", type=int, default=13)
    ap.add_argument("--cap", type=int, default=300)
    a = ap.parse_args()
    ds = list(range(1, a.D + 1))

    print("== 1. 交叉验证论文已核验过的值（d=2）==")
    for x in (11, 7, 2, 1):
        p, q = p_d(x, 2)
        print("   p_2(%d) = %d   (q=%d, n=%d)" % (x, p, q, x.bit_length()))

    print("\n== 2. 猜想 conj:13 检验：是否存在 x 使 p_d(x) = %d ==" % a.m)
    w = witnesses_for(a.m, ds, a.r)
    for d in ds:
        mark = "有见证 ⇒ 猜想被否证" if w[d] else "未找到（x ≤ %d）" % a.r
        print("   d=%-3d %-28s %s" % (d, w[d] if w[d] else "", mark))

    print("\n== 3. 问题 1 的数据：⋃_{d≤%d} M_d（x ≤ %d）在 ≤ %d 上仍缺的非 3 倍数 ==" % (a.D, a.r, a.cap))
    miss = union_gaps(ds, a.r, a.cap)
    print("   缺口个数 = %d" % len(miss))
    print("   最小缺口 = %s" % (miss[0] if miss else "无（全部取到）"))
    print("   前 20 个缺口 = %s" % miss[:20])


if __name__ == "__main__":
    main()
