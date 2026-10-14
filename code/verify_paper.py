# -*- coding: utf-8 -*-
"""论文《考拉兹截断：更正与部分结果》附录 A 的可复现验证程序。

运行： python verify_paper.py     （纯标准库，约十几秒）
说明： 本程序按定义 2.1（截断规则）与命题 2.5（集合 A 的成员判据）逐项计算。
       **程序只是核验，不构成证明。**
"""
import math


def n(x):
    """二进制位长（三种写法一致）。"""
    return x.bit_length()


def kappa(m):
    """最小的不小于 m 的 3 的幂。"""
    t = 3
    while t < m:
        t *= 3
    return t


def inA(m):
    """命题 2.5：m ∈ A ⟺ 3∤m 且 2m > 3^κ(m)（严格不等式）。"""
    return m >= 2 and m % 3 != 0 and 2 * m > kappa(m)


def pd(x, d, cap=10 ** 7):
    """定义 2.1：返回 (p_d(x), q_d(x))。"""
    nx = n(x)
    k = q = 0
    while cap > 0:
        cap -= 1
        if x & 1:
            x = 3 * x + 1
            q += 1
        else:
            x >>= 1
            k += 1
            if k == d * nx:
                return x, q
    raise RuntimeError("截断条件未在 cap 步内满足")


def N3(a, b):
    return (b - a) - (b // 3 - a // 3)


def report():
    print("A0  位长三种定义的一致性（x <= 100000）")
    bad = [x for x in range(1, 100001)
           if not (x.bit_length() == math.floor(math.log2(x)) + 1 == math.ceil(math.log2(x + 1)))]
    print("    不一致个数：%d  %s" % (len(bad), "OK" if not bad else "FAIL " + str(bad[:5])))

    L = 3 ** 12
    pref = [0] * (L + 1)
    for m in range(1, L + 1):
        pref[m] = pref[m - 1] + (1 if inA(m) else 0)

    print("A1  闭式 A(3^K) = A(3^K-1) = (3^K-1)/2")
    ok = True
    for K in range(1, 13):
        a, b = pref[3 ** K], pref[3 ** K - 1]
        e = (3 ** K - 1) // 2
        ok &= (a == e and b == e)
        if K <= 6 or K >= 10:
            print("    K=%2d  A(3^K)=%7d  A(3^K-1)=%7d  期望 %7d  %s"
                  % (K, a, b, e, "OK" if (a == e and b == e) else "FAIL"))
    print("    K=1..12 全部 %s" % ("OK" if ok else "FAIL"))

    print("A2  分层计数 #A_k = 3^(k-1)")
    for k in range(1, 11):
        c = sum(1 for m in range(3 ** (k - 1) + 1, 3 ** k) if inA(m))
        print("    k=%2d  #A_k=%6d  期望 %6d  %s" % (k, c, 3 ** (k - 1), "OK" if c == 3 ** (k - 1) else "FAIL"))

    print("A3  每层 = (3^k/2, 3^k) 内非 3 倍数")
    for k in range(1, 11):
        c1 = sum(1 for m in range(3 ** (k - 1) + 1, 3 ** k) if inA(m))
        c2 = N3(3 ** k // 2, 3 ** k)
        print("    k=%2d  逐层=%6d  公式=%6d  %s" % (k, c1, c2, "OK" if c1 == c2 else "FAIL"))

    print("A4  p_1(x) ∈ A（x <= 200000）")
    bad = [x for x in range(1, 200001) if not inA(pd(x, 1)[0])]
    print("    反例数：%d  %s" % (len(bad), "OK" if not bad else "FAIL " + str(bad[:5])))

    print("A5  定理 3.3：3 ∤ p_d(x)（x <= 2000, d <= 12）")
    bad = sum(1 for d in range(1, 13) for x in range(1, 2001) if pd(x, d)[0] % 3 == 0)
    print("    反例数：%d  %s" % (bad, "OK" if bad == 0 else "FAIL"))

    print("A6  命题 2.2：q_d(x) >= 1（x <= 5000, d <= 12）")
    bad = sum(1 for d in range(1, 13) for x in range(1, 5001) if pd(x, d)[1] == 0)
    print("    反例数：%d  %s" % (bad, "OK" if bad == 0 else "FAIL"))

    print("A7  定理 5.3：M_d ∩ {1,2,3}（1 ∈ M_d 当且仅当 d >= 2；2 恒成立；3 恒不出现）")
    for d in range(1, 13):
        S = sorted({pd(x, d)[0] for x in range(1, 4001)} & {1, 2, 3})
        print("    d=%2d  %s" % (d, S))

    print("A8  见证元：p_d(2)=2（一切 d）；1∈M_d：d=3 用 x=11，奇 d>=5 用 x=7")
    for d in range(1, 15):
        w = 11 if d == 3 else 7
        print("    d=%2d  p_d(2)=%d   p_d(%d)=%d" % (d, pd(2, d)[0], w, pd(w, d)[0]))

    print("A9  比商极值（3^7 <= N <= 3^12）")
    lo = min((pref[N] / N, N) for N in range(3 ** 7, L + 1))
    hi = max((pref[N] / N, N) for N in range(3 ** 7, L + 1))
    print("    min A(N)/N = %.6f (N=%d)   期望趋于 1/3" % lo)
    print("    max A(N)/N = %.6f (N=%d)   期望趋于 1/2" % hi)
    print("    N=2*3^(K-1) 处（K=12）：%.6f   期望趋于 5/12" % (pref[2 * 3 ** 11] / (2 * 3 ** 11)))

    print("A10 定理 4.3：A_1 \\ (3A+1) 与 A_2 \\ (3A+2) 的最小元素")
    A = [m for m in range(2, 3000) if inA(m)]
    A1 = [m for m in A if m % 3 == 1]
    A2 = [m for m in A if m % 3 == 2]
    t1 = {3 * a + 1 for a in A}
    t2 = {3 * a + 2 for a in A}
    print("    A_1 \\ (3A+1) 最小几个：%s" % [m for m in A1 if m not in t1][:6])
    print("    A_2 \\ (3A+2) 最小几个：%s" % [m for m in A2 if m not in t2][:6])


if __name__ == "__main__":
    report()
