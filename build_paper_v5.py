# -*- coding: utf-8 -*-
"""论文《考拉兹截断：更正与部分结果》第五版生成器（LaTeX 源 + Word 版）。

唯一的正文来源是下面的 CONTENT 列表；本脚本由它同时产出 .tex 与 .docx（.pdf 由
.docx 经 LibreOffice 转换），从而保证三个版本内容一致。

第五版修订（并入第四、五轮外部核对，全部经程序核验）：
  n(x) 的等价定义；q>=1 恒成立（删除"无奇步"情形）；情形 (i) 的模 3 值为 2；
  3A+1 / 3A+2 的严格性反例；1∈M_d 分 d=3（x=11）与奇数 d>=5（x=7）；
  A(3^K)=(3^K-1)/2（K<=12）；密度上下极限的取值点。
第五版后续清理（结构 + 数值口径）：
  1. 定理/定义/命题/注/猜想一律交给 LaTeX 自动编号，标题不再手写编号；
  2. 定理 3.1 的引理独立成 lemma 环境（\\label{lem:ab}）并加 \\label；
  3. 所有证明内容完整落在 proof 环境中；引理不再嵌套在 proof 内；
  4. 交叉引用改用 \\T{label}：LaTeX 侧展开为 \\ref，Word 侧展开为脚本算出的编号；
  5. 定理 5.1 中 A(3^{K-1}) 改为更严谨的 A(3^{K-1}-1)，层长度改为 3^j/2；
  6. 注 5.2 与表 A.1 的 A9 栏改用程序实际输出（min 0.333232 @ N=3280）；
  7. 表 A.1 与程序输出 A0..A10 逐行对应；表 A.2 补齐 K=7,9,11；
  8. 附录用 \\appendix + \\numberwithin，参考文献用 thebibliography。
所有数值断言由 code/verify_paper.py 核验。程序只是核验，不构成证明。
"""
import os
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = os.path.dirname(os.path.abspath(__file__))

# 定理类环境的 Kind -> 中文名（共享一个计数器，按 section 重置）
NAMES = {"thm": "定理", "lem": "引理", "prop": "命题", "cor": "推论",
         "conj": "猜想", "def": "定义", "rem": "注"}

CONTENT = [
("title", "考拉兹截断：更正与部分结果"),
("author", "据原始手稿《一个猜想》审读、修订与补证（第五版）"),

("abs_h", "摘要"),
("abs", "设 $T$ 为考拉兹映射，$n(x)=\\lfloor\\log_2x\\rfloor+1$ 为 $x$ 的二进制位长。对 $d\\in\\mathbb{Z}^+$ 定义截断算子 $T_d$：自 $x$ 出发沿 $T$ 迭代，"
        "当“除以 2”的次数**首次**达到 $d\\,n(x)$ 时停止（该步必为除以 2），其值记为 $p_d(x)$，令 $M_d=\\{p_d(x):x\\in\\mathbb{Z}^+\\}$。"
        "本文是修正性研究笔记，内容分四类，界限明确。**定理**：(1) 截断条件对任意 $x,d$ 在前 $2d\\,n(x)+1$ 步内被满足（证明不引用考拉兹猜想）；"
        "(2) 对一切 $x,d$ 有 $3\\nmid p_d(x)$，从而 $M_d\\cap3\\mathbb{Z}=\\varnothing$——这否证了手稿的猜想 $M_d+3\\mathbb{Z}=\\mathbb{Z}^+\\ (d\\ge2)$；"
        "(3) 对一切 $d$ 有 $2\\in M_d$，对一切 $d\\ge2$ 有 $1\\in M_d$；$d=1$ 时 $1\\notin M_1$ 未证，但 $22=p_1(25)\\in M_1$ 且 $22\\equiv1\\pmod3$，"
        "故 $1\\in M_1+3\\mathbb{Z}$，于是对一切 $d\\ge1$ 有 $M_d+3\\mathbb{Z}=\\{m:3\\nmid m\\}$；"
        "(4) 对 $A=\\{m\\ge2:3\\nmid m,\\ \\exists k\\ge1,\\ 3^k/2<m<3^k\\}$，有 $A(3^K)=A(3^K-1)=\\frac{3^K-1}{2}$，"
        "故 $\\limsup A(N)/N=1/2$、$\\liminf A(N)/N=1/3$（$A$ 的自然密度不存在，平均密度 $5/12$）；"
        "(5) $3A+1\\subseteq A_1$、$3A+2\\subseteq A_2$ 且两个包含均严格。"
        "**猜想**：$M_1=A$（第 6 节：本文只证其一方向并给出数值支持）。**数值观察**：$M_d\\ (d\\ge2)$ 对非 3 倍数的覆盖率数据。"
        "**更正**：第 1.1 节集中列出本文对初稿与中期版本共九处错误的更正。"),
("kw", "关键词：考拉兹映射；二进制位长；截断算子；模 3 不变量；自然密度；log-周期函数；MSC 11B83, 37P99"),

("h1", "1  引言与更正清单"),
("p", "考拉兹映射为 $T(x)=x/2$（$x$ 偶）、$3x+1$（$x$ 奇）；考拉兹猜想断言每个正整数最终进入循环 $1\\to4\\to2\\to1$，至今未决。"
      "原始手稿《一个猜想》对轨道施加“按二进制位长截断”的规则（定义 \\T{def:trunc}），研究截断值集合 $M_d$，"
      "声称 $M_1$ 等于集合 $A$、猜想 $M_d+3\\mathbb{Z}=\\mathbb{Z}^+\\ (d\\ge2)$、并称不变集 $C=M_1$ 的自然密度不存在而上下极限为 $\\frac12$ 与 $\\frac13$。"),
("p", "本文的结论：手稿的覆盖猜想**为假**（定理 \\T{thm:mod3}、推论 \\T{cor:subset}），其唯一正确的改写 $M_d+3\\mathbb{Z}=\\{m:3\\nmid m\\}$ **为真且可证**（定理 \\T{thm:cover}）；"
      "手稿的密度断言**为真**（定理 \\T{thm:density}）；而 $M_1=A$ 的两个方向本文只给出其中一个（第 6 节），第二个方向仍为猜想。"),
("h2", "1.1  更正清单（初稿与中期版本共九处）"),
("item", "**C1（缩放恒等式）**：初稿断言 $p_d(2^{ds}x)=2^{ds}p_d(x)$。**为假**：$p_1(2)=2$ 而 $2p_1(1)=4$；$d\\ge2$ 时比值一般不是整数。已删除。"),
("item", "**C2（错误举例）**：初稿写 $5=p_1(5)$。**错**：$p_1(5)=2$。$5\\in M_1$ 的见证元是 $x=17$。"),
("item", "**C3（终止性证明）**：初稿用 $x_i<x/2^{\\,i-n(x)}$ 并称“$n(x)$ 步内到达 $1$”。该不等式不成立（$x=7$：$x_4=17>x/2$），且“到达 $1$”即考拉兹猜想。已改用计数论证（定理 \\T{thm:term}）。"),
("item", "**C4（$A$ 的定义）**：必须用**严格**不等式 $3^k/2<m<3^k$；非严格不等式会破坏 $M_1=A$。"),
("item", "**C5（$x=2$ 的奇偶）**：初稿称 $p_d(2)=1$（$d$ 偶）。**错**：第 $2d$ 次除以 2 的序数 $2d$ 恒为偶，故 $p_d(2)=2$ 对一切 $d$。"),
("item", "**C6（计数闭式）**：初稿写成 $A(3^K)=\\frac{3^K\\pm1}{2}$ 等多种互相矛盾的形式。正确为 $A(3^K)=A(3^K-1)=\\frac{3^K-1}{2}$（$K\\le12$ 已核验）。"),
("item", "**C7（$n(x)$ 的写法）**：中期版本曾误写为 $\\lceil\\log_2x\\rceil+1$；它与正确位长不等（例如 $x=3$ 时分别为 3 与 2）。"
        "正确写法是 $\\lfloor\\log_2x\\rfloor+1$，并且它等价于 $\\lceil\\log_2(x+1)\\rceil$，与程序 x.bit\\_length() 一致；本版统一写 $n(x)=\\lfloor\\log_2x\\rfloor+1$。"),
("item", "**C8（“无奇步”情形）**：中期版本在定理 \\T{thm:mod3} 中保留了“轨道中从未出现奇步”的情形。**该情形不可能发生**：若从未出现奇步则 $x=2^{t}$，到 $1$ 只需 $t$ 次除以 2，"
        "而阈值 $d\\,n(x)=d(t+1)>t$，截断条件永远无法满足。已删除该情形。"),
("item", "**C9（$1\\in M_d$ 的见证元）**：中期版本用 $x=7$ 覆盖一切奇数 $d\\ge3$。**对 $d=3$ 不成立**（$p_3(7)=4$，因 $3d=9$ 为奇而轨道第 9 次除以 2 落在 4 上）。"
        "正确做法：$d=3$ 用 $x=11$（$p_3(11)=1$），奇数 $d\\ge5$ 用 $x=7$（$3d$ 为 $\\ge15$ 的奇数，轨道第 $15,17,19,\\dots$ 次除以 2 均落在 1 上）。"),
("h2", "1.2  本文的结论分类"),
("p", "为免混淆，本文对每条陈述标注类别：**定理**（已给出完整证明）、**命题**（已证但较简单）、**猜想**（未证）、**数值观察**（仅有程序核验）。"),

("h1", "2  定义与记号"),
("p", "记号：$\\mathbb{Z}^+$ 为正整数集；$n(x)=\\lfloor\\log_2x\\rfloor+1$（等价于 $x$ 的二进制位数，也等价于 $\\lceil\\log_2(x+1)\\rceil$，$x\\ge1$）；"
      "$\\lfloor\\cdot\\rfloor$、$\\lceil\\cdot\\rceil$ 为下取整与上取整；$3\\nmid m$ 表示 $m$ 不是 3 的倍数；$3\\mathbb{Z}=\\{3c:c\\in\\mathbb{Z}\\}$。"),
("def", "def:trunc", "截断算子",
        ["设 $d\\in\\mathbb{Z}^+$，$x\\in\\mathbb{Z}^+$。自 $x$ 出发沿 $T$ 迭代，记 $a$ 为已执行的 $3u+1$ 型步骤（下称**奇步**）数、$b$ 为已执行的除以 2 的步数。"
         "当 $b$ **首次**达到 $d\\,n(x)$ 时停止（由定理 \\T{thm:term}，该时刻必存在且该步必为除以 2）。此时记",
         ("eq", "$q_d(x)=a,\\qquad p_d(x)=\\text{停止时的值},\\qquad M_d=\\{p_d(x):x\\in\\mathbb{Z}^+\\}.$")]),
("prop", "prop:qge", "$q_d\\ge1$", ["对一切 $x,d$ 有 $q_d(x)\\ge1$（即轨道中至少出现一次奇步）。"]),
("proof", ["反设 $q_d(x)=0$。则轨道全程为除以 2，故 $x=2^{t}$ 且 $n(x)=t+1$，到 $1$ 恰需 $t$ 次除以 2，此后无法继续（$1$ 为奇，下一步必为奇步，与 $q=0$ 矛盾）。"
           "于是可达的 $b$ 值只到 $t$，而阈值 $d\\,n(x)=d(t+1)\\ge t+1>t$，截断条件不可能满足。"]),
("rem", "rem:Ared", "$m\\ge2$ 是冗余条件",
        ["$A$ 的定义中条件 $m\\ge2$ 其实是冗余的：由严格不等式 $3^k/2<m<3^k$ 与 $k\\ge1$ 立即得 $m>3/2$，故 $m\\ge2$ 自动成立，"
         "特别地 $m=1$ 永远不满足（$3/2<1$ 为假）。中期版本曾写“若允许 $m=1$ 则 $1\\in A$（取 $k=1$）”，那是把严格不等式误读为 $3/2\\le1$；此说明由更正 C4 覆盖。"]),
("def", "def:A", "集合 $A$", ["$A=\\{m\\in\\mathbb{Z}^+:3\\nmid m,\\ \\exists k\\ge1,\\ 3^k/2<m<3^k\\}$。"]),
("prop", "prop:kappa", "成员判据",
        ["对 $m\\ge2$ 且 $3\\nmid m$，设 $3^{\\kappa(m)}$ 为最小的 $\\ge m$ 的 3 的幂。则",
         ("eq", "$m\\in A\\iff 2m>3^{\\kappa(m)}\\iff\\exists k:\\ 3^k/2<m<3^k.$")]),
("proof", ["若 $2m>3^{\\kappa(m)}$，取 $k=\\kappa(m)$：由 $m\\le3^{\\kappa(m)}$ 与 $3\\nmid m$ 得 $m<3^{\\kappa(m)}$，故 $3^k/2<m<3^k$。"
           "反之设 $3^k/2<m<3^k$，则 $k\\ge\\kappa(m)$，故 $3^{\\kappa(m)}\\le3^k<2m$。"]),

("h1", "3  终止性与模 3 不变量"),
("lem", "lem:ab", "奇步计数",
        ["记 $a(i)$、$b(i)$ 为前 $i$ 步中奇步与除 2 步的个数（$a(i)+b(i)=i$）。则 $a(i)\\le b(i)+1$。"]),
("proof", ["奇数 $u$ 经 $3u+1$ 得偶数，故紧接奇步之后的必是除 2 步；相邻两个奇步之间至少夹一次除以 2，故 $a(i)\\le b(i)+1$。"]),
("thm", "thm:term", "截断条件恒可达", ["对任意 $x,d\\in\\mathbb{Z}^+$，截断条件在前 $2d\\,n(x)+1$ 步之内被满足。"]),
("proof", ["由引理 \\T{lem:ab}，若 $b(i)=m$ 则 $i\\le2m+1$。因每次除以 2 使 $b$ 增 1、奇步不改变 $b$，$b$ 必取到每个非负整数值；"
           "取 $i_0=\\min\\{i:b(i)=d\\,n(x)\\}$，则 $i_0\\le2d\\,n(x)+1$，且由最小性 $b(i_0-1)=d\\,n(x)-1$，故第 $i_0$ 步为除以 2。"]),
("rem", "rem:noconj", "", ["本证明不使用考拉兹猜想，也不使用轨道有界性。"]),
("h2", "3.1  模 3 不变量"),
("thm", "thm:mod3", "截断值永不整除 3", ["对一切 $x,d\\in\\mathbb{Z}^+$ 有 $3\\nmid p_d(x)$，即 $M_d\\cap3\\mathbb{Z}=\\varnothing$。"]),
("proof", ["记 $p=p_d(x)$，考虑截断的最后一步：它把偶数 $y$ 变为 $y/2=p$。由命题 \\T{prop:qge}（$q_d(x)\\ge1$），最后一步之前的那一步只有两种可能。",
           "- 情形 (i)：前一步是奇步。即 $y=3u+1$（$u$ 为奇数），于是 $p=\\dfrac{3u+1}{2}$。由 $3u+1\\equiv1\\pmod3$ 与 $2^{-1}\\equiv2\\pmod3$ 得 $p\\equiv2\\pmod3$，故 $3\\nmid p$。",
           "- 情形 (ii)：前一步是除 2 步。即 $y=z/2$，其中 $z$ 为截断之前最后一次奇步的输出 $z=3v+1$。设 $y$ 到 $p$ 还做了 $s\\ge1$ 次除以 2，"
           "即 $z=p\\cdot2^{s}$。由 $z\\equiv1\\pmod3$ 与 $\\gcd(2,3)=1$ 得 $p\\equiv2^{-s}\\not\\equiv0\\pmod3$。",
           "两种情形穷尽了最后一步之前的所有可能（“从未出现奇步”不可能，见更正 C8 与命题 \\T{prop:qge}）。故 $3\\nmid p_d(x)$。"]),
("cor", "cor:subset", "手稿猜想的否证",
        ["$M_d+3\\mathbb{Z}\\subseteq\\{m\\in\\mathbb{Z}^+:3\\nmid m\\}\\subsetneq\\mathbb{Z}^+$；特别地 $3\\notin M_d+3\\mathbb{Z}$ 对一切 $d$。"
         "故手稿的猜想“$M_d+3\\mathbb{Z}=\\mathbb{Z}^+\\ (d\\ge2)$”为假。"]),

("h1", "4  $d=1$ 与集合 $A$ 的计数"),
("h2", "4.1  分层与计数闭式"),
("p", "对 $k\\ge1$ 记 $A_k=A\\cap(3^{k-1},3^k)$。例如 $A_1=\\{2\\}$，$A_2=\\{5,7,8\\}$，$A_3=\\{14,16,17,19,20,22,23,25,26\\}$。"),
("thm", "thm:count", "分层计数", ["对一切 $k\\ge1$，$\\#A_k=3^{k-1}$，且 $A(3^K)=\\sum_{k=1}^{K}\\#A_k=\\dfrac{3^K-1}{2}$。"]),
("proof", ["**分层**：由命题 \\T{prop:kappa}，$m\\in A\\iff\\exists k:\\ 3^k/2<m<3^k$。对 $m\\in(3^{k-1},3^k)$ 有 $\\kappa(m)=k$，"
           "且 $3^k/2>3^{k-1}$（$k\\ge1$），故 $A_k=\\{m:3^k/2<m<3^k,\\ 3\\nmid m\\}$。这些区间（$k=1,2,\\dots$）两两不交，且 $A$ 的元素恰好落在这些区间内。",
           "**计数**：区间 $(3^k/2,3^k)$ 内的整数个数为 $3^k-\\lfloor3^k/2\\rfloor-1=\\frac{3^k-1}{2}$（$3^k$ 为奇数）；其中 3 的倍数个数为 "
           "$\\#\\{j:3^k/2<3j<3^k\\}=\\#\\{j:3^{k-1}/2<j<3^{k-1}\\}=\\frac{3^{k-1}-1}{2}$。两者相减即得 $\\#A_k=3^{k-1}$：",
           ("eq", "$\\#A_k=\\frac{3^k-1}{2}-\\frac{3^{k-1}-1}{2}=3^{k-1}.$"),
           "**求和**：$A(3^K)=\\sum_{k=1}^{K}3^{k-1}=\\frac{3^K-1}{2}$（几何级数）。又 $3^K\\equiv0\\pmod3$ 故 $3^K\\notin A$，于是 $A(3^K-1)=A(3^K)$。"]),
("rem", "rem:vals", "数值核对",
        ["$A(3)=A(2)=1$，$A(9)=A(8)=4$，$A(27)=A(26)=13$，$A(81)=40$，$A(243)=121$，$A(729)=364$，$A(2187)=1093$，"
         "$A(6561)=3280$，$A(59049)=29524$，$A(531441)=265720$，全部等于 $(3^K-1)/2$（附录 A 的 A1 栏）。"]),
("h2", "4.2  与乘 3 的关系（包含且严格）"),
("thm", "thm:times3", "三分结构",
        ["按 $m\\bmod3$ 分解 $A=A_1\\sqcup A_2$（余 1、余 2 两类），$A\\cap3\\mathbb{Z}=\\varnothing$，且",
         ("eq", "$3A+1\\subseteq A_1,\\qquad 3A+2\\subseteq A_2,$"),
         "两个包含均**严格**：$A_1\\setminus(3A+1)\\supseteq\\{19,46,55,64,73,\\dots\\}$，$A_2\\setminus(3A+2)\\supseteq\\{2,5,14,20,41,47,\\dots\\}$。"]),
("proof", ["设 $a\\in A$，取 $k$ 使 $3^k/2<a<3^k$。对 $m=3a+1$：由 $3^{k+1}/2<3a+1<3^{k+1}$（两边乘 3 再加 1 且 $3^{k+1}/2=3\\cdot3^k/2<3a+1$）得 $m\\in A$，且 $m\\equiv1\\pmod3$；对 $3a+2$ 同理。",
           "**严格性**：$19\\in A$（$27/2<19<27$ 且 $19\\equiv1$）而 $19=3\\cdot6+1$，但 $6\\notin A$（$6$ 被 3 整除），故 $19\\notin3A+1$；"
           "$2\\in A$（$3/2<2<3$）而 $2=3\\cdot0+2$，$0\\notin A$；$5=3\\cdot1+2$ 而 $1\\notin A$；$14=3\\cdot4+2$ 而 $4\\notin A$。"]),

("h1", "5  密度与覆盖定理"),
("thm", "thm:density", "$A$ 的密度",
        ["$\\displaystyle\\limsup_{N\\to\\infty}A(N)/N=1/2$，$\\displaystyle\\liminf_{N\\to\\infty}A(N)/N=1/3$；"
         "故 $A$ 的自然密度不存在，比商按 $\\log_3N$ 周期振荡，平均密度为 $5/12$。"]),
("proof", ["由定理 \\T{thm:count}：$A(3^K)=(3^K-1)/2$，故",
           ("eq", "$A(3^K)/3^K=1/2-1/(2\\cdot3^K)\\longrightarrow1/2,$"),
           "给出 $\\limsup\\ge\\frac12$。另一方面，$A$ 在第 $K$ 层内的元素是 $(3^K/2,3^K)$ 中的非 3 倍数，而该层下界处累计值为",
           ("eq", "$A(N)=\\sum_{j=1}^{K}N_3\\!\\left(\\frac{3^{j}}{2},\\ \\min(N,3^{j})\\right)+O(1),\\qquad N_3(a,b)=\\#\\{n\\in(a,b):3\\nmid n\\},$"),
           "取 $N=3^K-1$ 得 $A(3^K-1)=\\frac{3^K-1}{2}$，因此",
           ("eq", "$\\frac{A(3^K-1)}{3^K-1}=\\frac{(3^K-1)/2}{3^K-1}=\\frac12.$"),
           "取 $N=\\lfloor3^K/2\\rfloor$ 时只有前 $K-1$ 层被计入，$A(N)=A(3^{K-1}-1)+O(1)=\\frac{3^{K-1}-1}{2}+O(1)$，于是",
           ("eq", "$A(\\lfloor3^K/2\\rfloor)/\\lfloor3^K/2\\rfloor=1/3+O(3^{-K}),$"),
           "给出 $\\liminf\\le\\frac13$；把上述两条估计与层长度的几何衰减（第 $j$ 层 $(3^j/2,3^j)$ 的长度为 $3^j/2$，故 $j$ 小的层贡献可忽略）合起来即得上下极限分别为 $\\frac12$ 与 $\\frac13$。"
           "平均值由 $\\log_3N$ 的一个周期上的积分给出 $5/12$。"]),
("rem", "rem:A9", "数值核对",
        ["附录 A 的 A9 栏（实际运行输出）：在 $3^7\\le N\\le3^{12}$ 上 $\\min A(N)/N=0.333232$（最小点 $N=3280=\\lfloor3^8/2\\rfloor$），"
         "$\\max A(N)/N=0.500000$（最大点 $N=531440=3^{12}-1$）；在 $N=2\\cdot3^{K-1}$ 处（$K=12$）比商为 $0.416665$，趋于 $5/12$。"
         "若把扫描范围放宽到 $N<3^7$，最小值出现在 $N=1093=\\lfloor3^7/2\\rfloor$，此时 $A(1093)=A(3^6)=364$、比商 $364/1093=0.333028$；"
         "表 A.1 的 A9 栏只记录程序实际扫描区间。三者与定理 \\T{thm:density} 一致。"]),
("thm", "thm:cover", "覆盖定理",
        ["$2\\in M_d$ 对一切 $d\\ge1$；$1\\in M_d$ 对一切 $d\\ge2$（$d=1$ 时 $1\\notin M_1$ 未证，见推论 \\T{cor:notinM1}，但这不影响下面的覆盖等式）。因此对一切 $d\\ge1$，",
         ("eq", "$M_d+3\\mathbb{Z}=\\{m\\in\\mathbb{Z}^+:3\\nmid m\\}.$")]),
("proof", ["**$2\\in M_d$**：取 $x=2$（$n(2)=2$，阈值 $2d$）。轨道 $2\\to1\\to4\\to2\\to1\\to\\cdots$ 的第 $j$ 次除以 2 所得之值当 $j$ 偶时为 $2$、$j$ 奇时为 $1$。因 $2d$ 恒为偶数，故 $p_d(2)=2$。",
           "**$1\\in M_d$（$d$ 为偶数）**：取 $x=1$（$n(1)=1$，阈值 $d$）。轨道 $1\\to4\\to2\\to1\\to\\cdots$ 的第 $j$ 次除以 2 所得之值当 $j$ 奇时为 $2$、$j$ 偶时为 $1$。因 $d$ 为偶数，故 $p_d(1)=1$。",
           "**$1\\in M_d$（$d=3$）**：取 $x=11$（$n(11)=4$，阈值 $12$）。轨道 $11\\to34\\to17\\to52\\to26\\to13\\to40\\to20\\to10\\to5\\to16\\to8\\to4\\to2\\to1$ 恰在第 12 次除以 2 时到达 $1$，故 $p_3(11)=1$。",
           "**$1\\in M_d$（奇数 $d\\ge5$）**：取 $x=7$（$n(7)=3$，阈值 $3d\\ge15$ 为奇数）。轨道 $7\\to22\\to11\\to34\\to17\\to52\\to26\\to13\\to40\\to20\\to10\\to5\\to16\\to8\\to4\\to2\\to1$ "
           "在第 $15$ 次除以 2 时到达 $1$；此后每 2 次除以 2 回到 $1$，故第 $15,17,19,\\dots$ 次（即一切 $\\ge15$ 的奇数次）除以 2 所得之值均为 $1$，特别 $p_d(7)=1$。",
           "**$d=1$**：本文不需要断言 $1\\notin M_1$。已有 $22=p_1(25)\\in M_1$ 且 $22\\equiv1\\pmod3$，故 $1=22-21\\in M_1+3\\mathbb{Z}$。",
           "**覆盖等式**：反向包含（$\\subseteq$）由推论 \\T{cor:subset}；正向包含（$\\supseteq$）由“每个非 3 倍数与 $1$ 或 $2$ 同余模 3”与上述 $1$（对 $d\\ge2$）及 $1\\in M_1+3\\mathbb{Z}$、$2$ 的成员性给出。"]),
("rem", "rem:cover", "", ["定理 \\T{thm:cover} 是手稿猜想唯一需要修改之处：把 $\\mathbb{Z}^+$ 换成 $\\{m:3\\nmid m\\}$。它对 $d=1$ 同样成立。"]),

("h1", "6  猜想与未决问题"),
("conj", "conj:M1A", "$M_1=A$", ["$M_1=A$，即 $\\{p_1(x):x\\in\\mathbb{Z}^+\\}=\\{m:3\\nmid m,\\ \\exists k\\ge1,\\ 3^k/2<m<3^k\\}$。"]),
("p", "**已知**：包含 $M_1\\subseteq A$ 有数值支持（$x\\le2\\times10^5$ 时 $p_1(x)\\in A$ 无反例，附录 A 的 A4 栏），其证明可用辅助量 $S=3^a2^{\\,n(x)-b}-z$："
      "每次奇步 $S\\mapsto3S-1$、每次除以 2 $S\\mapsto S/2$，$S$ 恒为正整数，终态给出 $S=3^{q}-p>0$，即 $p<3^q$；再用“把 $3x+1$ 换成 $3x$”的估计给出下界 $3^q/2<p$，"
      "合起来得 $p\\in A$（**本文未补齐下界那一步**）。反向包含 $A\\subseteq M_1$ 依赖“反向考拉兹树”构造可达性（可任意乘 2；若 $y\\equiv1\\pmod3$ 可反向做 $(y-1)/3$，"
      "再用经典方程 $\\frac{y}{2^b}=\\frac{p}{3^a}-\\frac{C}{3^a2^b}$ 保证首项为整数），**本文未给出独立证明**。因此本版把 $M_1=A$ 列为猜想。"),
("cor", "cor:notinM1", "$1\\notin M_1$，条件性", ["若猜想 \\T{conj:M1A} 成立，则由 $1\\notin A$（注 \\T{rem:Ared}）立得 $1\\notin M_1$。本文未能无条件证明 $1\\notin M_1$。"]),
("item", "**问题 1（$M_1\\subseteq A$ 的严格证明）**：需补下界 $3^q/2<p$ 的辅助量与逐步正性验证（见上）。"),
("item", "**问题 2（$A\\subseteq M_1$）**：需给出反向构造的完整算法，并证明所构造首项恰满足截断阈值。"),
("item", "**问题 3（$M_d$ 的精确刻画，$d\\ge2$）**：是否 $\\bigcup_{d\\ge1}M_d=\\{m:3\\nmid m\\}$？数值观察：在 $p\\le2\\times10^4$ 范围内，"
        "$M_1$ 覆盖非 3 倍数约 $56\\%$、$M_2$ 约 $58\\%$、$M_3$ 约 $38\\%$、$M_6$ 约 $9\\%$；覆盖率随 $d$ 下降源于见证元规模爆炸，不足以判定。"),
("item", "**问题 4（不变集）**：手稿称 $C=M_1$ 满足 $T(C)=C$。本文未处理。"),
("item", "**问题 5（$A(N)$ 的 $O(1)$ 项）**：第 5 节给出一阶估计；$O(1)$ 项的确切界可进一步精确化。"),

("h1", "附录 A  可复现验证程序"),
("p", "程序见 code/verify\\_paper.py（纯标准库）。它按定义 \\T{def:trunc} 与命题 \\T{prop:kappa} 逐项计算。**程序只是核验，不构成证明。**"),
("table", "tab:verify", "核验项目与结果（实际运行输出，对应程序输出 A0–A10）", "lll",
 [("核验项", "范围", "结果"),
  ("A0 位长三定义一致（x.bit\\_length()、$\\lfloor\\log_2x\\rfloor+1$、$\\lceil\\log_2(x+1)\\rceil$）", "$x\\le10^5$", "不一致 0"),
  ("A1 闭式 $A(3^K)=A(3^K-1)=(3^K-1)/2$", "$K\\le12$", "全部 OK"),
  ("A2 分层计数 $\\#A_k=3^{k-1}$", "$k\\le10$", "全部 OK"),
  ("A3 每层 = $(3^k/2,3^k)$ 内非 3 倍数", "$k\\le10$", "全部一致"),
  ("A4 $p_1(x)\\in A$", "$x\\le2\\times10^5$", "反例 0"),
  ("A5 定理 \\T{thm:mod3}：$3\\nmid p_d(x)$", "$x\\le2000,\\ d\\le12$", "反例 0"),
  ("A6 命题 \\T{prop:qge}：$q_d(x)\\ge1$", "$x\\le5000,\\ d\\le12$", "反例 0"),
  ("A7 定理 \\T{thm:cover}：$M_d\\cap\\{1,2,3\\}$", "$d\\le12$", "$d=1:\\{2\\}$；$d\\ge2:\\{1,2\\}$"),
  ("A8 见证元：$p_d(2)=2$（一切 $d$）；$1\\in M_d$：$d=3$ 用 $x=11$，奇 $d\\ge5$ 用 $x=7$", "$d\\le14$", "全部成立"),
  ("A9 比商极值", "$3^7\\le N\\le3^{12}$", "$\\min=0.333232$（$N=3280$）；$\\max=0.500000$（$N=531440$）"),
  ("A10 定理 \\T{thm:times3} 的严格性", "$\\le3000$", "$A_1\\setminus(3A+1)$ 最小元 19；$A_2\\setminus(3A+2)$ 最小元 2")]),
("table", "tab:Apow", "$A(N)$ 在 3 的幂处的值（全部等于 $(3^K-1)/2$）", "rrrr",
 [("$K$", "$3^K$", "$A(3^K)$", "$(3^K-1)/2$"),
  ("1", "3", "1", "1"), ("2", "9", "4", "4"), ("3", "27", "13", "13"),
  ("4", "81", "40", "40"), ("5", "243", "121", "121"), ("6", "729", "364", "364"),
  ("7", "2187", "1093", "1093"), ("8", "6561", "3280", "3280"), ("9", "19683", "9841", "9841"),
  ("10", "59049", "29524", "29524"), ("11", "177147", "88573", "88573"), ("12", "531441", "265720", "265720")]),
("rem", "rem:Averify", "", ["表 \\T{tab:verify} 中“反例数 0”仅表示在所列范围内未发现反例，不构成对相应定理的证明。"]),

("h1", "参考文献"),
("ref", "lagarias", "J. C. Lagarias (ed.), The Ultimate Challenge: The 3x+1 Problem, American Mathematical Society, 2010."),
("ref", "tao", "T. Tao, Almost all Collatz orbits attain almost bounded values, Forum of Mathematics, Pi 10 (2022), e12."),
("ref", "manuscript", "原始手稿《一个猜想》，未刊稿。"),
]


def compute_numbers(content):
    """按 section 顺序为定理类环境与表格编号，供 Word 侧展开 \\T{...}。"""
    num, sec, cnt, tcnt = {}, "", 0, 0
    for e in content:
        kind = e[0]
        if kind == "h1":
            head = e[1].split("  ", 1)[0].strip()
            if head.startswith("附录"):
                sec = head.replace("附录", "").strip()
            elif head.isdigit():
                sec = head
            else:
                sec = ""
            cnt = tcnt = 0
        elif kind in NAMES:
            cnt += 1
            num[e[1]] = "%s.%d" % (sec, cnt)
        elif kind == "table":
            tcnt += 1
            num[e[1]] = "%s.%d" % (sec, tcnt)
    return num


NUM = compute_numbers(CONTENT)


def resolve(s, mode):
    """把 \\T{label} 展开为 LaTeX 的 \\ref 或 Word 侧的编号。"""
    if mode == "tex":
        return re.sub(r"\\T\{([^}]+)\}", lambda m: r"\ref{" + m.group(1) + "}", s)
    return re.sub(r"\\T\{([^}]+)\}", lambda m: NUM.get(m.group(1), "?"), s)


# ------------------------------------------------------------------ LaTeX
def texbold(s):
    out = []
    for i, part in enumerate(s.split("**")):
        if not part:
            continue
        out.append(part if i % 2 == 0 else r"\textbf{" + part + "}")
    return "".join(out)


def tex_flow(segs):
    """段落/公式/项目符号 -> LaTeX 行列表。"""
    out, ul = [], []

    def flush_ul():
        if ul:
            out.append(r"\begin{itemize}" + "".join(r"\item " + x for x in ul) + r"\end{itemize}")
            ul.clear()

    for s in segs:
        if isinstance(s, tuple) and s[0] == "eq":
            flush_ul()
            out.append(r"\[" + s[1].strip("$") + r"\]")
            continue
        for line in resolve(str(s), "tex").split("\n"):
            line = line.strip()
            if not line:
                continue
            if line.startswith("- "):
                ul.append(texbold(line[2:].strip()))
            else:
                flush_ul()
                out.append(texbold(line))
    flush_ul()
    return out


def build_tex(path):
    L = [r"\documentclass[11pt]{article}",
         r"\usepackage[a4paper,margin=2.3cm]{geometry}",
         r"\usepackage{amsmath,amssymb,amsthm}",
         r"\usepackage{ctex}",
         r"\usepackage{booktabs,longtable,array}",
         r"\usepackage[hidelinks]{hyperref}",
         r"\newtheorem{theorem}{定理}[section]",
         r"\newtheorem{lemma}[theorem]{引理}",
         r"\newtheorem{proposition}[theorem]{命题}",
         r"\newtheorem{corollary}[theorem]{推论}",
         r"\theoremstyle{definition}\newtheorem{definition}[theorem]{定义}",
         r"\theoremstyle{definition}\newtheorem{conjecture}[theorem]{猜想}",
         r"\theoremstyle{remark}\newtheorem{remark}[theorem]{注}",
         r"\renewcommand{\proofname}{证明}",
         r"\title{考拉兹截断：更正与部分结果}",
         r"\author{据原始手稿《一个猜想》审读、修订与补证（第五版）}",
         r"\date{\today}",
         r"\begin{document}\maketitle",
         r"\begin{abstract}"]
    absblock, body, items = [], [], []
    bib = False

    def flush():
        if items:
            body.append(r"\begin{itemize}" + "".join(r"\item " + x for x in items) + r"\end{itemize}")
            items.clear()

    for entry in CONTENT:
        kind = entry[0]
        if kind != "item":
            flush()
        if kind == "abs":
            absblock.append(texbold(resolve(entry[1], "tex")))
        elif kind == "kw":
            absblock.append(r"\noindent\textbf{关键词：}" + texbold(resolve(entry[1].split("：", 1)[-1], "tex")))
        elif kind in ("title", "author", "abs_h"):
            continue
        elif kind == "h1":
            parts = entry[1].split("  ", 1)
            head = parts[0]
            title = parts[1] if len(parts) > 1 else parts[0]
            if head.strip().startswith("附录"):
                body.append(r"\appendix")
                body.append(r"\numberwithin{table}{section}")
                body.append(r"\section{" + title + "}")
            elif head.strip() == "参考文献":
                body.append(r"\begin{thebibliography}{9}")
                bib = True
            else:
                body.append(r"\section{" + title + "}")
        elif kind == "h2":
            body.append(r"\subsection{" + entry[1].split("  ", 1)[-1] + "}")
        elif kind == "p":
            body.append(texbold(resolve(entry[1], "tex")))
        elif kind in ("eq", "eqi"):
            body.append(r"\[" + entry[1].strip("$") + r"\]")
        elif kind in NAMES:
            _, label, note, segs = entry
            opt = "[" + note + "]" if note else ""
            env = {"thm": "theorem", "lem": "lemma", "prop": "proposition",
                   "cor": "corollary", "conj": "conjecture", "def": "definition",
                   "rem": "remark"}[kind]
            body.append(r"\begin{" + env + "}" + opt + r"\label{" + label + "}")
            body += tex_flow(segs)
            body.append(r"\end{" + env + "}")
        elif kind == "proof":
            body.append(r"\begin{proof}")
            body += tex_flow(entry[1])
            body.append(r"\end{proof}")
        elif kind == "item":
            items.append(texbold(resolve(entry[1], "tex")))
            continue
        elif kind == "table":
            _, label, caption, colspec, rows = entry
            body.append(r"\begin{table}[htbp]\centering\small")
            body.append(r"\caption{" + resolve(caption, "tex") + r"}")
            body.append(r"\label{" + label + r"}")
            body.append(r"\begin{tabular}{" + colspec + r"}\toprule")
            body.append(" & ".join(resolve(c, "tex") for c in rows[0]) + r" \\ \midrule")
            for row in rows[1:]:
                body.append(" & ".join(resolve(c, "tex") for c in row) + r" \\")
            body.append(r"\bottomrule\end{tabular}")
            body.append(r"\end{table}")
        elif kind == "ref":
            if bib:
                body.append(r"\bibitem{" + entry[1] + r"} " + resolve(entry[2], "tex"))
            else:
                body.append(r"\noindent " + resolve(entry[2], "tex") + r"\par")
    flush()
    if bib:
        body.append(r"\end{thebibliography}")
    L += absblock
    L.append(r"\end{abstract}")
    L += body
    L.append(r"\end{document}")
    open(path, "w", encoding="utf-8").write("\n".join(L) + "\n")

# ------------------------------------------------------------------ Word
GREEK = {"\\mathbb{Z}": "Z", "\\nmid": "∤", "\\mid": "∣", "\\le": "≤", "\\ge": "≥", "\\ne": "≠",
         "\\in": "∈", "\\notin": "∉", "\\subseteq": "⊆", "\\subsetneq": "⊊", "\\varnothing": "∅",
         "\\to": "→", "\\mapsto": "↦", "\\cdot": "·", "\\times": "×", "\\infty": "∞", "\\longrightarrow": "→",
         "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉", "\\kappa": "κ",
         "\\equiv": "≡", "\\exists": "∃", "\\forall": "∀", "\\cap": "∩", "\\cup": "∪", "\\sqcup": "⊔",
         "\\sum": "Σ", "\\pm": "±", "\\dots": "…", "\\cdots": "⋯", "\\qquad": "  ", "\\quad": " ",
         "\\;": " ", "\\:": " ", "\\!": "", "\\%": "%", "\\limsup": "limsup", "\\liminf": "liminf",
         "\\setminus": "∖", "\\supseteq": "⊇", "\\supset": "⊃", "\\frac": "", "\\dfrac": "",
         "\\log": "log", "\\min": "min", "\\max": "max", "\\gcd": "gcd", "\\#": "#", "\\{": "{", "\\}": "}",
         "\\pmod": "(mod ", "\\text": "", "\\mathbf": "", "\\displaystyle": "", "\\,": "", "\\ ": " ",
         "\\not\\equiv": "≢", "\\bmod": " mod ", "\\bigcup": "∪", "\\bigcap": "∩", "\\iff": " ⇔ ",
         "\\left": "", "\\right": "", "\\not": "¬"}
SUP = {"0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
       "k": "ᵏ", "n": "ⁿ", "d": "ᵈ", "s": "ˢ", "t": "ᵗ", "K": "ᴷ", "a": "ᵃ", "b": "ᵇ", "j": "ʲ",
       "m": "ᵐ", "i": "ⁱ", "+": "⁺", "-": "⁻", "−": "⁻"}
SUB = {"0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
       "k": "ₖ", "n": "ₙ", "a": "ₐ", "i": "ᵢ", "j": "ⱼ", "m": "ₘ", "r": "ᵣ", "s": "ₛ", "t": "ₜ", "x": "ₓ"}


def _tex_arg(s, i):
    """读取 LaTeX 参数（{...} 或单个记号），返回 (内容, 新位置)。"""
    while i < len(s) and s[i] == " ":
        i += 1
    if i >= len(s):
        return None
    if s[i] == "{":
        d = 0; j = i
        while j < len(s):
            if s[j] == "{": d += 1
            elif s[j] == "}":
                d -= 1
                if d == 0: break
            j += 1
        return s[i+1:j], j + 1
    if s[i] == "\\":                      # 形如 \frac\alpha\beta
        j = i + 1
        while j < len(s) and s[j].isalpha():
            j += 1
        return s[i:j], j
    return s[i], i + 1                    # 形如 \frac12 = \frac{1}{2}


def _frac(s):
    for cmd in ("\\dfrac", "\\frac"):
        while cmd in s:
            k = s.index(cmd)
            a = _tex_arg(s, k + len(cmd))
            if a is None:
                s = s[:k]; continue
            num, i2 = a
            b = _tex_arg(s, i2)
            if b is None:
                s = s[:k] + num; continue
            den, i3 = b
            s = s[:k] + "(" + num + ")/(" + den + ")" + s[i3:]
    return s


def tex2uni(s):
    s = s.replace("$", "")
    s = _frac(s)
    s = re.sub(r"\\pmod\s*\{([^}]*)\}", r"(mod \1)", s)
    s = re.sub(r"\\pmod\s*([0-9A-Za-z])", r"(mod \1)", s)
    s = s.replace("\\{", "\x01").replace("\\}", "\x02")   # 保护集合花括号
    for a, b in sorted(GREEK.items(), key=lambda x: -len(x[0])):
        s = s.replace(a, b)
    # 上标/下标必须在删除花括号之前处理，否则 ^{K-1} 会被拆成 ᴷ-1（旧版即因此把
    # A(3^{K-1}-1) 印成 A(3ᴷ-1-1)，集合花括号也会整对丢失）
    out = []; i = 0
    while i < len(s):
        if s[i] == "^" and i + 1 < len(s):
            if s[i+1] == "{":
                j = s.find("}", i)
                if j > 0:
                    out.append("^(" + s[i+2:j] + ")"); i = j + 1; continue
            out.append(SUP.get(s[i+1], "^" + s[i+1])); i += 2; continue
        if s[i] == "_" and i + 1 < len(s):
            if s[i+1] == "{":
                j = s.find("}", i)
                if j > 0:
                    out.append("_" + s[i+2:j]); i = j + 1; continue
            out.append(SUB.get(s[i+1], "_" + s[i+1])); i += 2; continue
        out.append(s[i]); i += 1
    s = "".join(out)
    s = s.replace("\\", "").replace("{", "").replace("}", "")
    # 兜底：把残留的 LaTeX 命令名替换为 Unicode（按长度降序）
    SYM = {"supseteq": "⊇", "subseteq": "⊆", "setminus": "∖", "subsetneq": "⊊", "nmid": "∤",
           "in": "∈", "notin": "∉", "le": "≤", "ge": "≥", "ne": "≠", "to": "→", "mapsto": "↦",
           "cap": "∩", "cup": "∪", "cdot": "·", "times": "×", "infty": "∞", "varnothing": "∅",
           "exists": "∃", "forall": "∀", "equiv": "≡", "kappa": "κ", "lfloor": "⌊", "rfloor": "⌋",
           "lceil": "⌈", "rceil": "⌉", "frac": "", "dfrac": "", "displaystyle": "", "sum": "Σ",
           "dots": "…", "cdots": "⋯", "qquad": "  ", "quad": " ", "pmod": "(mod ", "text": "",
           "mathbf": "", "limsup": "limsup", "liminf": "liminf", "log": "log", "min": "min",
           "max": "max", "gcd": "gcd", "pm": "±", "sqcup": "⊔", "mathbb": "", "bmod": " mod ",
           "bigcup": "∪", "bigcap": "∩", "iff": " ⇔ ", "left": "", "right": "", "not": "¬"}
    for k in sorted(SYM, key=len, reverse=True):
        if k in ("in", "to", "le", "ge", "ne", "cap", "cup", "pm", "max", "min", "log", "gcd"):
            continue  # 这些短名易与已转换文本冲突，仅依赖 GREEK 表处理
        s = s.replace(k, SYM[k])
    s = s.replace("∤ ", "∤")
    return s.replace("\x01", "{").replace("\x02", "}")


def eastasia(run, font="宋体"):
    rPr = run._element.get_or_add_rPr()
    rF = rPr.find(qn('w:rFonts'))
    if rF is None:
        rF = OxmlElement('w:rFonts'); rPr.append(rF)
    rF.set(qn('w:eastAsia'), font)


def add_rich(par, text, italic=False, size=None):
    def setf(r):
        if size:
            r.font.size = size
        r.font.name = "Cambria Math"; eastasia(r, "宋体")
    for part in re.split(r"(\*\*.+?\*\*)", text, flags=re.S):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            r = par.add_run(tex2uni(part[2:-2])); r.bold = True; setf(r)
        else:
            r = par.add_run(tex2uni(part)); r.italic = italic; setf(r)


def add_para(doc, text, indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=None):
    p = doc.add_paragraph(); p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)
    add_rich(p, text, size=size)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet"); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_rich(p, text)
    return p


def add_body(doc, segs, size=None):
    """把段落/公式段序列写成 Word 段落。"""
    for s in segs:
        if isinstance(s, tuple) and s[0] == "eq":
            add_para(doc, s[1], indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=size)
            continue
        for line in resolve(str(s), "txt").split("\n"):
            line = line.strip()
            if not line:
                continue
            if line.startswith("- "):
                add_bullet(doc, line[2:].strip())
            else:
                add_para(doc, line, size=size)


def build_docx(path):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Cambria Math"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.2); s.left_margin = s.right_margin = Cm(2.3)
    refs = 0
    for entry in CONTENT:
        kind = entry[0]
        if kind == "title":
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(tex2uni(entry[1])); r.bold = True; r.font.size = Pt(16)
            r.font.name = "Cambria Math"; eastasia(r, "黑体")
        elif kind == "author":
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(entry[1]); r.font.size = Pt(10.5); eastasia(r)
        elif kind == "abs_h":
            p = doc.add_paragraph(); r = p.add_run(entry[1]); r.bold = True; r.font.size = Pt(12); eastasia(r, "黑体")
        elif kind == "abs":
            add_para(doc, resolve(entry[1], "txt"), size=Pt(10))
        elif kind == "kw":
            add_para(doc, resolve(entry[1], "txt"), indent=False, size=Pt(10))
        elif kind in ("h1", "h2"):
            p = doc.add_paragraph(); r = p.add_run(tex2uni(entry[1]))
            r.bold = True; r.font.size = Pt(14 if kind == "h1" else 12)
            r.font.name = "Cambria Math"; eastasia(r, "黑体")
            p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(6)
        elif kind == "p":
            add_para(doc, resolve(entry[1], "txt"))
        elif kind in ("eq", "eqi"):
            add_para(doc, entry[1], indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif kind in NAMES:
            _, label, note, segs = entry
            title = "%s %s" % (NAMES[kind], NUM[label])
            if note:
                title += "（%s）" % tex2uni(note)
            if kind == "rem":
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Pt(21)
                r = p.add_run(title + "  "); r.italic = True
                r.font.name = "Cambria Math"; eastasia(r, "楷体")
                first = resolve(str(segs[0]), "txt")
                add_rich(p, first)
                add_body(doc, segs[1:])
            else:
                first = resolve(str(segs[0]), "txt")
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Pt(21)
                r = p.add_run(title + "  "); r.bold = True; eastasia(r, "黑体")
                add_rich(p, first)
                add_body(doc, segs[1:])
        elif kind == "proof":
            segs = entry[1]
            first = True
            for s in segs:
                if isinstance(s, tuple) and s[0] == "eq":
                    add_para(doc, s[1], indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    continue
                for line in resolve(str(s), "txt").split("\n"):
                    line = line.strip()
                    if not line:
                        continue
                    if line.startswith("- "):
                        add_bullet(doc, line[2:].strip())
                        continue
                    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Pt(21)
                    if first:
                        r = p.add_run("证明  "); r.bold = True; r.italic = True
                        r.font.name = "Cambria Math"; eastasia(r, "楷体")
                        first = False
                    add_rich(p, line)
        elif kind == "item":
            add_bullet(doc, resolve(entry[1], "txt"))
        elif kind == "table":
            _, label, caption, colspec, rows = entry
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_rich(p, "表 %s  %s" % (NUM[label], caption), size=Pt(9.5))
            tb = doc.add_table(rows=len(rows), cols=len(rows[0])); tb.style = "Light Grid Accent 1"
            for i, row in enumerate(rows):
                for j, cell in enumerate(row):
                    cp = tb.cell(i, j).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_rich(cp, resolve(cell, "txt"), size=Pt(9))
                    if i == 0:
                        for r in cp.runs:
                            r.bold = True
        elif kind == "ref":
            refs += 1
            p = doc.add_paragraph(); add_rich(p, "[%d] %s" % (refs, resolve(entry[2], "txt")), size=Pt(10))
            p.paragraph_format.left_indent = Pt(21); p.paragraph_format.first_line_indent = Pt(-21)
    doc.save(path)


if __name__ == "__main__":
    build_docx(os.path.join(OUT, "考拉兹截断_论文.docx"))
    build_tex(os.path.join(OUT, "考拉兹截断_论文.tex"))
    print("OK  编号表：", ", ".join("%s=%s" % (k, v) for k, v in NUM.items()))
