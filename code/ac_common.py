# -*- coding: utf-8 -*-
"""防伪层公用常量与工具（签发端 make_protected_pdf.py 与校验端 verify_pdf.py 共用）。"""
import hashlib
import os
import re

OWNER_CN = "吉林大学物理学院11260724"
WATERMARK = "本文件由吉林大学物理学院11260724制作"

# 变体选择符 U+FE00..U+FE0F：肉眼与所见即所得界面中不可见，但可程序还原
ZW = [chr(0xFE00 + i) for i in range(16)]
ZWSET = set(ZW)
NULCH = "\ufe00"                     # 页面级不可见通道用的字符
SEP = "\u00a4"                       # 通道分隔符（¤，正文中不会出现）

# 隐形水印载荷（正则用于校验时把水印从正文中剥离；放宽空白容忍度以适配不同提取器）
PAYLOAD_RE = re.compile(
    re.escape(WATERMARK) +
    r"｜文档\s*ID\s*[0-9A-F]{32}｜第\s*\d+\s*/\s*\d+\s*页"
    r"｜发行\s*[0-9A-Za-z:+.\-]+｜正文指纹\s*[0-9a-f]{16}")

FONTS = [r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\simkai.ttf",
         r"C:\Windows\Fonts\simfang.ttf", r"C:\Windows\Fonts\msyh.ttc",
         r"C:\Windows\Fonts\simsun.ttc"]


def pick_font():
    for f in FONTS:
        if os.path.exists(f):
            return f
    return None


def zw_encode(hexstr):
    """把十六进制串编码为变体选择符序列（元数据里可原样保存，肉眼不可见）。"""
    return "".join(ZW[int(c, 16)] for c in hexstr if c in "0123456789abcdefABCDEF")


def zw_decode(text):
    nib = {c: "%x" % i for i, c in enumerate(ZW)}
    return "".join(nib[c] for c in text if c in nib)


def channel_encode(hexstr):
    """页面级不可见数据通道：每个十六进制位用「分隔符 + 1..16 个不可见字符」表示。

    实测：中文字体没有这些字形，提取时每个不可见字符变成 1 个 NUL，
    分隔符（¤，正文中不出现）用于可靠地切分各个 NUL 串。"""
    return "".join(SEP + NULCH * (int(c, 16) + 1) for c in hexstr
                   if c in "0123456789abcdefABCDEF")


def channel_decode(text):
    """还原通道内容：分隔符后 NUL 串长度减 1 即一个十六进制位。"""
    return "".join("%x" % (len(m.group(1)) - 1)
                   for m in re.finditer(re.escape(SEP) + r"(\x00{1,16})", text))


def strip_marks(text):
    """去掉隐形水印载荷、零宽隐写字符与不可见数据通道，还原"纯正文"。"""
    text = "".join(ch for ch in text if ch not in ZWSET)
    text = re.sub(r"[\x00" + re.escape(SEP) + r"]", "", text)
    return PAYLOAD_RE.sub("", text)


def norm(text):
    return re.sub(r"\s+", "", text)


def body_hash(pages_text):
    """正文指纹：剥离全部防伪标记后的规范化文本的 SHA-256。"""
    return hashlib.sha256(norm(strip_marks(pages_text)).encode("utf-8")).hexdigest()


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()
