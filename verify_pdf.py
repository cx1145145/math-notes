# -*- coding: utf-8 -*-
"""校验 PDF 的防伪层是否完整、是否被改动。

用法：
    python verify_pdf.py --input "考拉兹截断_完整证明.pdf" [--owner-pw 口令]

逐项检查：/Info 标准与自定义键、XMP 字段、trailer /ID、每页隐形水印载荷、零宽隐写、
内嵌溯源附件、正文指纹复算，最后给出结论。
"""
import argparse
import json
import os
import re
import sys

import fitz

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "code"))
from ac_common import (OWNER_CN, WATERMARK, PAYLOAD_RE, pick_font,  # noqa: E402
                       zw_decode, body_hash, file_hash)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--owner-pw", default=None)
    a = ap.parse_args()
    doc = fitz.open(a.input)
    if doc.needs_pass and a.owner_pw:
        doc.authenticate(a.owner_pw)

    print("=" * 78)
    print("文件：%s" % os.path.abspath(a.input))
    print("页数：%d    加密：%s" % (doc.page_count, "是" if doc.is_encrypted else "否"))
    print("文件 SHA-256：%s" % file_hash(a.input))

    ok = True

    # ---- 1) 标准元数据
    md = doc.metadata or {}
    print("\n[1] /Info 标准字段")
    for k in ("title", "author", "subject", "keywords", "creator", "producer",
              "creationDate", "modDate"):
        print("    %-13s: %s" % (k, md.get(k, "")))
    if OWNER_CN not in (md.get("author") or ""):
        ok = False
        print("    !! Author 字段未包含所有者标识")

    # ---- 2) 自定义 Info 键
    info_xref = None
    try:
        info_xref = int(doc.xref_get_key(-1, "Info")[1].split()[0])
    except Exception:
        pass
    custom = {}
    if info_xref is not None:
        for key in ("DocID", "ContentHash", "IssuedAt", "PageCount", "OwnerHex"):
            try:
                val = doc.xref_get_key(info_xref, key)[1]
                custom[key] = val.strip("()")
            except Exception:
                custom[key] = None
    print("\n[2] /Info 自定义键")
    for k, v in custom.items():
        print("    %-13s: %s" % (k, v))
    if not custom.get("DocID") or not custom.get("ContentHash"):
        ok = False
        print("    !! 缺少 DocID / ContentHash")

    # ---- 3) XMP
    xmp = doc.get_xml_metadata() or ""
    print("\n[3] XMP 元数据（%d 字节）" % len(xmp))
    for tag in ("jlu:Owner", "jlu:DocID", "jlu:ContentSHA256", "jlu:IssuedAt"):
        m = re.search(r"<%s>(.*?)</%s>" % (tag, tag), xmp, re.S)
        print("    %-16s: %s" % (tag, m.group(1).strip() if m else "（缺失）"))
        if not m:
            ok = False

    # ---- 4) trailer /ID
    try:
        idval = doc.xref_get_key(-1, "ID")[1]
    except Exception as exc:                                     # pragma: no cover
        idval = "（读取失败：%s）" % exc
    print("\n[4] trailer /ID：%s" % idval)
    ids = re.findall(r"<([0-9a-fA-F]+)>", idval)

    # ---- 5) 每页隐形水印 + 零宽隐写
    pages_text = [p.get_text() for p in doc]
    all_text = "".join(pages_text)
    hits = PAYLOAD_RE.findall(all_text)
    print("\n[5] 隐形水印：检出 %d 条载荷（%d 页 × 2 处，页眉/页脚各一）" % (len(hits), doc.page_count))
    if hits:
        print("    示例：%s" % hits[0])
    else:
        ok = False
        print("    !! 未检出隐形水印")
    zw_page = zw_decode(all_text)
    doc_id = custom.get("DocID") or ""
    _m = re.search(r"<jlu:StegoZW>(.*?)</jlu:StegoZW>", xmp, re.S)
    zw_meta = zw_decode(md.get("keywords") or "") or zw_decode(_m.group(1) if _m else "")
    print("    零宽隐写（/Keywords 或 XMP 解出）：%s" % (zw_meta or "（无）"))
    print("    页面不可见字形残留的 NUL 标记：%d 个（部分阅读器会以 NUL 呈现，可作存在性签名）"
          % all_text.count("\x00"))
    if zw_meta and doc_id and zw_meta.upper() != doc_id.upper():
        ok = False
        print("    !! 元数据零宽隐写与 /Info 的 DocID 不一致")
    if not zw_meta:
        ok = False
        print("    !! 未检出元数据零宽隐写")

    # ---- 6) 内嵌溯源附件
    names = []
    try:
        names = doc.embfile_names() or []
    except Exception:                                            # pragma: no cover
        pass
    print("\n[6] 内嵌附件：%s" % (names or "（无）"))
    prov = {}
    for n in names:
        try:
            raw = doc.embfile_get(n)
            prov = json.loads(raw.decode("utf-8"))
            print("    %s 内容：" % n)
            for k, v in prov.items():
                if k != "layers":
                    print("        %-20s: %s" % (k, v))
            print("        %-20s: %s" % ("layers", ", ".join(prov.get("layers", []))))
        except Exception as exc:
            print("    %s 解析失败：%s" % (n, exc))
    if not prov:
        ok = False
        print("    !! 缺少溯源附件")

    # ---- 7) 正文指纹复算
    recomputed = body_hash(all_text)
    claimed = (prov.get("content_sha256") or custom.get("ContentHash") or "")
    print("\n[7] 正文指纹（剥离水印后复算）")
    print("    记录值：%s" % claimed)
    print("    复算值：%s" % recomputed)
    if recomputed != claimed:
        ok = False
        print("    !! 不一致：正文可能被改动，或水印结构被破坏")
    else:
        print("    一致 ✔")

    # ---- 8) 文档 ID 一致性交叉核对
    print("\n[8] 文档 ID 交叉核对（应与 %s 完全一致）" % (doc_id or "?"))
    xmp_id = re.search(r"<jlu:DocID>(.*?)</jlu:DocID>", xmp, re.S)
    pay_id = re.search(r"文档\s*ID\s*([0-9A-Fa-f]{32})", hits[0]) if hits else None
    places = {
        "/Info 自定义键": custom.get("DocID") or "",
        "XMP jlu:DocID": xmp_id.group(1).strip() if xmp_id else "",
        "trailer /ID": ids[0] if ids else "",
        "内嵌溯源附件": prov.get("document_id", ""),
        "元数据零宽隐写": zw_meta,
        "页面隐形水印": pay_id.group(1) if pay_id else "",
    }
    for k, v in places.items():
        same = bool(v) and v.upper() == doc_id.upper()
        print("    %-16s: %s   %s" % (k, v or "（无）", "✔" if same else "✗"))
        if not same:
            ok = False

    print("\n" + "=" * 78)
    print("结论：%s" % ("通过 —— 防伪层完整且与正文指纹一致（可认定为 %s 制作的原始文件）"
                      % OWNER_CN if ok else "未通过 —— 请查看上面的 !! 项"))
    print("=" * 78)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
