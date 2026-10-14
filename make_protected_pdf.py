# -*- coding: utf-8 -*-
"""给 PDF 加上多层防伪标记。

已实现的防伪层：
  1. 隐形水印：每页多处 render_mode=3 文本（肉眼不可见、可选中/可检索/可提取）；
  2. 零宽字符隐写：文档 ID 以变体选择符编码，藏在不可见文本里；
  3. /Info 元数据：Author 等标准字段 + 自定义键 DocID / ContentHash / IssuedAt / OwnerHex；
  4. XMP 元数据：dc/xmp/pdf 命名空间 + 自定义防伪命名空间；
  5. 文档唯一 ID：写入 trailer /ID（两串一致）与上述各处；
  6. 正文指纹绑定：剥离水印后的正文 SHA-256，任何正文改动都会导致校验失败；
  7. 内嵌溯源附件 provenance.json（含 ID、时间、哈希、页数、工具版本）；
  8. 可选 OCG 图层（图层名为所有者标识），以及可选 AES-256 权限保护（--encrypt）。

用法：
    python make_protected_pdf.py --input _build/plain.pdf --output 考拉兹截断_完整证明.pdf
    python make_protected_pdf.py --input in.pdf --output out.pdf --encrypt
"""
import argparse
import datetime
import hashlib
import json
import os
import secrets
import sys

import fitz

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "code"))
from ac_common import (OWNER_CN, WATERMARK, PAYLOAD_RE, ZWSET,  # noqa: E402
                       pick_font, zw_encode, body_hash, file_hash)

TITLE = "考拉兹截断：完整证明与开放问题"
NS = "https://nsx.jlu.edu.cn/phys/anti-counterfeit/1.0/"


def xmp_packet(doc_id, text_hash, issued, pages, extra=""):
    return """<?xpacket begin="\ufeff" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
 <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about=""
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:xmp="http://ns.adobe.com/xap/1.0/"
    xmlns:pdf="http://ns.adobe.com/pdf/1.3/"
    xmlns:jlu="%s">
   <dc:title><rdf:Alt><rdf:li xml:lang="x-default">%s</rdf:li></rdf:Alt></dc:title>
   <dc:creator><rdf:Seq><rdf:li>%s</rdf:li></rdf:Seq></dc:creator>
   <dc:description><rdf:Alt><rdf:li xml:lang="x-default">数学预印本；本文件带有不可见防伪水印（%s）</rdf:li></rdf:Alt></dc:description>
   <xmp:CreatorTool>%s 制作</xmp:CreatorTool>
   <pdf:Producer>PyMuPDF %s 防伪流水线</pdf:Producer>
   <jlu:Owner>%s</jlu:Owner>
   <jlu:DocID>%s</jlu:DocID>
   <jlu:ContentSHA256>%s</jlu:ContentSHA256>
   <jlu:IssuedAt>%s</jlu:IssuedAt>
   <jlu:PageCount>%d</jlu:PageCount>%s
  </rdf:Description>
 </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>""" % (NS, TITLE, OWNER_CN, WATERMARK, OWNER_CN, fitz.VersionBind,
                         OWNER_CN, doc_id, text_hash, issued, pages, extra)


def build(src_path, out_path, encrypt=False, owner_pw=None):
    src = fitz.open(src_path)
    pages = src.page_count
    doc_id = secrets.token_hex(16).upper()
    issued = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    text_hash = body_hash("".join(p.get_text() for p in src))
    font = pick_font()
    if not font:
        print("警告：找不到中文字体，隐形水印将用内置字体（可能无法正确提取中文）")

    # 8) 可选的防伪图层（OCG）
    ocg = 0
    try:
        ocg = src.add_ocg("防伪层 " + OWNER_CN, on=True)
    except Exception as exc:                                    # pragma: no cover
        print("  （OCG 图层不可用：%s）" % exc)

    # 1) + 2) 每页的隐形水印与零宽隐写
    #    只放在页边空白（上、下各一处），避免与正文文字交织：提取时每条载荷都是完整连续串，
    #    校验端才能可靠地"剥离水印后复算正文指纹"。
    for i, page in enumerate(src, 1):
        if font:
            page.insert_font(fontname="wmk", fontfile=font)
        payload = "%s｜文档ID %s｜第 %d/%d 页｜发行 %s｜正文指纹 %s" % (
            WATERMARK, doc_id, i, pages, issued, text_hash[:16])
        h = page.rect.height
        for y in (24, h - 18):
            try:
                page.insert_text((36, y), payload, fontname="wmk", fontsize=5.5,
                                 render_mode=3, oc=ocg)
            except Exception:
                page.insert_text((36, y), payload, fontname="wmk", fontsize=5.5, render_mode=3)

    # 3) 标准元数据（Keywords 里夹带一份零宽隐写的文档 ID）
    now = fitz.get_pdf_now()
    src.set_metadata({
        "title": TITLE,
        "author": OWNER_CN,
        "subject": "数学预印本（考拉兹截断）；含不可见防伪水印",
        "keywords": "考拉兹猜想;考拉兹截断;防伪;%s;%s" % (doc_id, zw_encode(doc_id)),
        "creator": OWNER_CN + " 制作",
        "producer": "PyMuPDF %s 防伪流水线" % fitz.VersionBind,
        "creationDate": now,
        "modDate": now,
    })

    # 3b) /Info 自定义键（ASCII，便于程序读取）
    try:
        info = src.xref_get_key(-1, "Info")
        info_xref = int(info[1].split()[0])
        for key, val in (("DocID", doc_id), ("ContentHash", text_hash),
                         ("IssuedAt", issued), ("PageCount", str(pages)),
                         ("OwnerHex", OWNER_CN.encode("utf-16-be").hex().upper())):
            src.xref_set_key(info_xref, key, "(%s)" % val)
    except Exception as exc:                                    # pragma: no cover
        print("  （自定义 Info 键写入失败：%s）" % exc)

    # 4) XMP（另带一份零宽隐写载荷）
    try:
        extra = "\n   <jlu:StegoZW>%s</jlu:StegoZW>" % zw_encode(doc_id)
        src.set_xml_metadata(xmp_packet(doc_id, text_hash, issued, pages, extra))
    except Exception as exc:                                    # pragma: no cover
        print("  （XMP 写入失败：%s）" % exc)

    # 5) trailer /ID
    try:
        src.xref_set_key(-1, "ID", "[<%s> <%s>]" % (doc_id.lower(), doc_id.lower()))
    except Exception as exc:                                    # pragma: no cover
        print("  （/ID 写入失败：%s）" % exc)

    # 7) 内嵌溯源附件
    prov = {
        "owner": OWNER_CN,
        "watermark": WATERMARK,
        "document_id": doc_id,
        "title": TITLE,
        "issued_at": issued,
        "page_count": pages,
        "content_sha256": text_hash,
        "source_pdf_sha256": file_hash(src_path),
        "generator": "make_protected_pdf.py (PyMuPDF %s)" % fitz.VersionBind,
        "layers": ["invisible-text-watermark", "zero-width-steganography",
                   "info-metadata", "xmp-metadata", "trailer-id",
                   "content-hash-binding", "embedded-provenance", "ocg-layer"],
        "note": "content_sha256 是对剥离全部隐形水印后的正文规范化文本计算的 SHA-256；"
                "verify_pdf.py 可复算并核对。",
    }
    try:
        src.embfile_add("provenance.json",
                        json.dumps(prov, ensure_ascii=False, indent=2).encode("utf-8"),
                        filename="provenance.json",
                        desc="防伪溯源信息 anti-counterfeit provenance")
    except Exception as exc:                                    # pragma: no cover
        print("  （内嵌附件失败：%s）" % exc)

    kw = {}
    if encrypt:
        kw = dict(encryption=fitz.PDF_ENCRYPT_AES_256,
                  owner_pw=owner_pw or secrets.token_urlsafe(12), user_pw="",
                  permissions=int(fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT))
    try:                                    # 只保留用到的字形，避免整份中文字体进文件
        src.subset_fonts()
    except Exception as exc:                                    # pragma: no cover
        print("  （字体子集化跳过：%s）" % exc)
    src.save(out_path, garbage=4, deflate=True, clean=True, **kw)
    src.close()

    # 自洽性检查：重新打开，剥离水印后复算正文指纹
    chk = fitz.open(out_path)
    if kw:
        chk.authenticate(kw["owner_pw"])
    got = body_hash("".join(p.get_text() for p in chk))
    ok = (got == text_hash)
    chk.close()
    return dict(doc_id=doc_id, issued=issued, text_hash=text_hash, recomputed=got,
                pages=pages, self_consistent=ok, encrypt=bool(kw),
                owner_pw=kw.get("owner_pw", ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--encrypt", action="store_true", help="额外加 AES-256 权限保护（禁止修改）")
    a = ap.parse_args()
    rep = build(a.input, a.output, encrypt=a.encrypt)
    print("=" * 74)
    print("已生成防伪 PDF：", a.output)
    print("  所有者标识   : %s" % OWNER_CN)
    print("  文档唯一 ID  : %s" % rep["doc_id"])
    print("  签发时间     : %s" % rep["issued"])
    print("  页数         : %d" % rep["pages"])
    print("  正文指纹     : %s" % rep["text_hash"])
    print("  指纹自洽     : %s" % ("是" if rep["self_consistent"] else "否"))
    if rep["encrypt"]:
        print("  权限保护     : AES-256（所有者口令 %s）" % rep["owner_pw"])
    print("  文件 SHA-256 : %s" % file_hash(a.output))
    print("=" * 74)
    print("校验命令： python verify_pdf.py --input \"%s\"" % a.output)


if __name__ == "__main__":
    main()
