#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成「降AI 对照稿」docx：在改后稿的基础上，改过/新增的句子标绿色。

用法：
  python3 make_diff_docx.py --original 原稿.docx --revised 改后.md --out 对照稿.docx
  python3 make_diff_docx.py --original 原稿.txt  --revised 改后.txt --out 对照稿.docx --highlight
  python3 make_diff_docx.py --original 原稿.docx --revised 改后.md --out 对照稿.docx --show-deleted

参数：
  --original       原始稿（.docx / .txt / .md，docx 零依赖直读）
  --revised        改后稿（.docx / .txt / .md）
  --out            输出 docx 路径
  --highlight      用绿色荧光笔高亮（默认用绿色字体）
  --show-deleted   被删掉的原文用灰色删除线显示（默认不显示，输出可直接用）

零依赖：仅用标准库（zipfile / xml.etree / difflib），豆包沙盒无需 pip 安装。
"""
import argparse
import difflib
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
GREEN = "00B050"      # 绿色字体
GRAY = "808080"       # 删除线灰


# ---------- 读取 ----------
def read_text(path):
    """读 docx / txt / md，返回纯文本（docx 按段落拼回，保留换行结构）。"""
    low = path.lower()
    if low.endswith(".docx"):
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml")
        root = ET.fromstring(xml)
        paras = []
        for pnode in root.iter(f"{{{W}}}p"):
            texts = [t.text or "" for t in pnode.iter(f"{{{W}}}t")]
            paras.append("".join(texts))
        # 去掉首尾空段，中间空段保留为分隔
        while paras and not paras[0].strip():
            paras.pop(0)
        while paras and not paras[-1].strip():
            paras.pop()
        return "\n".join(paras)
    with open(path, encoding="utf-8") as f:
        return f.read()


_SENT_RE = re.compile(r"[^。！？；;!?]*[。！？；;!?]+|[^。！？；;!?]+$")


def sentences(text):
    """切成 [(文本, 段落序号)]：段落内按句切，句末标点跟随句子。"""
    out = []
    for pi, para in enumerate(text.replace("\r\n", "\n").split("\n")):
        para = para.strip()
        if not para:
            continue
        for m in _SENT_RE.finditer(para):
            s = m.group(0).strip()
            if s:
                out.append((s, pi))
    return out


# ---------- 差异对齐 ----------
def build_runs(orig, rev, show_deleted):
    """返回 [(段落号, 文本, 样式)]，样式 in {plain, green, del}。"""
    os_ = sentences(orig)
    rs = sentences(rev)
    sm = difflib.SequenceMatcher(None, [t for t, _ in os_], [t for t, _ in rs], autojunk=False)
    runs = []

    def add(j_idx, style):
        t, pi = rs[j_idx]
        runs.append((pi, t, style))

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for j in range(j1, j2):
                add(j, "plain")
        elif tag in ("replace", "insert"):
            for j in range(j1, j2):
                add(j, "green")
        elif tag == "delete" and show_deleted:
            # 删除内容挂在相邻改后句所在段落之后展示
            anchor_pi = rs[j1][1] if j1 < len(rs) else (rs[-1][1] if rs else 0)
            for i in range(i1, i2):
                runs.append((anchor_pi, os_[i][0], "del"))
    return runs, len(os_), len(rs)


# ---------- docx 生成 ----------
def _esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _rpr(style, highlight):
    if style == "green":
        color = f'<w:color w:val="{GREEN}"/>'
        hl = '<w:highlight w:val="green"/>' if highlight else ""
        return f"<w:rPr>{color}{hl}</w:rPr>"
    if style == "del":
        return f'<w:rPr><w:strike/><w:color w:val="{GRAY}"/></w:rPr>'
    return ""


def write_docx(out_path, body_paras, title_lines):
    """body_paras: [[(文本, 样式), ...], ...]；title_lines: [str]"""
    parts = []

    def para_xml(runs, pstyle=""):
        pr = f"<w:pPr>{pstyle}</w:pPr>" if pstyle else ""
        rs = "".join(
            f'<w:r>{_rpr(st, HIGHLIGHT)}<w:t xml:space="preserve">{_esc(t)}</w:t></w:r>'
            for t, st in runs
        )
        return f"<w:p>{pr}{rs}</w:p>"

    for line in title_lines:
        parts.append(para_xml([(line, "plain")], '<w:jc w:val="center"/>'))
    # 标题与正文之间空一行
    parts.append("<w:p/>")
    for runs in body_paras:
        parts.append(para_xml(runs) if runs else "<w:p/>")

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}"><w:body>{"".join(parts)}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
        "</w:body></w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)


# ---------- 主流程 ----------
def main():
    global HIGHLIGHT
    ap = argparse.ArgumentParser(description="生成降AI对照稿：改过的句子标绿色")
    ap.add_argument("--original", required=True, help="原稿（.docx/.txt/.md）")
    ap.add_argument("--revised", required=True, help="改后稿（.docx/.txt/.md）；传 - 表示从标准输入读取（推荐：不产生中间文件）")
    ap.add_argument("--out", required=True, help="输出 docx 路径")
    ap.add_argument("--highlight", action="store_true", help="绿色荧光笔高亮（默认绿色字体）")
    ap.add_argument("--show-deleted", action="store_true", help="灰色删除线显示被删内容")
    args = ap.parse_args()
    HIGHLIGHT = args.highlight

    orig = read_text(args.original)
    if args.revised == "-":
        rev = sys.stdin.read()
    else:
        rev = read_text(args.revised)
    if not orig.strip():
        sys.exit("原稿是空的")
    if not rev.strip():
        sys.exit("改后稿是空的")

    runs, n_orig, n_rev = build_runs(orig, rev, args.show_deleted)
    if not runs:
        sys.exit("改后稿没解析出任何句子")

    # 按段落聚合（保持改后稿的段落结构）
    max_pi = max(pi for pi, _, _ in runs)
    buckets = [[] for _ in range(max_pi + 1)]
    for pi, t, st in runs:
        buckets[pi].append((t, st))

    n_green = sum(1 for _, _, st in runs if st == "green")
    n_del = sum(1 for _, _, st in runs if st == "del")
    pct = round(n_green * 100 / max(1, n_rev))
    title_lines = [
        "降AI 对照稿（绿色 = 相对原稿改过/新增的句子）",
        f"原稿 {n_orig} 句 → 改后 {n_rev} 句｜改动 {n_green} 句（{pct}%）"
        + (f"｜删除 {n_del} 句（灰色删除线）" if args.show_deleted else ""),
    ]
    write_docx(args.out, buckets, title_lines)
    print(f"✅ 已生成 {args.out}")
    print(f"   原稿 {n_orig} 句 → 改后 {n_rev} 句，标绿 {n_green} 句（{pct}%）")
    if n_green == 0:
        print("   ⚠️ 没有任何标绿句子——请检查改后稿是否真的与原稿不同")


if __name__ == "__main__":
    main()
