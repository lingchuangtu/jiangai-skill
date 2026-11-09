#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
split_units.py —— 把论文切成「第 N 段 / 第 N.M 句」的编号单元，方便逐段改写

用法：
    python3 split_units.py 初稿.docx                    # 打印到屏幕
    python3 split_units.py 初稿.docx --out units.md     # 写文件
    python3 split_units.py 初稿.txt --long 45           # 只列出 ≥45 字的长句
"""
import sys, re, zipfile, argparse
from pathlib import Path

def read_text(path: str) -> str:
    p = Path(path)
    if p.suffix.lower() == ".docx":
        with zipfile.ZipFile(p) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
        xml = re.sub(r"</w:p>", "\n", xml)
        return re.sub(r"<[^>]+>", "", xml)
    return p.read_text(encoding="utf-8", errors="ignore")

def strip_punct(s: str) -> str:
    return re.sub(r"[，。；：？！——（）“”…、《》\s\"'‘’]", "", s)

def split_paras(text: str):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    if len(paras) <= 1:
        paras = [p.strip() for p in text.split("\n") if p.strip()]
    return paras

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--out", default="")
    ap.add_argument("--long", type=int, default=0, help="只列出 >=N 字的句子")
    args = ap.parse_args()

    paras = split_paras(read_text(args.file))
    lines, long_sents = [], []

    for pi, p in enumerate(paras, 1):
        n = len(strip_punct(p))
        lines.append(f"\n## 第 {pi} 段（{n} 字）\n")
        lines.append(f"> 原段：{p}\n")
        sents = [s.strip() for s in re.split(r"[。！？；!?;]", p) if strip_punct(s)]
        for si, s in enumerate(sents, 1):
            ln = len(strip_punct(s))
            flag = "  ⚠️超长" if ln >= 45 else ("  ▫️短" if ln <= 10 else "")
            lines.append(f"- {pi}.{si}（{ln} 字）{s}{flag}")
            if ln >= 45:
                long_sents.append((pi, si, ln, s))

    if args.long:
        print(f"\n=== {args.long} 字及以上的句子：{len(long_sents)} 句 ===")
        for pi, si, ln, s in long_sents:
            print(f"  {pi}.{si}（{ln} 字）：{s[:70]}")
        return

    body = "\n".join(lines)
    if args.out:
        Path(args.out).write_text(body, encoding="utf-8")
        print(f"已写出：{args.out}")
        print(f"段落 {len(paras)} / 超长句 {len(long_sents)} 句")
    else:
        print(body)
        print(f"\n段落 {len(paras)} / 超长句（≥45字） {len(long_sents)} 句 —— 这些优先拆")

if __name__ == "__main__":
    main()
