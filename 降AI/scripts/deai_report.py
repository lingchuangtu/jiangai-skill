#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deai_report.py —— 一条命令出完整降AI报告：改前分数 + 改后分数 + Word 对照稿（改动标绿）。

改后全文从标准输入读入（--revised -），**全程不产生任何中间文件**，
所以交付时只会有一个文件：Word 对照稿。

用法：
  python3 deai_report.py --original 初稿.docx --out 初稿-降AI对照稿.docx --revised - <<'EOF'
  （改后全文，保留段落换行）
  EOF

可选：
  --highlight     对照稿用绿色荧光笔（默认绿色字体）
  --show-deleted  被删掉的原文用灰色删除线显示
  --top N         列出最该改的 N 句（默认 5）

零依赖：仅标准库。改前/改后分数直接调用同目录的 deai_score.py。
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import make_diff_docx as mdd  # noqa: E402

SCORE_PY = HERE / "deai_score.py"


def score(path=None, text=None):
    """调 deai_score.py 拿 JSON 结果。path 为文件路径，text 为标准输入内容。"""
    cmd = [sys.executable, str(SCORE_PY), path if path else "-", "--json"]
    r = subprocess.run(cmd, input=text or "", capture_output=True, text=True)
    if r.returncode != 0:
        msg = (r.stdout or "").strip() or (r.stderr or "").strip()
        raise SystemExit(f"打分失败：{msg}")
    try:
        return json.loads(r.stdout)
    except json.JSONDecodeError:
        raise SystemExit(f"打分输出无法解析：{r.stdout[:300]}")


def main():
    ap = argparse.ArgumentParser(description="一条命令出降AI报告（分数 + 标绿对照稿）")
    ap.add_argument("--original", required=True, help="原稿（.docx/.txt/.md）")
    ap.add_argument("--revised", required=True, help="传 - 从标准输入读改后全文（推荐，零中间文件）")
    ap.add_argument("--out", required=True, help="输出对照稿 docx 路径")
    ap.add_argument("--highlight", action="store_true", help="绿色荧光笔高亮（默认绿色字体）")
    ap.add_argument("--show-deleted", action="store_true", help="灰色删除线显示被删内容")
    ap.add_argument("--top", type=int, default=5, help="列出最该改的 N 句")
    args = ap.parse_args()

    orig = mdd.read_text(args.original)
    if not orig.strip():
        raise SystemExit("原稿是空的，请检查文件")

    if args.revised == "-":
        if sys.stdin.isatty():
            raise SystemExit("没收到改后全文：请用 heredoc 或管道把改后文喂进来（--revised - <<'EOF' ... EOF）")
        rev = sys.stdin.read()
    else:
        rev = mdd.read_text(args.revised)
    if not rev.strip():
        raise SystemExit("改后全文是空的")

    # 改前 / 改后分数
    before = score(path=args.original)
    after = score(text=rev)

    # 生成对照稿
    mdd.HIGHLIGHT = bool(args.highlight)
    runs, n_orig, n_rev = mdd.build_runs(orig, rev, args.show_deleted)
    if not runs:
        raise SystemExit("改后稿没解析出任何句子")
    max_pi = max(pi for pi, _, _ in runs)
    buckets = [[] for _ in range(max_pi + 1)]
    for pi, t, st in runs:
        buckets[pi].append((t, st))
    n_green = sum(1 for _, _, st in runs if st == "green")
    n_del = sum(1 for _, _, st in runs if st == "del")
    pct = round(n_green * 100 / max(1, n_rev))
    title = [
        "降AI 对照稿（绿色 = 相对原稿改过/新增的句子）",
        f"原稿 {n_orig} 句 → 改后 {n_rev} 句｜改动 {n_green} 句（{pct}%）"
        + (f"｜删除 {n_del} 句（灰色删除线）" if args.show_deleted else "")
        + f"｜AI 特征分 {before['score']} → {after['score']}",
    ]
    mdd.write_docx(args.out, buckets, title)

    # 终端摘要（给用户看的部分由模型复述，别贴全文）
    print(f"\n=== 降AI 报告 ===")
    print(f"分数：{before['score']} → {after['score']}（{after['verdict']}）")
    print(f"句子：原 {n_orig} → 改 {n_rev}，标绿 {n_green} 句（{pct}%）")
    print(f"对照稿：{args.out}")
    if n_green == 0:
        print("⚠️ 没有任何标绿句子——请确认改后稿确实与原稿不同")
    worst = after.get("worst_sentences") or []
    if worst:
        print("\n改后仍需留意的句子（分数已达标可忽略）：")
        for m in worst[: args.top]:
            print(f"  第{m['para']}段第{m['sent']}句（{m['len']}字）：{m['text'][:50]}")


if __name__ == "__main__":
    main()
