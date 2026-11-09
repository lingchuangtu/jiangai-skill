#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
deai_score.py —— 论文文本「AI 腔」体检与打分（纯标准库，无依赖）

用法：
    python3 deai_score.py 论文.txt
    python3 deai_score.py 初稿.docx
    python3 deai_score.py - < 改后.txt        # 从标准输入读，不产生中间文件
    python3 deai_score.py 论文.txt --json          # 机读输出
    python3 deai_score.py 论文.txt --top 8         # 列出最像 AI 的 8 句

设计依据：AI 检测器主要吃两类信号
  1) 困惑度（perplexity）：用词是否「太顺、太常见、太可预测」
  2) 突发性（burstiness）：句长与句式是否有人的起伏
此外，中文检测还大量依赖模板化结构（首先/其次/综上所述）、机械并列、名词化堆叠。
本脚本把上述特征量化成 0-100 分，分数越高越像人写。
"""
import sys, re, json, zipfile, statistics, argparse
from pathlib import Path

# ---------------- 读取 ----------------
def read_text(path: str) -> str:
    if path == "-":
        return sys.stdin.read()
    p = Path(path)
    if p.suffix.lower() == ".docx":
        with zipfile.ZipFile(p) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
        xml = re.sub(r"</w:p>", "\n", xml)
        txt = re.sub(r"<[^>]+>", "", xml)
        return txt
    return p.read_text(encoding="utf-8", errors="ignore")

# ---------------- 词表 ----------------
BAD_CONNECT = [
    "首先", "其次", "再次", "最后", "此外", "与此同时", "综上所述", "总而言之",
    "值得注意的是", "不难发现", "显而易见", "由此可见", "换言之", "总而言之",
    "一方面", "另一方面", "不仅", "而且", "既", "又",
    "总的来说", "总体而言", "在一定程度上", "从某种意义上讲", "换言之",
    "需要注意的是", "需要指出的是", "不可否认", "毋庸置疑", "不言而喻",
]
BAD_HOLLOW = [
    "深刻影响", "重要意义", "现实意义", "理论意义", "积极作用", "重要作用",
    "全面提升", "有效促进", "显著改善", "日益凸显", "愈发重要", "至关重要",
    "具有重要的现实意义", "提供了新的思路", "奠定坚实基础", "具有广阔前景",
    "在当今社会", "随着科技的不断发展", "随着社会的不断进步", "在信息化时代",
    "在数字化背景下", "在这个快速发展的时代",
]
BAD_VERBOSE = [
    "进行了分析", "进行了研究", "进行了探讨", "进行了阐述", "进行了论述",
    "具有一定的", "起到了重要的作用", "起到了积极的作用",
    "对……进行了", "作出了重要贡献",
]
PATTERNS = [
    ("不仅…而且", r"不仅[^。；！？]{0,20}(而且|还|也)"),
    ("一方面…另一方面", r"一方面[^。；！？]{0,30}另一方面"),
    ("首先…其次", r"首先[^。！？]{0,120}其次"),
    ("既…又", r"既[^。；！？]{0,12}又"),
    ("为…提供了", r"为[^。；！？]{0,20}提供了"),
    ("随着…的发展", r"随着[^。；！？]{0,25}(发展|推进|深入)"),
    ("在…背景下", r"在[^。；！？]{0,20}背景下"),
    ("三重排比", r"([^。；！？]{2,10})[，、]\1?[^。；！？]{0,6}[，、][^。；！？]{0,6}[，、]"),
]
PUNCT_SET = "，。；：？！—（）“”…、《》,."

def strip_punct(s: str) -> str:
    return re.sub(r"[，。；：？！——（）“”…、《》\s\"'‘’]", "", s)

def split_paras(text: str):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    if len(paras) <= 1:
        paras = [p.strip() for p in text.split("\n") if p.strip()]
    return paras

def split_sents(para: str):
    parts = re.split(r"[。！？；!?;]", para)
    return [p.strip() for p in parts if strip_punct(p)]

def linear(x, lo, hi, reverse=False):
    """把 x 从 [lo,hi] 映射到 0..1；reverse=True 表示越小越好。"""
    if hi == lo:
        return 0.0
    v = (x - lo) / (hi - lo)
    v = max(0.0, min(1.0, v))
    return 1 - v if reverse else v

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--top", type=int, default=6)
    args = ap.parse_args()

    text = read_text(args.file)
    paras = split_paras(text)
    if not paras:
        print("没读到文字，检查文件是否为空或为扫描版 PDF（请先转文字）")
        sys.exit(1)

    plain = "".join(paras)
    total_chars = len(strip_punct(plain)) or 1

    sents, sent_meta = [], []
    for pi, p in enumerate(paras):
        for si, s in enumerate(split_sents(p)):
            n = len(strip_punct(s))
            if n == 0:
                continue
            sents.append(n)
            sent_meta.append({"para": pi + 1, "sent": si + 1, "len": n, "text": s})

    if len(sents) < 5:
        print("文本太短（不足 5 句），打分不可靠。至少给一段完整正文。")
        sys.exit(1)

    mean_len = statistics.mean(sents)
    sd_len = statistics.pstdev(sents) if len(sents) > 1 else 0.0
    short_ratio = sum(1 for n in sents if n <= 10) / len(sents)
    long_ratio = sum(1 for n in sents if n >= 45) / len(sents)

    hits = {"连接/总结词": 0, "空话套话": 0, "冗余动词": 0}
    hit_detail = []
    for w in BAD_CONNECT:
        c = plain.count(w)
        if c:
            hits["连接/总结词"] += c
            hit_detail.append((w, c, "连接词"))
    for w in BAD_HOLLOW:
        c = plain.count(w)
        if c:
            hits["空话套话"] += c
            hit_detail.append((w, c, "套话"))
    for w in BAD_VERBOSE:
        if "……" in w:
            continue
        c = plain.count(w)
        if c:
            hits["冗余动词"] += c
            hit_detail.append((w, c, "冗余"))

    pat_hits = []
    for name, rx in PATTERNS:
        c = len(re.findall(rx, plain))
        if c:
            pat_hits.append((name, c))

    bad_total = sum(hits.values())
    bad_per_k = bad_total / total_chars * 1000
    pat_total = sum(c for _, c in pat_hits)

    de_density = plain.count("的") / total_chars * 100
    para_lens = [len(strip_punct(p)) for p in paras]
    para_cv = (statistics.pstdev(para_lens) / statistics.mean(para_lens)) if len(para_lens) > 1 and statistics.mean(para_lens) else 0.0
    punct_kinds = len({ch for ch in plain if ch in PUNCT_SET})

    sub = {
        "句长波动": (linear(sd_len, 4.0, 11.0), f"标准差 {sd_len:.1f} 字（目标≥8.5）"),
        "长短交替": (linear(short_ratio, 0.03, 0.22) * 0.6 + linear(long_ratio, 0.02, 0.18) * 0.4,
                     f"短句≤10字占 {short_ratio*100:.0f}%（目标15–30%），长句≥45字占 {long_ratio*100:.0f}%（目标≥8%）"),
        "套话密度": (linear(bad_per_k, 1.0, 9.0, reverse=True), f"每千字 {bad_per_k:.1f} 处（目标≤2）"),
        "机械并列": (linear(pat_total, 0, 5, reverse=True), f"并列/模板句 {pat_total} 处（目标≤1）"),
        "段落差异": (linear(para_cv, 0.10, 0.45), f"段长变异系数 {para_cv:.2f}（目标≥0.35）"),
        "的字段位": (linear(de_density, 3.0, 6.0, reverse=True), f"每百字「的」{de_density:.1f} 个（目标≤4.0）"),
        "标点多样": (linear(punct_kinds, 3, 6), f"使用 {punct_kinds} 种标点（目标≥5）"),
    }
    weights = {"句长波动": .22, "长短交替": .16, "套话密度": .26, "机械并列": .12,
               "段落差异": .10, "的字段位": .06, "标点多样": .08}
    score = round(sum(sub[k][0] * weights[k] for k in sub) * 100, 1)

    # 找出最像 AI 的句子：长度接近均值 + 含套话 + 以连接词开头
    def ai_likeness(m):
        s = m["text"]
        pen = abs(m["len"] - mean_len) / max(mean_len, 1)
        for w in BAD_CONNECT + BAD_HOLLOW:
            if w in s:
                pen += 0.45
        if re.match(r"^(首先|其次|此外|综上所述|随着|在当今|一方面)", s):
            pen += 0.5
        return pen
    worst = sorted(sent_meta, key=ai_likeness, reverse=True)[: args.top]

    verdict = ("很像人写" if score >= 90 else "基本可用" if score >= 80 else "需要二改" if score >= 70 else "AI 痕迹明显")

    if args.json:
        print(json.dumps({
            "score": score, "verdict": verdict, "chars": total_chars,
            "sentences": len(sents), "paragraphs": len(paras),
            "metrics": {k: {"score": round(v[0], 3), "note": v[1]} for k, v in sub.items()},
            "hits": hits, "patterns": pat_hits,
            "worst_sentences": worst,
        }, ensure_ascii=False, indent=1))
        return

    print(f"\n=== 降AI体检：{score}/100（{verdict}）===")
    print(f"字数 {total_chars} / 段落 {len(paras)} / 句子 {len(sents)}\n")
    for k, (v, note) in sub.items():
        bar = "█" * int(round(v * 10))
        print(f"{k:<6} {v*100:5.0f} {bar:<10} {note}")
    if hit_detail:
        print("\n【命中的 AI 词】")
        for w, c, kind in sorted(hit_detail, key=lambda x: -x[1])[:14]:
            print(f"  · {w} ×{c}（{kind}）")
    if pat_hits:
        print("\n【命中的模板结构】")
        for name, c in pat_hits:
            print(f"  · {name} ×{c}")
    print("\n【最该改的句子】")
    for m in worst:
        print(f"  第{m['para']}段第{m['sent']}句（{m['len']}字）：{m['text'][:60]}")
    print("\n下一步：按 references/rewrite-playbook.md 逐段改，改完再跑一次，目标 ≥85 分。")

if __name__ == "__main__":
    main()
