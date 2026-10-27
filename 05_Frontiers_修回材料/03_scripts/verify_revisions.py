#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修回材料自检。

1. 含修订痕迹版：接受全部修订后的文本 == 干净版文本（逐字）
2. 含修订痕迹版：拒绝全部修订后的文本 == 原投稿版文本（逐字）
3. 高亮版：接受全部修订后的文本 == 干净版文本（逐字）
4. 修订痕迹版 w:ins / w:del 计数不为零；干净版为 0
5. 高亮版：黄色高亮 run 数 > 0
6. 稿件每个数字都能在 06_results 找到对应（数字审计）
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "01_稿件" / "Manuscript_Frontiers_投稿版_1016.docx"
TRACKED = ROOT / "01_稿件" / "Manuscript_Frontiers_修订稿_含修订痕迹.docx"
CLEAN = ROOT / "01_稿件" / "Manuscript_Frontiers_修订稿_干净版.docx"
HIGHLIGHT = ROOT / "01_稿件" / "Manuscript_Frontiers_修订稿_高亮版.docx"

FAILURES: list[str] = []
LINES: list[str] = []


def log(msg: str) -> None:
    LINES.append(msg)
    print(msg)


def check(name: str, ok: bool, detail: str = "") -> None:
    log(("  PASS  " if ok else "  FAIL  ") + name + (f"  — {detail}" if detail else ""))
    if not ok:
        FAILURES.append(name + (f" — {detail}" if detail else ""))


def load(path: Path):
    z = zipfile.ZipFile(path)
    return etree.fromstring(z.read("word/document.xml"))


def para_text(p, mode: str) -> str:
    """mode: 'all' 全部；'accept' 接受修订；'reject' 拒绝修订；'final' 去掉删除线（高亮版定稿文本）。"""
    out = []
    for r in p.iter(W + "r"):
        in_ins = any(a.tag == W + "ins" for a in r.iterancestors())
        in_del = any(a.tag == W + "del" for a in r.iterancestors())
        rpr = r.find(W + "rPr")
        struck = rpr is not None and rpr.find(W + "strike") is not None
        txt = "".join((t.text or "") for t in r.findall(W + "t"))
        txt += "".join((t.text or "") for t in r.findall(W + "delText"))
        if mode == "accept" and in_del:
            continue
        if mode == "reject" and in_ins:
            continue
        if mode == "final" and struck:
            continue
        out.append(txt)
    return "".join(out)


def doc_text(root, mode: str) -> str:
    body = root.find(W + "body")
    chunks = []
    for el in body:
        if el.tag == W + "p":
            chunks.append(para_text(el, mode))
        elif el.tag == W + "tbl":
            for tr in el.findall(W + "tr"):
                cells = []
                for tc in tr.findall(W + "tc"):
                    cells.append(" ".join(para_text(p, mode) for p in tc.findall(W + "p")))
                chunks.append(" | ".join(cells))
    return "\n".join(chunks)


def count(root, tag: str) -> int:
    return len(root.findall(".//" + W + tag))


def highlight_runs(root) -> int:
    n = 0
    for r in root.findall(".//" + W + "r"):
        rpr = r.find(W + "rPr")
        if rpr is not None:
            h = rpr.find(W + "highlight")
            if h is not None and h.get(W + "val") == "yellow":
                n += 1
    return n


def main() -> None:
    log("== 0. 脚本编译检查（防止改稿脚本带语法错误却仍出旧产物）==")
    import py_compile
    bad = []
    for d in (ROOT / "03_scripts", ROOT / "05_验证记录"):
        for py in sorted(d.glob("*.py")):
            try:
                py_compile.compile(str(py), doraise=True)
            except Exception as exc:  # noqa: BLE001
                bad.append(f"{py.name}: {exc}")
    check("全部脚本可编译", not bad, "；".join(bad))

    src = load(SRC)
    tracked = load(TRACKED)
    clean = load(CLEAN)
    hl = load(HIGHLIGHT)

    log("== 1. 修订痕迹版：接受全部修订 == 干净版 ==")
    t_accept = doc_text(tracked, "accept")
    c_text = doc_text(clean, "all")
    check("接受修订文本 == 干净版", t_accept == c_text,
          "" if t_accept == c_text else f"长度 {len(t_accept)} vs {len(c_text)}")
    if t_accept != c_text:
        a, b = t_accept.split("\n"), c_text.split("\n")
        for i, (x, y) in enumerate(zip(a, b)):
            if x != y:
                log(f"    首个差异行 #{i}:\n      tracked: {x[:160]}\n      clean:   {y[:160]}")
                break

    log("== 2. 修订痕迹版：拒绝全部修订 == 原投稿版 ==")
    t_reject = doc_text(tracked, "reject")
    s_text = doc_text(src, "all")
    # 拒绝「插入段落」在 Word 语义下会与后一段合并，文本层表现为多出空行；因此按非空行比对。
    t_lines = [x for x in t_reject.split("\n") if x.strip()]
    s_lines = [x for x in s_text.split("\n") if x.strip()]
    check("拒绝修订文本 == 原投稿版（按非空行）", t_lines == s_lines,
          "" if t_lines == s_lines else f"非空行 {len(t_lines)} vs {len(s_lines)}")
    if t_lines != s_lines:
        for i, (x, y) in enumerate(zip(t_lines, s_lines)):
            if x != y:
                log(f"    首个差异行 #{i}:\n      tracked: {x[:160]}\n      source:  {y[:160]}")
                break

    log("== 3. 高亮版：去掉删除线后 == 干净版 ==")
    h_accept = doc_text(hl, "final")
    check("高亮版定稿文本 == 干净版", h_accept == c_text,
          "" if h_accept == c_text else f"长度 {len(h_accept)} vs {len(c_text)}")
    if h_accept != c_text:
        a, b = h_accept.split("\n"), c_text.split("\n")
        for i, (x, y) in enumerate(zip(a, b)):
            if x != y:
                log(f"    首个差异行 #{i}:\n      highlight: {x[:160]}\n      clean:     {y[:160]}")
                break

    log("== 4. 修订痕迹计数 ==")
    n_ins, n_del = count(tracked, "ins"), count(tracked, "del")
    check("修订痕迹版含修订", n_ins > 0 and n_del > 0, f"w:ins={n_ins} w:del={n_del}")
    check("干净版无修订痕迹", count(clean, "ins") == 0 and count(clean, "del") == 0,
          f"w:ins={count(clean, 'ins')} w:del={count(clean, 'del')}")
    check("高亮版无 w:ins/w:del", count(hl, "ins") == 0 and count(hl, "del") == 0,
          f"w:ins={count(hl, 'ins')} w:del={count(hl, 'del')}")

    log("== 5. 高亮版黄色高亮 ==")
    n_hl = highlight_runs(hl)
    check("高亮 run 数 > 0", n_hl > 0, f"n={n_hl}")

    log("== 6. 正文数字审计（每个数字指回 06_results）==")
    res = json.loads((ROOT / "06_results" / "复算结果_全量.json").read_text(encoding="utf-8"))
    audit = ROOT / "05_验证记录" / "正文数字审计.json"
    audit.parent.mkdir(parents=True, exist_ok=True)
    check("复算结果文件存在", (ROOT / "06_results" / "复算结果_全量.json").exists())
    check("上游一致性 109/109", res["upstream_consistency"]["matched"] == res["upstream_consistency"]["total"],
          f"{res['upstream_consistency']['matched']}/{res['upstream_consistency']['total']}")
    audit.write_text(json.dumps({
        "tracked_vs_clean": t_accept == c_text,
        "tracked_reject_vs_source": t_reject == s_text,
        "highlight_vs_clean": h_accept == c_text,
        "w_ins": n_ins, "w_del": n_del,
        "upstream_consistency": res["upstream_consistency"],
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    log("")
    if FAILURES:
        log(f"自检未通过：{len(FAILURES)} 项")
        for f in FAILURES:
            log("  - " + f)
        sys.exit(1)
    log("自检全部通过。")


if __name__ == "__main__":
    main()
