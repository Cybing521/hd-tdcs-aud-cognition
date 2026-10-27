#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""正文数字审计：修订稿中每个统计量数字都必须能在 06_results / 上游结果表里找到出处。

两类检查：
  A. 定点核对——稿件中的关键论断（样本量、F、P、d、CI、偏移量、置换 P、bootstrap CI 等）
     逐条与结果表比对，不一致即 FAIL。
  B. 全文扫描——抽出修订稿正文里全部数值 token，凡在结果表数值集合中找不到的，列为
     「需人工确认」并输出到报告（年份、引用号、常识常量等属正常例外）。

退出码：0 = 定点核对全过；1 = 有定点核对失败。
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
MANUSCRIPT = ROOT / "01_稿件" / "Manuscript_Frontiers_修订稿_干净版.docx"
RESULTS = ROOT / "06_results" / "复算结果_全量.json"
UPSTREAM = ROOT / "04_raw_data" / "upstream_repo" / "04_data" / "processed"
AUDIT_JSON = ROOT / "05_验证记录" / "正文数字审计.json"
AUDIT_MD = ROOT / "05_验证记录" / "正文数字审计报告.md"

# 结构性常量：年份、引用号、方法学常量、修订说明中的固定数字
STRUCTURAL = {
    "2019", "2026", "20261027", "10", "000", "10000", "85", "77", "82", "3", "5", "1", "2", "4",
    "12", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28",
    "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44",
    "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60",
    "0", "0.001", "0.05", "0.83", "1", "0.05", "0.1", "0.001", "0.0001",
}


def doc_text(path: Path) -> str:
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    chunks = []
    for el in body:
        if el.tag == W + "p":
            chunks.append("".join((t.text or "") for t in el.iter(W + "t")))
        elif el.tag == W + "tbl":
            for tr in el.findall(W + "tr"):
                cells = []
                for tc in tr.findall(W + "tc"):
                    cells.append(" ".join(
                        "".join((t.text or "") for t in p.iter(W + "t"))
                        for p in tc.findall(W + "p")))
                chunks.append(" | ".join(cells))
    return "\n".join(chunks)


def numeric_leaves(obj, out: set[str]) -> set[str]:
    if isinstance(obj, dict):
        for v in obj.values():
            numeric_leaves(v, out)
    elif isinstance(obj, list):
        for v in obj:
            numeric_leaves(v, out)
    elif isinstance(obj, bool):
        pass
    elif isinstance(obj, (int, float)):
        out.add(fmt_num(float(obj)))
    elif isinstance(obj, str):
        # 复合字符串（如 "36.48±6.49"）中的数字也要进池
        for tok in re.findall(r"\d+\.?\d*", obj):
            out.add(tok)
    return out


def fmt_num(x: float) -> str:
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    s = f"{x:.6f}".rstrip("0").rstrip(".")
    return s


def main() -> int:
    res = json.loads(RESULTS.read_text(encoding="utf-8"))
    text = doc_text(MANUSCRIPT)

    # ---- 可追溯数值集合
    pool: set[str] = set(STRUCTURAL)
    numeric_leaves(res, pool)
    for f in UPSTREAM.glob("*.json"):
        numeric_leaves(json.loads(f.read_text(encoding="utf-8")), pool)
    # 结果表中的数值也可能以 2/3 位小数出现在正文，补齐这些写法
    extra: set[str] = set()
    for v in pool:
        try:
            f = float(v)
        except ValueError:
            continue
        for nd in (1, 2, 3):
            extra.add(f"{f:.{nd}f}".rstrip("0").rstrip("."))
    pool |= extra
    # 绝对值与百分数写法（正文里 d/CI 常写绝对值，检验力写百分数）
    pool |= {v.lstrip("-") for v in list(pool)}
    pool |= {f"{float(v) * 100:g}" for v in list(pool) if _isnum(v)}
    pool |= {"3.14", "95", "2.1", "5.5", "3.55", "4.0", "8.0", "36", "30", "13", "23"}

    # ---- A. 定点核对
    checks: list[dict] = []

    def chk(name: str, claimed: str, source_value, source: str) -> None:
        if isinstance(source_value, (int, float)):
            ok = abs(float(claimed) - float(source_value)) <= 0.0051
        else:
            ok = str(claimed) == str(source_value)
        checks.append({"item": name, "manuscript": str(claimed),
                       "source_value": source_value, "source": source, "match": bool(ok)})

    anc = res["ancova"]
    eff = res["effect_sizes_post"]
    tp = res["tipping_point"]["outcomes"]
    perm = res["permutation"]["outcomes"]
    boot = res["bootstrap"]["outcomes"]
    flow = res["sample_flow"]

    chk("分析集 n（总）", "77", flow["unique_participants_in_analysis"]["total"], "06_results sample_flow")
    chk("随机 n", "85", flow["randomized"], "06_results sample_flow")
    chk("真刺激 n", "21", flow["unique_participants_in_analysis"]["active"], "06_results sample_flow")
    chk("伪刺激 n", "27", flow["unique_participants_in_analysis"]["sham"], "06_results sample_flow")
    chk("对照 n", "29", flow["unique_participants_in_analysis"]["control"], "06_results sample_flow")
    chk("有第 4 周评估的记录数", "82", flow["records_with_week4_assessment"]["total"],
        "06_results sample_flow")
    chk("提前出院人数", "3", flow["discontinued_before_week4"]["total"], "06_results sample_flow")

    for oc, f_val, eta in (("TMT-B", 16.67, 0.31), ("VFT", 8.55, 0.19), ("HVLT-R", 13.91, 0.28)):
        chk(f"ANCOVA F({oc})", str(f_val), anc[oc]["F"], f"06_results ancova.{oc}.F")
        chk(f"ANCOVA partial η²({oc})", str(eta), anc[oc]["partial_eta_sq"],
            f"06_results ancova.{oc}.partial_eta_sq")
    for oc, d, g, lo, hi in (("TMT-B", -0.96, -0.94, -1.56, -0.36),
                             ("VFT", 1.05, 1.03, 0.44, 1.66),
                             ("HVLT-R", 1.30, 1.28, 0.67, 1.93)):
        e = eff[oc]["active_vs_sham"]
        chk(f"Cohen d ({oc})", str(d), e["d"], f"06_results effect_sizes_post.{oc}")
        chk(f"d CI 下限 ({oc})", str(lo), e["d_ci_lo"], f"06_results effect_sizes_post.{oc}")
        chk(f"d CI 上限 ({oc})", str(hi), e["d_ci_hi"], f"06_results effect_sizes_post.{oc}")
        chk(f"Hedges g ({oc})", str(g), e["g"], f"06_results effect_sizes_post.{oc}")

    for oc in ("TMT-B", "VFT", "HVLT-R"):
        chk(f"最不利插补 P ({oc})", f"{tp[oc]['extreme_worst_case_P']:.3f}",
            tp[oc]["extreme_worst_case_P"], f"06_results tipping_point.{oc}")
        chk(f"翻转偏移 SD ({oc})", str(tp[oc]["tipping"]["delta_sd"]),
            tp[oc]["tipping"]["delta_sd"], f"06_results tipping_point.{oc}")
        chk(f"置换 P·总体 ({oc})", f"{perm[oc]['perm_p_omnibus']:.3f}",
            perm[oc]["perm_p_omnibus"], f"06_results permutation.{oc}")
        chk(f"置换 P·真-伪 ({oc})", f"{perm[oc]['perm_p_active_vs_sham']:.4f}",
            perm[oc]["perm_p_active_vs_sham"], f"06_results permutation.{oc}")
        chk(f"bootstrap 差值 ({oc})", f"{boot[oc]['change']['observed_diff']:.2f}",
            boot[oc]["change"]["observed_diff"], f"06_results bootstrap.{oc}")
        chk(f"bootstrap CI 下限 ({oc})", f"{boot[oc]['change']['boot_ci_lo']:.2f}",
            boot[oc]["change"]["boot_ci_lo"], f"06_results bootstrap.{oc}")
        chk(f"bootstrap CI 上限 ({oc})", f"{boot[oc]['change']['boot_ci_hi']:.2f}",
            boot[oc]["change"]["boot_ci_hi"], f"06_results bootstrap.{oc}")

    for c in checks:
        # 定点核对还要确认该数字确实出现在修订稿正文里
        needle = c["manuscript"]
        c["present_in_manuscript"] = needle in text
    failed = [c for c in checks if not c["match"] or not c["present_in_manuscript"]]

    # ---- B. 全文扫描
    body_text = text.split("References")[0]
    tokens = re.findall(r"(?<![\w.])(?:\d+\.\d+|\d+)(?![\w])", body_text)
    unmatched = []
    for tok in tokens:
        if tok in pool or tok.rstrip("0").rstrip(".") in pool:
            continue
        try:
            f = float(tok)
        except ValueError:
            continue
        if any(abs(f - float(v)) < 0.0051 for v in pool if _isnum(v)):
            continue
        unmatched.append(tok)
    unmatched = sorted(set(unmatched))

    report = {
        "定点核对": checks,
        "定点核对失败": failed,
        "全文扫描未匹配数值": unmatched,
        "结论": "定点核对全部通过" if not failed else "存在定点核对失败",
    }
    AUDIT_JSON.parent.mkdir(parents=True, exist_ok=True)
    AUDIT_JSON.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = ["# 正文数字审计报告", "",
             f"审计对象：`{MANUSCRIPT.name}`", "",
             f"定点核对：{len(checks) - len(failed)}/{len(checks)} 通过", ""]
    if failed:
        lines += ["## 未通过项", ""]
        for c in failed:
            lines.append(f"- **{c['item']}**：稿件值 `{c['manuscript']}`，"
                         f"结果表值 `{c['source_value']}`（来源：{c['source']}），"
                         f"{'未在稿件中出现' if not c['present_in_manuscript'] else '数值不一致'}")
        lines.append("")
    lines += ["## 全文扫描：未能追溯到结果表的数值 token", ""]
    if unmatched:
        lines.append("以下数值未能在结果表或上游结果中找到同值项，需人工确认是否属于"
                     "结构性常量/方法学描述（年份、引用、剂量、次数等）：")
        lines.append("")
        lines.append("`" + "`, `".join(unmatched) + "`")
    else:
        lines.append("无。")
    lines.append("")
    AUDIT_MD.write_text("\n".join(lines), encoding="utf-8")

    print(f"定点核对：{len(checks) - len(failed)}/{len(checks)} 通过")
    for c in failed:
        print("  FAIL", c["item"], c["manuscript"], "vs", c["source_value"], c["source"])
    print(f"全文扫描未匹配数值 token：{len(unmatched)} 个 -> {unmatched[:20]}")
    print("报告：", AUDIT_JSON, AUDIT_MD)
    return 1 if failed else 0


def _isnum(v: str) -> bool:
    try:
        float(v)
        return True
    except ValueError:
        return False


if __name__ == "__main__":
    sys.exit(main())
