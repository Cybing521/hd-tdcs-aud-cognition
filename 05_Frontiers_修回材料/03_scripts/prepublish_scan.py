#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公开前扫描：逐个跟踪文件排查直接标识符与准标识符残留。

检查项：
  1. 表格里的标识符列名（姓名/住院号/门诊号/身份证/电话/住址/出生日期/病案号…）
  2. 单元格中的长数字串（≥6 位连续数字，身份证/住院号形态）
  3. 中文姓名形态的连续 2–4 个汉字（仅作提示，需人工判读）
  4. 日期形态（YYYY-MM-DD / YYYY年M月D日）
  5. 正文文本里的上述线索（docx / txt / md / py / json / csv）

输出：05_验证记录/公开前扫描报告.json
"""

from __future__ import annotations

import json
import re
import subprocess
import zipfile
from pathlib import Path

import openpyxl
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT / "04_raw_data" / "upstream_repo"
OUT = ROOT / "05_验证记录" / "公开前扫描报告.json"

IDENT_HEADER = ["姓名", "名字", "住院号", "门诊号", "身份证", "电话", "手机", "住址", "地址",
                "出生日期", "生日", "病案号", "床号", "单位", "家属", "联系人", "性别年龄", "年龄性别"]
LONG_NUM = re.compile(r"(?<!\d)\d{6,}(?!\d)")
DATE_LIKE = re.compile(r"(19|20)\d{2}\s*[-/年]\s*\d{1,2}\s*[-/月]\s*\d{1,2}")
HAN_NAME = re.compile(r"[一-龥]{2,4}")

TEXT_EXT = {".txt", ".md", ".py", ".json", ".csv", ".yml", ".yaml"}


def tracked_files() -> list[Path]:
    out = subprocess.run(["git", "-C", str(REPO), "ls-files", "-z"],
                         capture_output=True, text=True, check=True).stdout
    return [REPO / p for p in out.split("\0") if p]


def docx_text(p: Path) -> str:
    with zipfile.ZipFile(p) as z:
        if "word/document.xml" not in z.namelist():
            return ""
        root = etree.fromstring(z.read("word/document.xml"))
        parts = []
        for t in root.iter(W + "t"):
            parts.append(t.text or "")
            parts.append(" ")
        return "".join(parts)


def scan_text(name: str, text: str) -> list[dict]:
    hits = []
    for kw in IDENT_HEADER:
        for m in re.finditer(re.escape(kw), text):
            ctx = re.sub(r"\s+", " ", text[max(0, m.start() - 35):m.end() + 35])
            hits.append({"kind": "标识符关键词", "detail": f"{kw} | …{ctx}…"})
    for m in set(LONG_NUM.findall(text)):
        hits.append({"kind": "长数字串", "detail": m[:4] + "*" * (len(m) - 4)})
    for m in set(DATE_LIKE.findall(text)):
        hits.append({"kind": "日期形态", "detail": str(m)})
    return hits


def scan_xlsx(p: Path) -> list[dict]:
    hits = []
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            header = []
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                header = [str(c) for c in row if c is not None]
            for h in header:
                for kw in IDENT_HEADER:
                    if kw in h:
                        hits.append({"kind": "标识符列名", "detail": f"{ws.title}:{h}"})
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if v is None:
                        continue
                    s = str(v)
                    for m in set(LONG_NUM.findall(s)):
                        hits.append({"kind": "长数字串", "detail": f"{ws.title}:{m[:4]}{'*' * (len(m) - 4)}"})
                    if re.fullmatch(r"[一-龥]{2,4}", s) and s not in ("姓名",):
                        hits.append({"kind": "中文姓名形态（需人工判读）", "detail": f"{ws.title}:{s}"})
    finally:
        wb.close()
    return hits


def main() -> int:
    report = {"scanned": [], "findings": {}, "quasi_identifier_note": None}
    for p in tracked_files():
        rel = str(p.relative_to(REPO))
        report["scanned"].append(rel)
        hits: list[dict] = []
        try:
            if p.suffix.lower() == ".xlsx":
                hits = scan_xlsx(p)
            elif p.suffix.lower() == ".docx":
                hits = scan_text(rel, docx_text(p))
            elif p.suffix.lower() in TEXT_EXT:
                hits = scan_text(rel, p.read_text(encoding="utf-8", errors="ignore"))
            elif p.suffix.lower() == ".doc":
                # 旧版二进制 doc：只报「需人工判读」，不解析
                hits = [{"kind": "旧版 .doc（未解析）", "detail": rel}]
        except Exception as exc:  # noqa: BLE001
            hits = [{"kind": "扫描异常", "detail": f"{rel}: {exc}"}]
        if hits:
            report["findings"][rel] = hits

    report["quasi_identifier_note"] = (
        "去标识表保留 6 项人口学/饮酒变量（年龄、教育年限、首次饮酒年龄、累计使用时间、"
        "成瘾时间、戒酒次数）与 5 项认知评分。在 77 例规模下，这些准标识符组合理论上存在"
        "再识别可能。公开前请确认机构伦理与知情同意允许共享该层级数据；若不允许，"
        "只公开结果表与脚本（见 results/），不开 data/deid/。"
    )
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"扫描 {len(report['scanned'])} 个跟踪文件")
    if not report["findings"]:
        print("未发现直接标识符线索。")
    for rel, hits in report["findings"].items():
        print(f"\n{rel}")
        for h in hits[:12]:
            print("   ", h["kind"], "→", h["detail"])
        if len(hits) > 12:
            print(f"    …共 {len(hits)} 条")
    print("\n准标识符提示：\n" + report["quasi_identifier_note"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
