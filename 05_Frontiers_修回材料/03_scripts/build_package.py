#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""组装交付包与可复现性包。

07_交付包/  —— 给作者的投稿材料包（打包唯一输入）
  01_manuscript/  修订稿（含修订痕迹 / 干净版）、回复信、补充材料
  02_tables/      全部结果表合并为一个 xlsx
  03_scripts/     全部可复算脚本
  包内全 ASCII；不放 manifest.csv，不放任何 md / txt 说明文件。

07_交付/  —— 报告与导出件：两个 zip、PDF 渲染件、说明文档（包外）
"""

from __future__ import annotations

import csv
import hashlib
import json
import shutil
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side

ROOT = Path(__file__).resolve().parents[1]
PACK = ROOT / "07_交付包"
DELIV = ROOT / "07_交付"
RESULTS = ROOT / "06_results"
REPO = ROOT / "04_raw_data" / "upstream_repo"

# 结果表 -> (CSV 文件, 工作表名, 表题)
TABLE_SPECS = [
    ("表1_人口学与饮酒特征.csv", "Table1_baseline", "Table 1. Sociodemographic and drinking characteristics (mean±SD)"),
    ("表2_认知评分_基线与干预后.csv", "Table2_cognition", "Table 2. Cognitive scores at baseline and post-intervention (mean±SD)"),
    ("表3_变化量.csv", "Table3_change", "Table 3. Baseline-to-post-intervention change scores (mean Δ±SD)"),
    ("表4_ANCOVA.csv", "Table4_ancova", "Table 4. Analysis of covariance on post-intervention cognitive scores"),
    ("新增A_tipping_point.csv", "S3_worst_case", "Table S3. Worst-case analysis for the three participants without a week-4 assessment"),
    ("新增B_置换检验.csv", "S4_permutation", "Table S4. Permutation tests for the primary contrasts"),
    ("新增C_bootstrap.csv", "S4_bootstrap", "Table S4 (continued). Bootstrap confidence intervals for the primary contrasts"),
    ("新增E_诊断.csv", "S5_diagnostics", "Table S5. Distribution diagnostics and leave-one-participant-out stability"),
    ("上游一致性核对.csv", "Upstream_check", "Cross-check of every recomputed value against the previous delivery"),
]

THIN = Side(style="thin", color="000000")
MEDIUM = Side(style="medium", color="000000")
NO_SIDE = Side(style=None)


def read_csv(path: Path) -> list[dict]:
    with path.open(encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def build_result_workbook(out: Path) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for csv_name, sheet, title in TABLE_SPECS:
        rows = read_csv(RESULTS / csv_name)
        if not rows:
            continue
        ws = wb.create_sheet(sheet)
        ws.cell(row=1, column=1, value=title).font = Font(name="Times New Roman", size=10, bold=True)
        header = list(rows[0].keys())
        for j, h in enumerate(header, start=1):
            c = ws.cell(row=3, column=j, value=h)
            c.font = Font(name="Times New Roman", size=10, bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(top=MEDIUM, bottom=THIN, left=NO_SIDE, right=NO_SIDE)
        for i, r in enumerate(rows, start=4):
            for j, h in enumerate(header, start=1):
                v = r.get(h, "")
                if v is None:
                    v = ""
                try:
                    v = float(v) if (isinstance(v, str) and v.strip() != "" and
                                     v.replace(".", "", 1).replace("-", "", 1).isdigit()) else v
                except ValueError:
                    pass
                c = ws.cell(row=i, column=j, value=v)
                c.font = Font(name="Times New Roman", size=10)
                c.alignment = Alignment(horizontal="left" if j == 1 else "center",
                                        vertical="center", wrap_text=True)
                if i == 3 + len(rows):
                    c.border = Border(bottom=MEDIUM, left=NO_SIDE, right=NO_SIDE)
        for j, h in enumerate(header, start=1):
            width = max(12, min(60, int(len(str(h)) * 1.2) + 4))
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = width
        ws.freeze_panes = "A4"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def copy_ascii(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def assemble_pack() -> list[Path]:
    if PACK.exists():
        shutil.rmtree(PACK)
    files: list[Path] = []

    pairs = [
        ("01_稿件/Manuscript_Frontiers_修订稿_含修订痕迹.docx", "01_manuscript/Manuscript_Revised_Tracked_Changes.docx"),
        ("01_稿件/Manuscript_Frontiers_修订稿_干净版.docx", "01_manuscript/Manuscript_Revised_Clean.docx"),
        ("01_稿件/修订记录/Response_to_Reviewers.docx", "01_manuscript/Response_to_Reviewers.docx"),
        ("01_稿件/修订记录/Supplementary_Material.docx", "01_manuscript/Supplementary_Material.docx"),
    ]
    for s, d in pairs:
        copy_ascii(ROOT / s, PACK / d)
        files.append(PACK / d)

    build_result_workbook(PACK / "02_tables" / "Result_Tables.xlsx")
    files.append(PACK / "02_tables" / "Result_Tables.xlsx")

    for s, d in [
        ("03_scripts/recompute_all.py", "03_scripts/recompute_all.py"),
        ("03_scripts/revision_content.py", "03_scripts/revision_content.py"),
        ("03_scripts/build_revision.py", "03_scripts/build_revision.py"),
        ("03_scripts/docx_edit.py", "03_scripts/docx_edit.py"),
        ("03_scripts/verify_revisions.py", "03_scripts/verify_revisions.py"),
        ("03_scripts/dump_docx_paragraphs.py", "03_scripts/dump_docx_paragraphs.py"),
        ("05_验证记录/数字审计.py", "03_scripts/audit_manuscript_numbers.py"),
    ]:
        copy_ascii(ROOT / s, PACK / d)
        files.append(PACK / d)
    return files


def assemble_repro_package() -> Path:
    """可复现性包：脚本 + 去标识数据 + 说明 + 校验和（供期刊与 GitHub 使用）。"""
    rp = DELIV / "reproducibility_package"
    if rp.exists():
        shutil.rmtree(rp)
    (rp / "data" / "deid").mkdir(parents=True, exist_ok=True)
    (rp / "scripts").mkdir(parents=True, exist_ok=True)
    (rp / "results").mkdir(parents=True, exist_ok=True)

    # 去标识数据：文件名转 ASCII，保留原名对照
    rename = {
        "比较_deid.xlsx": "comparison_deid.xlsx",
        "真刺激组_deid.xlsx": "active_arm_deid.xlsx",
        "真刺激组比较_deid.xlsx": "active_arm_comparison_deid.xlsx",
        "伪刺激组_deid.xlsx": "sham_arm_deid.xlsx",
        "伪刺激组比较_deid.xlsx": "sham_arm_comparison_deid.xlsx",
        "酒依赖组_deid.xlsx": "control_arm_deid.xlsx",
        "酒依赖比较_deid.xlsx": "control_arm_comparison_deid.xlsx",
        "ambiguous_participants.json": "ambiguous_participants.json",
        "DEID_REPORT.json": "DEID_REPORT.json",
    }
    src_deid = REPO / "data" / "deid"
    for s, d in rename.items():
        copy_ascii(src_deid / s, rp / "data" / "deid" / d)

    for s in ("recompute_all.py", "revision_content.py", "build_revision.py", "docx_edit.py",
              "verify_revisions.py"):
        copy_ascii(ROOT / "03_scripts" / s, rp / "scripts" / s)
    copy_ascii(ROOT / "05_验证记录" / "数字审计.py", rp / "scripts" / "audit_manuscript_numbers.py")

    for p in sorted(RESULTS.glob("*.csv")):
        copy_ascii(p, rp / "results" / p.name)
    copy_ascii(RESULTS / "复算结果_全量.json", rp / "results" / "recomputed_results.json")

    readme = rp / "README.md"
    readme.write_text(
        "# Reproducibility package — HD-tDCS for cognition in alcohol dependence\n\n"
        "Manuscript: Effects of High-Definition Transcranial Direct Current Stimulation on Cognitive\n"
        "Function in Male Inpatients with Alcohol Dependence: A Randomized Sham-Controlled Trial\n\n"
        "## Contents\n\n"
        "| Path | Description |\n|---|---|\n"
        "| `data/deid/comparison_deid.xlsx` | De-identified analysis dataset, 77 records x 24 variables, one record per participant |\n"
        "| `data/deid/active_arm_deid.xlsx` | Active HD-tDCS arm (n=21) |\n"
        "| `data/deid/sham_arm_deid.xlsx` | Sham HD-tDCS arm (n=27) |\n"
        "| `data/deid/control_arm_deid.xlsx` | Routine-care arm (n=29) |\n"
        "| `data/deid/*_comparison_deid.xlsx` | Per-arm comparison tables with the drinking and cognitive variables |\n"
        "| `data/deid/ambiguous_participants.json` | Codes for the five participants affected by the record-verification step |\n"
        "| `data/deid/DEID_REPORT.json` | De-identification report: source file, sheet, row and column counts, and the action applied to each field |\n"
        "| `scripts/recompute_all.py` | Recomputation of every reported statistic plus the added robustness analyses |\n"
        "| `scripts/revision_content.py` | Manuscript text generated from the result tables |\n"
        "| `scripts/build_revision.py` | Builds the revised manuscript, response letter and supplementary material |\n"
        "| `scripts/docx_edit.py` | DOCX editing helpers (clean / tracked-changes / highlight output) |\n"
        "| `scripts/verify_revisions.py` | Self-check: accept-all-revisions equals the clean version, reject-all equals the submitted version |\n"
        "| `scripts/audit_manuscript_numbers.py` | Number audit: every statistic in the manuscript is traced to a result table |\n"
        "| `results/*.csv` | Derived result tables |\n"
        "| `results/recomputed_results.json` | Full recomputation output, including the upstream consistency check |\n"
        "| `checksums.md5` | SHA-256 checksums of all files in this package |\n\n"
        "## How to run\n\n"
        "```bash\n"
        "python3 -m venv .venv && . .venv/bin/activate\n"
        "pip install numpy scipy statsmodels openpyxl lxml\n"
        "python3 scripts/recompute_all.py        # regenerates results/ and prints the upstream consistency check\n"
        "python3 scripts/verify_revisions.py      # checks the three manuscript versions against each other\n"
        "python3 scripts/audit_manuscript_numbers.py  # traces every number in the manuscript to a result table\n"
        "```\n\n"
        "The scripts expect the de-identified tables in `data/deid/` and write their outputs to\n"
        "`results/`. Seeds are fixed (20261027), so the permutation and bootstrap results are exactly\n"
        "reproducible.\n\n"
        "## Filename mapping between the archived source workbooks and this package\n\n"
        "| Archived name | Package name |\n|---|---|\n"
        + "\n".join(f"| `{k}` | `data/deid/{v}` |" for k, v in rename.items())
        + "\n\n## De-identification\n\nParticipant names were removed at source and replaced with hashed study identifiers.\n"
        "The name-to-code mapping was kept locally and is not part of this package. No direct identifier,\n"
        "date of birth, or free-text clinical note is present in the de-identified tables. Identifiable\n"
        "records remain with the originating institution and are not redistributed.\n",
        encoding="utf-8")

    # 校验和
    lines = []
    for p in sorted(rp.rglob("*")):
        if p.is_file() and p.name != "checksums.md5":
            h = hashlib.sha256(p.read_bytes()).hexdigest()
            lines.append(f"{h}  {p.relative_to(rp)}")
    (rp / "checksums.md5").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return rp


def zip_dir(src: Path, out: Path, strip_prefix: str = "") -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(src.rglob("*")):
            if p.is_file():
                rel = p.relative_to(src)
                arc = str(rel)
                if strip_prefix:
                    arc = arc[len(strip_prefix):].lstrip("/")
                z.write(p, arc)


def main() -> None:
    DELIV.mkdir(parents=True, exist_ok=True)
    files = assemble_pack()
    print(f"交付包文件数：{len(files)}")

    zip_pack = DELIV / "交付包_hd_tdcs_aud_cognition_20261027.zip"
    zip_dir(PACK, zip_pack)
    print("zip:", zip_pack, zip_pack.stat().st_size, "bytes")

    rp = assemble_repro_package()
    zip_rp = DELIV / "reproducibility_package_hd_tdcs_aud_cognition_20261027.zip"
    zip_dir(rp, zip_rp)
    print("zip:", zip_rp, zip_rp.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
