# HD-tDCS and cognition in alcohol dependence

**Effects of High-Definition Transcranial Direct Current Stimulation on Cognitive Function in Male Inpatients with Alcohol Dependence: A Randomized Sham-Controlled Trial**

Single-center, single-blind, three-arm randomized trial (routine care / active HD-tDCS / sham HD-tDCS), Henan Psychiatric Hospital, Xinxiang, China, October 2019 – October 2020. 85 randomized; 77 analyzed (active 21 / sham 27 / control 29).

**Journal:** *Frontiers in Psychology* — Major Revision, manuscript ID **1308143**
**Repository:** <https://github.com/Cybing521/hd-tdcs-aud-cognition>

---

## What is in this repository

| Path | Contents |
|---|---|
| `data/deid/` | De-identified individual-level analysis tables (77 records × 24 variables, one record per participant), per-arm tables, the de-identification report, and the codes for the five participants affected by the record-verification step |
| `04_data/processed/` | Analysis outputs of the first delivery round (SPSS 22.0 re-derived in Python) |
| `05_Frontiers_修回材料/01_manuscript/` | Revised manuscript (tracked changes and clean), response letter, supplementary material, author information checklist |
| `05_Frontiers_修回材料/02_tables/` | All result tables in one workbook (9 sheets) |
| `05_Frontiers_修回材料/03_scripts/` | Every script that produces a reported number |
| `05_Frontiers_修回材料/04_results/` | Derived result tables (CSV/JSON) from the re-analysis, the manuscript number audit, and the pre-publication privacy scan |
| `01_论文正文/`, `02_投稿材料/`, `03_analysis/`, `HD-tDCS_AUD_cognition_delivery/` | First delivery round (Brain Stimulation submission), figures and scripts |
| `00_原始材料/` | The original Chinese manuscript (source of the intervention parameters) |
| `ccfa-review-reports/` | Internal manuscript QC reports |

## Data availability and de-identification

Participant names were removed at source and replaced with `study_id` pseudonyms; the name-to-code mapping was kept locally and is not part of this repository. The de-identification procedure, applied file by file, is recorded in `data/deid/DEID_REPORT.json` (source workbook, sheet, row/column counts, and the action taken for each field).

**Residual risk, stated plainly.** The de-identified tables retain six demographic and drinking variables (age, years of education, age at first drink, cumulative duration of use, dependence duration, prior abstinence attempts) together with five cognitive scores. In a sample of 77, that combination is quasi-identifying: anyone holding the source ward roster could in principle test candidate records. The `study_id` values are truncated local hashes and are pseudonyms, not an anonymization guarantee. Identifiable records remain with Henan Psychiatric Hospital and are not redistributed here. If the institutional ethics approval or the original consent does not cover sharing at this level of detail, use `05_Frontiers_修回材料/04_results/` and the scripts alone — the published inferential results are reproducible from the result tables without the participant-level file.

A pre-publication scan of every tracked file for direct identifiers (name/record-number/ID/phone/address/date-of-birth fields, long digit strings, date patterns) is archived as `05_Frontiers_修回材料/04_results/prepublication_privacy_scan.json`; its only hits are DOIs in the reference list.

## How to reproduce

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install numpy scipy statsmodels openpyxl lxml
python3 05_Frontiers_修回材料/03_scripts/recompute_all.py          # recomputes every reported statistic
python3 05_Frontiers_修回材料/03_scripts/verify_revisions.py      # checks the three manuscript versions against each other
python3 05_Frontiers_修回材料/03_scripts/audit_manuscript_numbers.py  # traces every number in the manuscript to a result table
```

Requirements: Python 3.14, SciPy 1.17, statsmodels 0.14, openpyxl 3.1, lxml. Random seeds are fixed (20261027), so the permutation and bootstrap results are exactly reproducible. `recompute_all.py` reads the master table from `data/deid/` and accepts either the archived file name or the ASCII name used in the standalone reproducibility package.

## Headline results and how they were checked

- All reported statistics were recomputed from the individual-level records in Python and cross-checked against the original SPSS 22.0 output: **109 of 109 checked values agree exactly** (`05_Frontiers_修回材料/04_results/上游一致性核对.csv`).
- Primary outcomes (TMT-B, VFT, HVLT-R): active HD-tDCS vs sham, Cohen's d = −0.96, 1.05, 1.30 (Hedges g −0.94, 1.03, 1.28); ANCOVA partial η² = 0.31, 0.19, 0.28.
- Robustness added at revision: permutation tests (10,000 permutations; omnibus P = 0.016 / 0.006 / <0.001; restricted permutation for the active–sham contrast P = 0.0016 / 0.0010 / 0.0002), bootstrap percentile CIs for the change-score difference (TMT-B −23.59 s [−40.43, −7.51]; VFT 2.89 [1.21, 4.64]; HVLT-R 4.39 [2.65, 6.22]), leave-one-participant-out re-analysis over all 77 records, and a worst-case (tipping-point) analysis for the three participants without a week-4 assessment.
- **Every number in the revised manuscript is traced to a result table** by `audit_manuscript_numbers.py` (46/46 targeted checks; 0 unmatched numeric tokens in a full-text scan).

## Citation

Please cite the manuscript itself:

> Effects of High-Definition Transcranial Direct Current Stimulation on Cognitive Function in Male Inpatients with Alcohol Dependence: A Randomized Sham-Controlled Trial. *Frontiers in Psychology* (in revision; manuscript ID 1308143).

## Contact

Correspondence: the corresponding author named in the manuscript, Department of Psychiatry, The Second Affiliated Hospital of Xinxiang Medical University (Henan Psychiatric Hospital), Xinxiang, Henan, China. Requests for identifiable data must go through the institution and be reviewed against the original ethics approval and consent.

## 中文说明

本仓库对应 *Frontiers in Psychology* 修回（稿件号 1308143）：`data/deid/` 是去标识后的个体级分析数据（77 例 × 24 变量，每人一条记录），`05_Frontiers_修回材料/` 放本轮修订的稿件三版、回复信、补充材料、全部可复算脚本与结果表。

**数据与隐私**：姓名在源头删除并替换为 `study_id` 伪名，映射只留本地；但表中仍保留 6 项人口学/饮酒变量与 5 项认知评分，在 77 例规模下属于准标识符——拿到住院花名册的人理论上可反查，故本仓库不构成绝对匿名化保证。含姓名原始表（`03_分析表格/*.xlsx`）由 `.gitignore` 排除，从未入库；公开前扫描结果见 `05_Frontiers_修回材料/04_results/prepublication_privacy_scan.json`。

**复现**：`recompute_all.py` 重算全部统计量并与上轮 SPSS 结果逐项核对（109/109 一致）；`verify_revisions.py` 校验修订稿三版互相对应；`audit_manuscript_numbers.py` 把正文每个数字指回结果表。
