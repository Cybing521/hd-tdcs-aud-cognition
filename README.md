# 西瓜味 — HD-tDCS 对男性酒依赖患者认知功能的影响

原始临床研究稿件，目标期刊 **Brain Stimulation**（Original Research）。分析集：85 随机，77 例纳入主要分析。

## 目录

| 路径 | 内容 |
|---|---|
| `00_原始材料/` | 微信源 DOC |
| `01_论文正文/` | ZH / EN / BRS 三份 DOCX + 图 1–4 |
| `02_投稿材料/` | Cover letter、Highlights、投稿清单 |
| `03_analysis/` | 清洗、统计、出图、建稿脚本 |
| `03_分析表格/` | 含姓名的原始工作簿，**仅本地保留** |
| `data/deid/` | 去标识分析表（`study_id`）与 `DEID_REPORT.json` |
| `04_data/processed/` | 统计 JSON、敏感性、参考文献核验 |
| `HD-tDCS_AUD_cognition_delivery/` | 投稿交付包（BRS 正文、图、脚本、投稿附件） |
| `ccfa-review-reports/` | 主审稿报告 v3 + CONSORT / 参考文献 QC |
| `ccfa.yaml` | 工作流状态与 open gates |

## 数据与隐私

- 原始表（`03_分析表格/*.xlsx`）含受试者姓名，已写入 `.gitignore`，不上仓。
- 仓内分析表在 `data/deid/`。姓名已删除，行标识为 `study_id`。
- 清洗日志中的五名跨组歧义受试者编码为 P-A…P-E（见 `data/deid/ambiguous_participants.json` 与 `04_data/processed/data_cleaning_log.json`）。
- 若本地原始表更新：在保留 `03_分析表格/*.xlsx` 的前提下重新生成 `data/deid/`，确认输出不含 `姓名` 列后再提交。

## 当前状态（2026-08-25）

- 分析集叙事：85 随机 → 77 分析（含重归）已写入摘要 / 方法 / 图注
- 敏感性 S1–S4 已入文；结论稳健（ANCOVA P≤0.0003）
- 新颖性为可辩护窄口径（无 “first” 硬声明）
- 图 1–4：CONSORT / 蒙太奇 / 干预后结局 / 变化轨迹（TNR 黑白，PDF+PNG）
- 摘要 ≤250 词；BRS 正文约 3016 词
- 权威审稿报告：`ccfa-review-reports/manuscript_review_v3_2026-08-24.md`

## 重建

```bash
.venv/bin/python 03_analysis/clean_and_analyze.py
.venv/bin/python 03_analysis/sensitivity_analyses.py
.venv/bin/python 03_analysis/build_full_manuscript.py
```

`clean_and_analyze.py` 与 `sensitivity_analyses.py` 读取 `data/deid/`，不读取含姓名的原始表。

## 作者自行完成

- 作者信息、伦理批件号、基金、CRediT
- 回顾性试验注册号与注册日期
- 分配隐藏机制、电极尺寸/材料、AE 病历回顾
