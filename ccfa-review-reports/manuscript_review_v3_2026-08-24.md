# Manuscript Review v3 — Independent Panel Audit (2026-08-24)

**Status update (2026-08-25):** Manuscript-side blockers B1–B3 and listed Majors/Minors under “我方可执行” have been applied in `build_full_manuscript.py` / `manuscript_content.py`; Figures 1–4 regenerated; three DOCX rebuilt; review archive consolidated to this file as the primary report. Author-only gates (registration, metadata, venue decision) remain open — see `ccfa.yaml` and `02_投稿材料/Submission_checklist.md`.

**Target (current):** Brain Stimulation (Elsevier) · **Manuscript:** BRS submission copy
**Panel:** 4 independent reviewer agents, run concurrently in isolation against the same authoritative snapshot; each report produced before seeing any other report. Coordinator (editorial) synthesis by the main agent. Full findings distilled below with reviewer-ID provenance; no finding added, dropped, or merged without provenance.

| Reviewer | Expertise | Recommendation | Confidence |
|---|---|---|---|
| R1-DESIGN | Clinical trial design; tDCS/HD-tDCS methodology | Major revision (reject risk if flow documentation not produced) | High |
| R2-STATS | Biostatistics, data integrity | Major revision | High (numeric), mod-high (pipeline) |
| R3-CONTRIBUTION | Addiction neuromodulation; journal fit | **Reject at Brain Stimulation as currently framed**; Major revision mandatory for any venue | High (venue fit, prior HD-tDCS AUD trials) |
| R4-REPORTING | CONSORT 2010, journal compliance, consistency | Major revision (reporting level) | High |

## Editorial verdict — 现在什么水平

**数字与合规层（强）**：R2 核对约 70 处统计锚点（摘要/正文/7 张表 vs JSON 底数）零偏差；BH-FDR 手工复算全部吻合；R4 独立验证摘要 228 词（限 250）、Highlights 5 条均 ≤85 字符、正文 2460 词（限 4000）、32 条参考文献严格按首次出现排序、期刊要求的 Integrity of the Blind 与 AI 声明均在。这一层达到可投稿水准。

**效度层（一票致命项，四位审稿人共识）**：
1. **分析集叙事不可重构**（R1 Blocker / R2 Major / R4 Major-可升级）：清洗日志权威口径 = 清洗前 27/26/29（82 条在档记录）→ 删除 5 条真刺激组误录/重复行（4 例跨组误录 + P-E 重复行 1 条）→ P-E 保留记录按原始随机化名单由真刺激组**重归**伪刺激组 → 21/27/29。正文从未陈述"重归"这一步（读者按正文算出 active=22 而非 21）；图 1 脚注 "One participant was retained in the sham arm" 措辞错误（应为 reassigned）；摘要 "77 male inpatients were assigned" 错误（85 随机、77 分析）。
2. **无敏感性分析**：删录与重归只发生在 active 组（−22%），恰是承载结论的组。→ **已完成**（见下）。
3. **新颖性声明与已发表文献冲突**（R3 Blocker）：已核实存在先前的 sham-controlled HD-tDCS 酒依赖试验（Nakamura-Palacios 2012 PMID 22956657 已核实；PMID 41780068 HD-tDCS-fMRS+naltrexone 已核实；Spagnolo ~2019 Front Psychiatry、Conti ~2024 报道于 BRS——细节待作者核实）。"First sham-controlled HD-tDCS trial of cognitive outcomes" 的窄口径可辩护，但三处 first 声明口径不一且未引用任何先前 HD-tDCS-AUD 试验，在本领域审稿人处等同失实。
4. **注册/伦理为占位符**（R1 M8 / R4 Blocker-pending）：2019 年入组已无法前瞻注册，回顾性注册（注明日期）是唯一诚实路径。

**战略层（R3，编辑采纳）**：BRS（IF 7.7，录取率 ~10%，70% 桌拒）当前框架下为纯行为结局、无机制载荷（无神经生理/影像）、未注册、事后数据修复——桌拒风险高。BRS 接受的成瘾-tDCS 工作共同点是机制/方法学载荷（本稿自引的 Dubuson 2021、Camchong 2023 皆如此）。**推荐替代：① Alcoholism: Clinical & Experimental Research（最佳契合：AUD 专科读者、重视神经心理终点）；② European Neuropsychopharmacology；③ Frontiers in Psychiatry（已发表过 sham-controlled HD-tDCS AUD 试验，编辑接受度有据）。→ 投稿目标变更需作者拍板。**

## 已完成的新证据：敏感性分析（R1 B1c + R2 Issue 2 要求）

`03_analysis/sensitivity_analyses.py` → `04_data/processed/sensitivity_analyses.json`。四个分析集（唯一受试者）：

| Set | n (A/S/C) | TMT-B ANCOVA F | VFT | HVLT-R | active-vs-sham d (TMT-B/VFT/HVLT-R) |
|---|---|---|---|---|---|
| S1 已发表 | 21/27/29 | 16.67*** | 8.55*** | 13.92*** | −0.96 / 1.05 / 1.30 |
| S2 P-E归真刺激 | 22/26/29 | 17.29*** | 9.30*** | 15.53*** | −1.07 / 1.11 / 1.46 |
| S3 剔除P-E | 21/26/29 | 17.01*** | 8.92*** | 14.94*** | −1.04 / 1.09 / 1.42 |
| S4 五争议例全归真刺激（最坏情形） | 26/22/29 | 13.30*** | 9.20*** | 9.15*** | −1.20 / 1.14 / 0.89 |

全部 P≤0.0003（ANCOVA）；Kruskal-Wallis 全部 P≤0.02。**结论在所有集合下存活，S1 精确复现已发表数字。** 另：21 vs 27 在 80% 效能下的最小可检测效应 d=0.83。

## 问题清单（去重，按严重度，含出处）

### Blocker / 必须先解决
| ID | 问题 | 处置 | 出处 |
|---|---|---|---|
| B1 | 分析集叙事：正文缺"重归"步骤、图 1 脚注措辞错、摘要 77/85 错误、需行级对账表 | 我方修复：正文+图+摘要重写；行级对账表入补充材料（需作者提供原始随机化名单日期/序号佐证） | R1-B1, R2-1, R4-1 |
| B2 | 敏感性分析缺失 | **已完成**（上表；写入正文 Statistical analysis + 补充） | R1-B1c, R2-2 |
| B3 | "First" 声明三处口径不一且不引先前 HD-tDCS-AUD 试验 | 我方修复：统一窄口径声明 + 引用已核实文献（Nakamura-Palacios 2012；PMID 41780068）；Spagnolo/Conti 待作者核实后再引 | R3-1 |
| B4 | 试验注册（CONSORT 23）占位 | 作者行动：回顾性注册并披露日期；删除"exemption"措辞 | R1-M8, R3-Obs, R4-3 |

### Major（稿件可修复）
| ID | 问题 | 处置 | 出处 |
|---|---|---|---|
| M1 | 评估者未盲、无检测偏倚缓解描述；VFT/HVLT-R 评分主观性 | 我方：改写 Integrity of the Blind + Limitations；盲法重评分敏感性分析需作者 | R1-M1, R3 |
| M2 | 伪刺激可信度/expectancy 从未评估（30s ramp-only 易破盲） | 我方：Limitations 明示；数据需作者 | R1-M2 |
| M3 | 随机序列生成与分配隐藏混写（"random number table" 是生成不是隐藏） | 我方：按 CONSORT 8–9 拆分改写；机制细节需作者 | R1-M3, R4 |
| M4 | 干预欠规范：回流电极位置/电极尺寸/电流密度/ramp/日程/依从性核查缺失 | 我方：回流电极（FP1,F7,C1,C3 / FP2,F8,C2,C4）等已有信息补入+蒙太奇图；电极尺寸等需作者 | R1-M4, R4-Obs18 |
| M5 | "routine care" 未定义、合并用药未报 | 需作者提供；我方留占位 | R1-M5, R4-Q2 |
| M6 | 脱毒完成标准与刺激起始时点未操作化 | 需作者 | R1-M6 |
| M7 | AE 报告无数字（建议回顾性病历审查） | 需作者 | R1-M7 |
| M8 | 伦理批件号/作者/基金/CRediT 占位 | 需作者 | R1-M8, R4 |
| M9 | 文献缺口：Nakamura-Palacios 2012 缺引；2023–26 NIBS-成瘾 meta；tDCS+认知训练；执行缺陷→复发前瞻研究 | 我方：补已核实文献；其余按主题缺口在 Discussion 承认 | R3-2 |
| M10 | 结论框架倒置："beyond routine care"（最弱对比）置于 sham 对比之前 | 我方：结论/Highlights/Discussion 以 sham-controlled 为主对比，"greater improvement" 替代 "improved" | R3 |
| M11 | 机制选择性论证为反向推断（HVLT-R 颞叶-边缘依赖与 DLPFC 叙事冲突） | 我方：改写为描述性 + 承认行为数据不能定位机制 | R3 |

### Minor（稿件可修复，批量处理）
摘要 q<0.05 误挂 pairwise（R2-3/R4-7）；Hedges g 承诺未报告（R2-5/R4-8，g=−0.94/1.03/1.28 已有）；F 统计量无 df、Welch 未标注（R2-4）；post-hoc power 循环论证→改 MDE d=0.83（R1-m7/R2-6）；Levene P 未报告+正态性筛查缺失→补 Levene P+KW 敏感性（R2-7，已算）；change/ANCOVA 家族未校正需声明 nominal（R2-8）；MMSE active-vs-sham d=0.59 CI 不含零需限定 omnibus 口径（R2-9）；η² 0.313→0.31（R2-10/R4-6）；Table 7 正文未引用（R4-4）；图表题注重复且措辞漂移（R4-5）；引用格式 (1-2)→[1,2]（R4-9）；>6 作者 et al. 规则（R4-10）；"85 screened" 混淆三重分母（R4-11）；无中期分析/终止声明（R4-12）；CONSORT 图 300dpi PNG 不达标→矢量 PDF 为主文件、脚注移出图内（R4-13）；表 2/3/5 冗余→合并（R4-14）；表内小数位统一（R4-15）；正文内部路径引用（R4-16）；consort_checklist.md 自身错误（R4-17）；complete-case 标注（R1-m2）；single-blind 精确化（R1-m8）；测验版本/施测细节（R1-m3，需作者）；guardian co-consent 解释（R1-M8/R4-Obs22）。

### 分歧（保留）
R3 判 "Reject at BRS" vs R1/R2/R4 判 "Major revision"——非事实分歧：同一缺陷集在专科期刊读作大修，在 BRS 门槛下读作桌拒风险。编辑采纳：缺陷修复对任何 venue 都是前置条件；venue 选择独立决策（见战略层）。

## 修订路线图

**我方可执行（进行中）**：B1 叙事重写、B2 敏感性入文、B3 新颖性改写+补引、M1–M4/M9–M11 及全部 Minor 的稿件侧修复、4 图（CONSORT 重绘 TNR 矢量 + HD-tDCS 4×1 蒙太奇示意 + 主结局箱线+散点 + 变化值/ANCOVA 图）、表 2/3/5 合并、重建三份 docx、consort_checklist 修正。
**作者必须完成（不可代笔）**：B4 注册、伦理批件号、作者/单位/通讯、基金、CRediT、routine care 与合并用药描述、脱毒操作化定义、评估者与刺激操作者关系、电极尺寸/材料、伪刺激 ramp 参数、AE 回顾性审查、Spagnolo/Conti 文献核实、guardian consent 说明、**投稿目标拍板（BRS vs ACER vs 其他）**。
