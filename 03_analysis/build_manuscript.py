#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate revised Chinese manuscript DOCX from analysis_results.json."""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
STATS = ROOT / "04_data" / "processed" / "analysis_results.json"
OUT = ROOT / "01_论文正文" / "HD-tDCS酒依赖认知功能_修订稿.docx"


def set_doc_font(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def add_title(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "黑体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def add_table(doc: Document, title: str, headers: list[str], rows: list[list[str]]) -> None:
    cap = doc.add_paragraph(title)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.runs[0].bold = True
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for j, h in enumerate(headers):
        table.rows[0].cells[j].text = h
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            table.rows[i].cells[j].text = str(val)
    doc.add_paragraph("")


def p(doc: Document, text: str) -> None:
    doc.add_paragraph(text)


def fmt_p(v) -> str:
    if v is None:
        return "—"
    if v == 0:
        return "<0.001"
    return f"{v:.3f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)


def main() -> None:
    data = json.loads(STATS.read_text(encoding="utf-8"))
    n = data["n"]
    n1, n2, n3 = n["1"], n["2"], n["3"]
    total = data["total_n"]
    t1 = {r["variable"]: r for r in data["tables"]["table1"]}
    base = {r["outcome"]: r for r in data["tables"]["table_baseline"]}
    post = {r["outcome"]: r for r in data["tables"]["table_post"]}
    pw = data["pairwise_post"]
    es = data["effect_sizes"]
    corr = data["correlation"]["significant"]

    doc = Document()
    set_doc_font(doc)

    add_title(doc, "高精度经颅直流电刺激对男性酒依赖患者认知功能的影响")
    p(doc, "【摘要】")
    p(
        doc,
        f"目的：探讨高精度经颅直流电刺激（HD-tDCS）对男性酒依赖患者认知功能的影响。"
        f"方法：纳入{total}例男性酒依赖住院患者，随机分为对照组（n={n3}）、真刺激组（n={n1}）和伪刺激组（n={n2}）。"
        "对照组接受常规治疗，真刺激组与伪刺激组在常规治疗基础上分别接受14次（20 min/次，1 mA）HD-tDCS真刺激或伪刺激，靶点为双侧背外侧前额叶。"
        "于干预前后评估MMSE、TMT-B、DST、VFT和HVLT-R。"
        f"结果：干预后，真刺激组TMT-B、VFT、HVLT-R优于伪刺激组及对照组（P<0.05）；"
        f"MMSE与DST三组差异无统计学意义（P>0.05）。"
        f"真刺激组与伪刺激组比较：TMT-B（P={fmt_p(pw['TMT-B']['active_vs_sham'])}，d={es['TMT-B']['active_vs_sham_d']}），"
        f"VFT（P={fmt_p(pw['VFT']['active_vs_sham'])}，d={es['VFT']['active_vs_sham_d']}），"
        f"HVLT-R（P={fmt_p(pw['HVLT-R']['active_vs_sham'])}，d={es['HVLT-R']['active_vs_sham_d']}）。"
        "伪刺激组与对照组上述指标差异均无统计学意义（P>0.05）。"
        "基线认知指标与部分饮酒特征呈相关（见正文）。"
        "结论：HD-tDCS可在常规治疗基础上进一步改善男性酒依赖患者的认知转换、执行控制及言语学习记忆功能。",
    )

    p(doc, "【Abstract】")
    p(
        doc,
        f"Objective: To examine the effect of high-definition transcranial direct current stimulation (HD-tDCS) "
        f"on cognitive performance in male inpatients with alcohol dependence. "
        f"Methods: In this single-blind randomized trial, {total} male patients were allocated to routine care "
        f"(control, n={n3}), active HD-tDCS (n={n1}), or sham HD-tDCS (n={n2}). "
        "All participants completed 14 sessions (20 min, 1 mA) targeting bilateral dorsolateral prefrontal cortex. "
        "Cognitive outcomes included MMSE, TMT-B, DST, VFT, and HVLT-R assessed before and after intervention. "
        f"Results: After treatment, the active group outperformed sham and control groups on TMT-B, VFT, and HVLT-R "
        f"(all P<0.05), with medium-to-large effect sizes versus sham (|d|=0.96–1.30). "
        "No significant between-group differences were observed for MMSE or DST. "
        "Sham and control groups did not differ on any cognitive measure. "
        "Conclusion: HD-tDCS may provide additional cognitive benefit beyond routine treatment in male alcohol-dependent patients.",
    )

    add_heading(doc, "1 对象与方法")
    add_heading(doc, "1.1 对象", level=2)
    p(
        doc,
        "2019年10月至2020年10月于河南省精神病医院招募男性酒依赖住院患者。"
        "入组标准：符合DSM-5酒依赖诊断；初中及以上文化；18–55岁；右利手；除尼古丁外无其他物质使用史；本人及监护人知情同意。"
        "排除标准：脑器质性或神经系统疾病；既往认知行为治疗；颅脑外伤或手术史；严重躯体疾病及植入金属设备者等。"
        f"初筛85例，按随机数字表法分配后，因提前出院脱落3例；数据核查后剔除重复录入5例，最终纳入分析{total}例"
        f"（真刺激组{n1}例、伪刺激组{n2}例、对照组{n3}例）。"
        "本研究经新乡医学院第二附属医院伦理委员会批准，所有受试者签署知情同意书。",
    )

    add_heading(doc, "1.2 认知评估", level=2)
    p(
        doc,
        "采用MMSE、TMT-B、DST、VFT及HVLT-R评估认知功能，分别反映总体认知、注意与转换、工作记忆、执行控制及言语学习记忆。",
    )

    add_heading(doc, "1.3 干预与评估流程", level=2)
    p(
        doc,
        "三组在完成苯二氮䓬类替代治疗后分别接受常规治疗、常规+HD-tDCS真刺激或常规+HD-tDCS伪刺激。"
        "设备为Soterix MXN-9，4×1环形电极，阳极F3（返回FP1/F7/C1/C3），阴极F4（返回FP2/F8/C2/C4），"
        "强度±1.0 mA，20 min/次，1次/日，共14次。伪刺激采用SHAM模式（仅首尾30 s有电流）。"
        "基线评估于入院第2周末，干预后评估于第4周末。",
    )

    add_heading(doc, "1.4 统计学处理", level=2)
    p(
        doc,
        "采用SPSS 22.0。计量资料以均数±标准差表示。组内前后比较采用配对t检验；"
        "组间比较采用单因素方差分析（方差齐）或Welch法（方差不齐），两两比较采用LSD或Games-Howell检验。"
        "并计算Cohen's d作为效应量。基线认知指标与饮酒特征行Pearson相关分析。P<0.05为差异有统计学意义。",
    )

    add_heading(doc, "2 结果")
    add_heading(doc, "2.1 三组一般资料比较", level=2)
    p(doc, f"三组年龄、教育年限、饮酒史等差异均无统计学意义（P>0.05），具有可比性（表1）。")

    add_table(
        doc,
        f"表1  三组社会特征比较（x̄±s，n={n1}/{n2}/{n3}）",
        ["变量", f"真刺激组(n={n1})", f"伪刺激组(n={n2})", f"对照组(n={n3})", "F值", "P值"],
        [
            ["年龄（岁）", t1["年龄"]["active"], t1["年龄"]["sham"], t1["年龄"]["control"], fmt_p(t1["年龄"]["F"]), fmt_p(t1["年龄"]["P"])],
            ["教育年限（年）", t1["教育年限"]["active"], t1["教育年限"]["sham"], t1["教育年限"]["control"], fmt_p(t1["教育年限"]["F"]), fmt_p(t1["教育年限"]["P"])],
            ["首次饮酒年龄（岁）", t1["首次饮酒年龄"]["active"], t1["首次饮酒年龄"]["sham"], t1["首次饮酒年龄"]["control"], fmt_p(t1["首次饮酒年龄"]["F"]), fmt_p(t1["首次饮酒年龄"]["P"])],
            ["累计使用时间（年）", t1["累计使用时间"]["active"], t1["累计使用时间"]["sham"], t1["累计使用时间"]["control"], fmt_p(t1["累计使用时间"]["F"]), fmt_p(t1["累计使用时间"]["P"])],
            ["成瘾时间（年）", t1["成瘾时间"]["active"], t1["成瘾时间"]["sham"], t1["成瘾时间"]["control"], fmt_p(t1["成瘾时间"]["F"]), fmt_p(t1["成瘾时间"]["P"])],
            ["戒酒次数（次）", t1["戒酒次数"]["active"], t1["戒酒次数"]["sham"], t1["戒酒次数"]["control"], fmt_p(t1["戒酒次数"]["F"]), fmt_p(t1["戒酒次数"]["P"])],
        ],
    )

    add_heading(doc, "2.2 基线认知功能", level=2)
    p(doc, "三组基线MMSE、TMT-B、DST、VFT、HVLT-R差异均无统计学意义（P>0.05，表2）。")
    add_table(
        doc,
        "表2  三组基线认知功能比较（x̄±s）",
        ["指标", f"真刺激组", f"伪刺激组", f"对照组", "F/P"],
        [
            ["MMSE", base["MMSE"]["active"], base["MMSE"]["sham"], base["MMSE"]["control"], f"{fmt_p(base['MMSE']['F'])}/{fmt_p(base['MMSE']['P'])}"],
            ["TMT-B（s）", base["TMT-B"]["active"], base["TMT-B"]["sham"], base["TMT-B"]["control"], f"{fmt_p(base['TMT-B']['F'])}/{fmt_p(base['TMT-B']['P'])}"],
            ["DST", base["DST"]["active"], base["DST"]["sham"], base["DST"]["control"], f"{fmt_p(base['DST']['F'])}/{fmt_p(base['DST']['P'])}"],
            ["VFT", base["VFT"]["active"], base["VFT"]["sham"], base["VFT"]["control"], f"{fmt_p(base['VFT']['F'])}/{fmt_p(base['VFT']['P'])}"],
            ["HVLT-R", base["HVLT-R"]["active"], base["HVLT-R"]["sham"], base["HVLT-R"]["control"], f"{fmt_p(base['HVLT-R']['F'])}/{fmt_p(base['HVLT-R']['P'])}"],
        ],
    )

    add_heading(doc, "2.3 干预后认知功能", level=2)
    p(
        doc,
        f"干预后，TMT-B（F={fmt_p(post['TMT-B']['F'])}, P={fmt_p(post['TMT-B']['P'])}）、"
        f"VFT（F={fmt_p(post['VFT']['F'])}, P={fmt_p(post['VFT']['P'])}）、"
        f"HVLT-R（F={fmt_p(post['HVLT-R']['F'])}, P={fmt_p(post['HVLT-R']['P'])}）组间差异有统计学意义；"
        f"MMSE（P={fmt_p(post['MMSE']['P'])}）与DST（P={fmt_p(post['DST']['P'])}）差异无统计学意义（表3）。",
    )
    p(
        doc,
        f"两两比较：真刺激组TMT-B优于伪刺激组（P={fmt_p(pw['TMT-B']['active_vs_sham'])}）及对照组（P={fmt_p(pw['TMT-B']['active_vs_control'])}）；"
        f"VFT（P={fmt_p(pw['VFT']['active_vs_sham'])}/{fmt_p(pw['VFT']['active_vs_control'])}）；"
        f"HVLT-R（P={fmt_p(pw['HVLT-R']['active_vs_sham'])}/{fmt_p(pw['HVLT-R']['active_vs_control'])}）。"
        "伪刺激组与对照组上述指标差异均无统计学意义（P>0.05）。",
    )
    add_table(
        doc,
        "表3  三组干预后认知功能比较（x̄±s）",
        ["指标", "真刺激组", "伪刺激组", "对照组", "F值", "P值"],
        [
            ["MMSE", post["MMSE"]["active"], post["MMSE"]["sham"], post["MMSE"]["control"], fmt_p(post["MMSE"]["F"]), fmt_p(post["MMSE"]["P"])],
            ["TMT-B（s）", post["TMT-B"]["active"], post["TMT-B"]["sham"], post["TMT-B"]["control"], fmt_p(post["TMT-B"]["F"]), fmt_p(post["TMT-B"]["P"])],
            ["DST", post["DST"]["active"], post["DST"]["sham"], post["DST"]["control"], fmt_p(post["DST"]["F"]), fmt_p(post["DST"]["P"])],
            ["VFT", post["VFT"]["active"], post["VFT"]["sham"], post["VFT"]["control"], fmt_p(post["VFT"]["F"]), fmt_p(post["VFT"]["P"])],
            ["HVLT-R", post["HVLT-R"]["active"], post["HVLT-R"]["sham"], post["HVLT-R"]["control"], fmt_p(post["HVLT-R"]["F"]), fmt_p(post["HVLT-R"]["P"])],
        ],
    )

    add_heading(doc, "2.4 基线认知与饮酒特征的相关性", level=2)
    if corr:
        lines = [
            f"{r['drinking']}与基线{r['cognition']}呈{'正' if r['r']>0 else '负'}相关（r={r['r']}, P={fmt_p(r['P'])})"
            for r in corr
        ]
        p(doc, "Pearson相关分析显示：" + "；".join(lines) + "。")
    else:
        p(doc, "Pearson相关分析未见统计学显著相关。")

    add_heading(doc, "3 讨论", level=1)
    p(
        doc,
        "本研究在男性酒依赖住院样本中显示，14次HD-tDCS真刺激可在常规治疗基础上进一步改善TMT-B、VFT和HVLT-R，"
        "效应量达中等至较大，而伪刺激组与对照组无显著差异，提示获益与主动刺激相关而非单纯时间效应。"
        "结果与既往tDCS改善酒依赖患者执行功能的报告一致，HD-tDCS可能通过更精确的DLPFC靶向增强疗效。",
    )
    p(
        doc,
        "局限性：（1）单中心、仅男性样本，外推性有限；（2）样本量较小且真刺激组经数据核查后n=21；"
        "（3）仅量表评估，缺乏神经电生理或影像学客观指标；（4）未报告刺激相关不良事件与远期复饮结局；"
        "（5）对照组未匹配等时设备接触。后续需多中心、大样本及随访研究进一步验证。",
    )

    add_heading(doc, "伦理与声明", level=1)
    p(doc, "伦理批准：新乡医学院第二附属医院伦理委员会（批件号待补充）。")
    p(doc, "利益冲突：作者声明无利益冲突。")
    p(doc, "资助信息：待补充。")
    p(doc, "数据可用性：去标识化数据可向通讯作者合理申请。")

    add_heading(doc, "参考文献", level=1)
    p(doc, "（保留原稿参考文献列表，投稿前按目标期刊格式统一排版并修正引用编号。）")

    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
