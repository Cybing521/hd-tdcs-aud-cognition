#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build full Chinese revised manuscript and English SCI manuscript."""

from __future__ import annotations

import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from manuscript_content import (
    DISCUSSION_EN,
    DISCUSSION_ZH,
    INTRO_EN,
    INTRO_ZH,
    assign_reference_numbers,
    reference_list,
    render_citations,
    unassigned_keys,
)

STATS_PATH = ROOT / "04_data" / "processed" / "analysis_results.json"
ENHANCED_PATH = ROOT / "04_data" / "processed" / "enhanced_analyses.json"
SENSITIVITY_PATH = ROOT / "04_data" / "processed" / "sensitivity_analyses.json"
ORIG_TXT = ROOT / "01_论文正文" / "知tDCS对酒依赖患者认知功能的影响(1).txt"
OUT_ZH = ROOT / "01_论文正文" / "HD-tDCS酒依赖认知功能_修订稿.docx"
OUT_EN = ROOT / "01_论文正文" / "HD-tDCS_alcohol_dependence_cognition_EN.docx"
OUT_EN_BRS = ROOT / "01_论文正文" / "HD-tDCS_alcohol_dependence_cognition_BRS.docx"
FIG_DIR = ROOT / "01_论文正文" / "figures"
CONSORT_FIG = FIG_DIR / "Figure1_CONSORT_flowchart.png"
MONTAGE_FIG = FIG_DIR / "Figure2_HD_tDCS_montage.png"
OUTCOMES_FIG = FIG_DIR / "Figure3_primary_outcomes_post.png"
CHANGE_FIG = FIG_DIR / "Figure4_change_trajectories.png"

RUNNING_TITLE = "HD-tDCS for AUD cognition"
KEYWORDS_EN = (
    "alcohol dependence; transcranial direct current stimulation; cognition; "
    "dorsolateral prefrontal cortex; randomized controlled trial"
)
HIGHLIGHTS_EN = [
    "Bilateral 4x1 HD-tDCS beat sham on TMT-B, VFT, and HVLT-R in AUD",
    "Active-vs-sham effects were large (Cohen d 0.96-1.30; Hedges g similar)",
    "ANCOVA and sensitivity analyses confirmed primary-outcome group effects",
    "MMSE and digit span showed no significant between-group differences",
    "Preregistered multicenter replication with relapse follow-up is needed",
]


def check_highlights(bullets: list[str]) -> None:
    """Elsevier Highlights: max 85 characters including spaces, per bullet."""
    for bullet in bullets:
        assert len(bullet) <= 85, f"Highlight exceeds 85 chars ({len(bullet)}): {bullet}"

TITLE_ZH = "高精度经颅直流电刺激对男性酒依赖患者认知功能的影响"
TITLE_EN = (
    "Effects of High-Definition Transcranial Direct Current Stimulation on Cognitive "
    "Function in Male Inpatients with Alcohol Dependence: A Randomized Sham-Controlled Trial"
)



def load_stats() -> dict:
    return json.loads(STATS_PATH.read_text(encoding="utf-8"))


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def fmt_p(v) -> str:
    """Fixed 3-decimal P value; <0.001 below threshold (never renders '0')."""
    if v is None or (_is_num(v) and v != v):
        return "—"
    if _is_num(v):
        if v < 0.001:
            return "<0.001"
        return f"{float(v):.3f}"
    return str(v)


def fmt_fdr(v) -> str:
    return fmt_p(v)


def fmt_f(v) -> str:
    if v is None or (_is_num(v) and v != v):
        return "—"
    return f"{float(v):.2f}"


def stat_str(name: str, v) -> str:
    """'P<0.001' / 'P=0.015' — avoids malformed 'P=<0.001'."""
    p = fmt_p(v)
    return f"{name}{p}" if p.startswith("<") else f"{name}={p}"


def ci_str(d: dict) -> str:
    """Cohen's d with 95% CI, e.g. '-0.96 [-1.56, -0.36]'."""
    return f"{d['d']:.2f} [{d['ci_lo']:.2f}, {d['ci_hi']:.2f}]"


def g_str(d: dict) -> str:
    """Hedges g when present, else empty."""
    if "g" not in d:
        return ""
    return f"{d['g']:.2f}"


def eta_range(anc: dict, keys: tuple[str, ...] = ("TMT-B", "VFT", "HVLT-R")) -> str:
    vals = [round(float(anc[k]["partial_eta_sq"]), 2) for k in keys]
    return f"{min(vals):.2f}-{max(vals):.2f}"


def data_verification_en() -> list[str]:
    return [
        "During database consolidation after intervention completion, 82 completing records were on file "
        "(active 27, sham 26, control 29). Five erroneous active-arm records were identified: four cross-arm "
        "misentries whose valid copies already sat in the sham file, and one duplicate active-arm row for a "
        "participant whose retained record was reassigned to the sham arm according to the original "
        "randomization log. Removing the five erroneous active-arm rows and applying that reassignment yielded "
        "the primary complete-case analysis set of n=77 (active 21, sham 27, control 29; Figure 1). No "
        "participant contributed more than one analyzed record. Sensitivity analyses that moved or dropped the "
        "reassigned record, or that analyzed all five ambiguous participants in the active arm as recorded, "
        "are reported in Statistical analysis and Results.",
    ]


def data_verification_zh() -> str:
    return (
        "1.1.1 数据核查与最终分析集  干预完成后在档完成记录82例（真刺激27、伪刺激26、对照29）。核查发现5条真刺激组误录/"
        "重复行：其中4例在伪刺激组已有正确记录，1例保留记录按原始随机化名单由真刺激组重归伪刺激组。删除上述5条误录/"
        "重复行并完成重归后，主要完整病例分析集为77例（真刺激21、伪刺激27、对照29；图1）；无受试者贡献超过1条分析记录。"
        "将重归记录改回真刺激、剔除该记录，或将5例争议受试者全部按录入臂归入真刺激的敏感性分析见统计学处理与结果。"
    )


def set_doc_font(doc: Document, western: str = "Times New Roman", east: str = "宋体") -> None:
    """Journal-like base styles: TNR/宋体, black headings, 1-inch margins."""
    style = doc.styles["Normal"]
    style.font.name = western
    style.font.size = Pt(12)
    style.font.color.rgb = RGBColor(0, 0, 0)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    pf = style.paragraph_format
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    for level in range(1, 4):
        try:
            hs = doc.styles[f"Heading {level}"]
        except KeyError:
            continue
        hs.font.name = western
        hs.font.color.rgb = RGBColor(0, 0, 0)
        hs.font.bold = True
        hs._element.rPr.rFonts.set(qn("w:eastAsia"), east)
        if level == 1:
            hs.font.size = Pt(14)
            hs.paragraph_format.space_before = Pt(18)
            hs.paragraph_format.space_after = Pt(8)
        elif level == 2:
            hs.font.size = Pt(12)
            hs.paragraph_format.space_before = Pt(12)
            hs.paragraph_format.space_after = Pt(6)
        else:
            hs.font.size = Pt(12)
            hs.paragraph_format.space_before = Pt(8)
            hs.paragraph_format.space_after = Pt(4)


def set_double_spacing(paragraph) -> None:
    paragraph.paragraph_format.line_spacing = 2.0
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)


def _set_run_font(run, *, western: str = "Times New Roman", east: str = "宋体", size_pt: float = 12) -> None:
    run.font.name = western
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)


def brs_cite(text: str) -> str:
    """Convert [1,2] / [18-24] to Brain Stimulation Vancouver in-text style (1, 2)."""

    def convert_inner(inner: str) -> str:
        tokens = [part.strip() for part in inner.split(",")]
        if len(tokens) == 1 and re.fullmatch(r"\d+-\d+", tokens[0]):
            start, end = tokens[0].split("-", 1)
            return f"({start}-{end})"
        numbers: list[str] = []
        for token in tokens:
            if re.fullmatch(r"\d+-\d+", token):
                start, end = map(int, token.split("-", 1))
                numbers.extend(str(i) for i in range(start, end + 1))
            else:
                numbers.append(token)
        return "(" + ", ".join(numbers) + ")"

    return re.sub(r"\[([^\]]+)\]", lambda match: convert_inner(match.group(1)), text)


def add_p(
    doc: Document,
    text: str,
    *,
    double_spaced: bool = False,
    body: bool = False,
    east: str = "宋体",
) -> None:
    """body=True → justified + first-line indent (paper-like paragraphs)."""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    _set_run_font(run, east=east, size_pt=12)
    if body:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # ~2 Chinese characters / 0.5 inch English first-line indent
        paragraph.paragraph_format.first_line_indent = Pt(24)
    if double_spaced:
        set_double_spacing(paragraph)


def add_labeled_paragraph(doc: Document, label: str, text: str, *, double_spaced: bool = False) -> None:
    paragraph = doc.add_paragraph()
    label_run = paragraph.add_run(f"{label}: ")
    label_run.bold = True
    _set_run_font(label_run, size_pt=12)
    body_run = paragraph.add_run(text)
    _set_run_font(body_run, size_pt=12)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if double_spaced:
        set_double_spacing(paragraph)


def add_title(doc: Document, text: str, east: str = "黑体") -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = east
    run.font.color.rgb = RGBColor(0, 0, 0)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    p.paragraph_format.space_after = Pt(12)


def add_heading(doc: Document, text: str, level: int = 1, *, double_spaced: bool = False) -> None:
    heading = doc.add_heading(text, level=level)
    # Force black TNR after Word theme styling
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体" if level == 1 else "宋体")
        run.bold = True
    if double_spaced:
        set_double_spacing(heading)


def _set_element_border(element, edges: dict[str, dict[str, str]]) -> None:
    """Replace w:tcBorders / w:tblBorders children for named edges."""
    tag = element.tag.split("}")[-1]
    borders_tag = "tcBorders" if tag == "tcPr" else "tblBorders"
    for child in list(element):
        if child.tag.split("}")[-1] == borders_tag:
            element.remove(child)
    borders = OxmlElement(f"w:{borders_tag}")
    for edge, attrs in edges.items():
        el = OxmlElement(f"w:{edge}")
        for key, val in attrs.items():
            el.set(qn(f"w:{key}"), val)
        borders.append(el)
    element.append(borders)


def apply_three_line_table(table) -> None:
    """Academic three-line table: top / header separator / bottom; no vertical rules."""
    thick = {"val": "single", "sz": "12", "space": "0", "color": "000000"}  # 1.5 pt
    mid = {"val": "single", "sz": "6", "space": "0", "color": "000000"}  # 0.75 pt
    none = {"val": "nil", "sz": "0", "space": "0", "color": "000000"}

    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    _set_element_border(
        tblPr,
        {
            "top": thick,
            "left": none,
            "bottom": thick,
            "right": none,
            "insideH": none,
            "insideV": none,
        },
    )

    # Header row bottom separator
    for cell in table.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        _set_element_border(
            tcPr,
            {
                "top": thick,
                "left": none,
                "bottom": mid,
                "right": none,
            },
        )

    # Body cells: no internal lines; last row keeps table bottom via tblBorders
    for row in table.rows[1:]:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            _set_element_border(
                tcPr,
                {
                    "top": none,
                    "left": none,
                    "bottom": none,
                    "right": none,
                },
            )


def _style_table_cell(cell, text: str, *, bold: bool = False, center: bool = True) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(str(text))
    run.bold = bold
    _set_run_font(run, size_pt=10)


def add_table(doc: Document, title: str, headers: list[str], rows: list[list[str]]) -> None:
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(10)
    cap.paragraph_format.space_after = Pt(6)
    run = cap.add_run(title)
    run.bold = True
    _set_run_font(run, size_pt=10.5)

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.autofit = True
    for j, h in enumerate(headers):
        _style_table_cell(table.rows[0].cells[j], h, bold=True, center=True)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            # First column usually a label → left align
            _style_table_cell(table.rows[i].cells[j], val, bold=False, center=(j > 0))
    apply_three_line_table(table)
    # Thin spacer after table
    spacer = doc.add_paragraph("")
    spacer.paragraph_format.space_after = Pt(6)

def abstract_zh(data: dict) -> str:
    n = data["n"]
    pw = data["pairwise_post"]
    es = data["effect_sizes"]
    total = data["total_n"]
    return (
        f"目的：探讨HD-tDCS对男性酒依赖患者认知功能的影响。"
        f"方法：85例男性酒依赖住院患者随机分为对照组（n=29）、真刺激组（n=28）和伪刺激组（n=28）；"
        f"数据核查后主要分析纳入{total}例（对照{n['3']}、真刺激{n['1']}、伪刺激{n['2']}）。"
        "三组分别接受常规治疗、常规联合14次HD-tDCS真刺激或伪刺激，干预前后评估MMSE、TMT-B、DST、VFT和HVLT-R。"
        f"结果：干预后真刺激组TMT-B、VFT、HVLT-R优于伪刺激组及对照组；"
        f"真刺激组与伪刺激组比较：TMT-B（{stat_str('P', pw['TMT-B']['active_vs_sham'])}，d={es['TMT-B']['active_vs_sham_d']}），"
        f"VFT（{stat_str('P', pw['VFT']['active_vs_sham'])}，d={es['VFT']['active_vs_sham_d']}），"
        f"HVLT-R（{stat_str('P', pw['HVLT-R']['active_vs_sham'])}，d={es['HVLT-R']['active_vs_sham_d']}）。"
        "MMSE与DST三组差异无统计学意义；伪刺激组与对照组各指标差异均无统计学意义。"
        "结论：相对伪刺激及常规治疗，双侧4×1 HD-tDCS可进一步改善男性酒依赖患者的认知转换、执行控制及言语学习记忆功能。"
    )


def abstract_en(data: dict) -> str:
    n = data["n"]
    es_ci = data["enhanced"]["effect_sizes_post_ci"]
    total = data["total_n"]
    return (
        f"Background: Cognitive impairment in alcohol dependence may increase relapse risk. "
        f"Methods: In this single-blind randomized trial, 85 male inpatients were randomized to routine care "
        f"(n=29), active HD-tDCS (n=28), or sham HD-tDCS (n=28); {total} were analyzed after data verification "
        f"(routine care n={n['3']}, active n={n['1']}, sham n={n['2']}). "
        "Active and sham arms completed 14 bilateral dorsolateral prefrontal cortex sessions (20 min, ±1.0 mA). "
        "Cognitive outcomes (MMSE, TMT-B, DST, VFT, and HVLT-R) were assessed at baseline and after intervention. "
        "Results: Active HD-tDCS surpassed sham and control on TMT-B, VFT, and HVLT-R "
        "(omnibus FDR q<0.05; versus sham, Cohen's d: "
        f"{ci_str(es_ci['TMT-B']['active_vs_sham'])}, {ci_str(es_ci['VFT']['active_vs_sham'])}, and "
        f"{ci_str(es_ci['HVLT-R']['active_vs_sham'])}; Hedges g="
        f"{g_str(es_ci['TMT-B']['active_vs_sham'])}/{g_str(es_ci['VFT']['active_vs_sham'])}/"
        f"{g_str(es_ci['HVLT-R']['active_vs_sham'])}). "
        "MMSE and DST did not differ across groups; sham and control groups were comparable. "
        "Conclusions: Relative to sham and routine care, bilateral 4×1 HD-tDCS was associated with greater "
        "improvement on domain-specific cognitive outcomes in male inpatients with alcohol dependence."
    )


def structured_abstract_en(data: dict) -> dict[str, str]:
    n = data["n"]
    enh = data["enhanced"]
    es_ci = enh["effect_sizes_post_ci"]
    anc = enh["ancova"]
    total = data["total_n"]
    return {
        "Background": (
            "Alcohol dependence features cognitive deficits that may increase relapse risk. "
            "High-definition tDCS (HD-tDCS) can focally engage the dorsolateral prefrontal cortex "
            "(DLPFC) via a 4×1 ring montage, but sham-controlled cognitive batteries in alcohol "
            "dependence remain scarce relative to craving and relapse endpoints."
        ),
        "Objective/Hypothesis": (
            "We tested whether 14 daily sessions of bilateral 4×1 HD-tDCS over the DLPFC improve "
            "cognition relative to sham HD-tDCS and routine care in male inpatients with alcohol "
            "dependence."
        ),
        "Methods": (
            f"In this single-blind three-arm trial, 85 male inpatients were randomized to routine care (n=29), "
            f"active HD-tDCS (n=28), or sham HD-tDCS (n=28); {total} were analyzed "
            f"(n={n['3']}/{n['1']}/{n['2']}) after data verification. Bilateral 4×1 sessions (±1.0 mA, "
            "20 min) targeted F3/F4 after detoxification. Primary outcomes were TMT-B, VFT, and HVLT-R. "
            "Analyses included post-intervention tests, baseline-adjusted ANCOVA, change scores, FDR "
            "correction, Cohen's d (95% CI), and Hedges g."
        ),
        "Results": (
            "Active HD-tDCS surpassed sham and control on TMT-B, VFT, and HVLT-R (omnibus FDR q<0.05; "
            f"versus sham d={ci_str(es_ci['TMT-B']['active_vs_sham'])}, "
            f"{ci_str(es_ci['VFT']['active_vs_sham'])}, and "
            f"{ci_str(es_ci['HVLT-R']['active_vs_sham'])}; "
            f"g={g_str(es_ci['TMT-B']['active_vs_sham'])}/"
            f"{g_str(es_ci['VFT']['active_vs_sham'])}/"
            f"{g_str(es_ci['HVLT-R']['active_vs_sham'])}). "
            f"ANCOVA confirmed group effects (partial η²={eta_range(anc)}), as did change scores and "
            "sensitivity analyses (ANCOVA P≤0.0003). MMSE and DST did not differ; sham equaled control."
        ),
        "Conclusion(s)": (
            "Bilateral HD-tDCS produced greater domain-specific cognitive improvement than sham, with "
            "parallel advantage over routine care. Preregistered multicenter confirmation is needed."
        ),
    }


def abstract_word_count(sections: dict[str, str]) -> int:
    return len(" ".join(sections.values()).split())


def add_title_page_brs(doc: Document) -> None:
    add_title(doc, TITLE_EN, east="Times New Roman")
    add_p(
        doc,
        "Author names and degrees: see Author Completion Checklist (end of manuscript).",
        double_spaced=True,
    )
    add_p(
        doc,
        "Department of Psychiatry, The Second Affiliated Hospital of Xinxiang Medical University "
        "(Henan Psychiatric Hospital), Xinxiang, Henan, China",
        double_spaced=True,
    )
    add_p(
        doc,
        "Corresponding author: see Author Completion Checklist (end of manuscript).",
        double_spaced=True,
    )
    add_p(doc, f"Running title: {RUNNING_TITLE}", double_spaced=True)
    add_p(doc, f"Keywords: {KEYWORDS_EN}", double_spaced=True)
    add_p(doc, "Prior abstract presentation: None reported.", double_spaced=True)
    add_p(
        doc,
        "Clinical trial registration: Not prospectively registered before enrollment (2019). "
        "Retrospective registration with disclosure of the registration date is planned before final upload; "
        "see Author Completion Checklist.",
        double_spaced=True,
    )
    doc.add_page_break()


def compute_post_hoc_power(data: dict) -> dict:
    from statsmodels.stats.power import TTestIndPower

    n_active = data["n"]["1"]
    n_sham = data["n"]["2"]
    ratio = n_sham / n_active
    analysis = TTestIndPower()
    es = data["effect_sizes"]
    mapping = {
        "TMT-B": abs(es["TMT-B"]["active_vs_sham_d"]),
        "VFT": abs(es["VFT"]["active_vs_sham_d"]),
        "HVLT-R": abs(es["HVLT-R"]["active_vs_sham_d"]),
    }
    out: dict[str, dict] = {}
    for outcome, d in mapping.items():
        out[outcome] = {
            "cohens_d": round(d, 2),
            "n_active": n_active,
            "n_sham": n_sham,
            "alpha": 0.05,
            "power": round(
                analysis.power(effect_size=d, nobs1=n_active, alpha=0.05, ratio=ratio, alternative="two-sided"),
                3,
            ),
        }
    path = ROOT / "04_data" / "processed" / "post_hoc_power.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return out


def methods_supplement_en(data: dict, power: dict) -> list[tuple[str, list[str]]]:
    mde = data.get("sensitivity", {}).get("minimum_detectable_effect", {}).get(
        "min_detectable_d_80power_21v27", 0.83
    )
    return [
        (
            "Outcome hierarchy, sample size, and multiple testing",
            [
                "TMT-B, VFT, and HVLT-R were prespecified as primary cognitive outcomes based on their relevance "
                "to executive control, set shifting, and verbal learning and memory in alcohol dependence. MMSE and "
                "DST were designated secondary outcomes reflecting global screening and working memory, respectively. "
                "Associations between drinking characteristics and baseline cognition were declared exploratory.",
                "No formal a priori sample-size calculation was documented in the original protocol. For a two-sided "
                f"α=0.05 comparison of active n={data['n']['1']} versus sham n={data['n']['2']}, the minimum "
                f"detectable Cohen's d at 80% power is {mde:.2f}. Observed active-versus-sham effects on the primary "
                "outcomes exceeded this threshold. Observed post hoc power estimates are reported only as descriptive "
                f"complements (TMT-B {power['TMT-B']['power']:.1%}, VFT {power['VFT']['power']:.1%}, "
                f"HVLT-R {power['HVLT-R']['power']:.1%}) and were not used for inferential claims.",
                "Benjamini-Hochberg false discovery rate (FDR) correction was applied across all five "
                "post-intervention omnibus between-group tests, and primary outcomes remained significant at "
                "FDR-adjusted q<0.05. The same correction was applied across the twenty exploratory correlation tests. "
                "Change-score and ANCOVA families were treated as confirmatory parallel analyses and are reported "
                "with nominal P values. No interim analyses were planned or conducted; the trial ran to completion.",
            ],
        ),
        (
            "Safety monitoring",
            [
                "Systematic adverse-event ascertainment using structured tDCS safety checklists was not prespecified "
                "in the original study protocol [@brunoni2011]. Session tolerability was monitored through routine inpatient "
                "clinical observation during each 20-min session, including inquiry about scalp tingling, skin "
                "redness, headache, and discomfort. Serious adverse events would have been recorded in medical records "
                "and reported to the ethics committee according to institutional policy.",
            ],
        ),
    ]


def methods_supplement_zh(data: dict, power: dict) -> list[str]:
    mde = data.get("sensitivity", {}).get("minimum_detectable_effect", {}).get(
        "min_detectable_d_80power_21v27", 0.83
    )
    return [
        "1.5 主要结局与样本量  预先设定的主要结局为干预后TMT-B、VFT和HVLT-R组间比较，次要结局为MMSE和DST；"
        "基线认知与饮酒特征的相关性为探索性分析。原方案未记录正式样本量估算。"
        f"在α=0.05（双侧）、真刺激组n={data['n']['1']}与伪刺激组n={data['n']['2']}时，80%效能下最小可检测Cohen's d约为"
        f"{mde:.2f}；主要结局观察到的真–伪刺激效应超过该阈值。事后效能仅作描述性补充"
        f"（TMT-B {power['TMT-B']['power']:.1%}，VFT {power['VFT']['power']:.1%}，"
        f"HVLT-R {power['HVLT-R']['power']:.1%}），不作为推断依据。未计划亦未实施期中分析。",
        "1.6 多重比较  对5项干预后组间比较采用Benjamini-Hochberg FDR校正；三项主要结局校正后仍保持显著（q<0.05）。"
        "变化值与ANCOVA作为平行确证分析，报告名义P值。",
        "1.7 安全性监测  原研究方案未预设结构化tDCS不良事件量表[@brunoni2011]。每次20 min刺激期间通过住院常规临床观察记录耐受情况，包括头皮刺痛、皮肤发红、头痛及不适主诉；"
        "若发生严重不良事件，将按机构要求写入病历并报告伦理委员会。",
    ]


def add_safety_results(doc: Document, data: dict, lang: str, nummap: dict, *, double_spaced: bool = False) -> None:
    n = data["n"]
    if lang == "zh":
        add_heading(doc, "2.8 安全性与耐受性", level=2)
        add_p(
            doc,
            render_citations(
                "本研究未采用结构化不良事件量表进行系统采集[@brunoni2011]。干预期间未见需中止试验的严重刺激相关事件报告；"
                "临床人员会在每次刺激后询问头皮刺痛、皮肤发红、头痛及不适，但未形成结构化数据库。"
                "真刺激组1例、伪刺激组2例因提前出院未完成全部14次干预（见图1）。",
                nummap,
            ),
        )
        add_table(
            doc,
            "表5  安全性与随访完成情况",
            ["项目", f"真刺激组(n={n['1']})", f"伪刺激组(n={n['2']})", f"对照组(n={n['3']})"],
            [
                ["干预期间提前出院（研究层面，见图1）", "1", "2", "0"],
                ["结构化不良事件采集", "否", "否", "否"],
                ["非结构化耐受性询问", "是", "是", "不适用"],
                ["常规临床观察", "是", "是", "是"],
            ],
        )
    else:
        add_heading(doc, "Safety and tolerability", level=2, double_spaced=double_spaced)
        add_p(
            doc,
            brs_cite(
                render_citations(
                    "Structured adverse-event collection was not prespecified [@brunoni2011]. No serious "
                    "stimulation-related events requiring trial discontinuation were reported during routine inpatient "
                    "monitoring. Clinicians asked about scalp tingling, skin redness, headache, and discomfort after "
                    "sessions, but responses were not recorded in a structured database. One active and two sham "
                    "participants discontinued before completing all 14 sessions (Figure 1).",
                    nummap,
                )
            ),
            double_spaced=double_spaced,
        )
        add_table(
            doc,
            "Table 5. Safety monitoring and session completion",
            ["Item", f"Active (n={n['1']})", f"Sham (n={n['2']})", f"Control (n={n['3']})"],
            [
                ["Early discontinuation before 14 sessions (study level; Figure 1)", "1", "2", "0"],
                ["Structured adverse-event collection", "No", "No", "No"],
                ["Informal session tolerability monitoring", "Yes", "Yes", "N/A"],
                ["Routine clinical monitoring", "Yes", "Yes", "Yes"],
            ],
        )


def add_structured_abstract(doc: Document, sections: dict[str, str], *, double_spaced: bool = False) -> None:
    add_heading(doc, "Abstract", level=1, double_spaced=double_spaced)
    for label, text in sections.items():
        add_labeled_paragraph(doc, label, text, double_spaced=double_spaced)


def add_highlights(doc: Document, bullets: list[str], *, double_spaced: bool = False) -> None:
    add_heading(doc, "Highlights", level=1, double_spaced=double_spaced)
    for bullet in bullets:
        add_p(doc, f"• {bullet}", double_spaced=double_spaced)


def add_english_methods(doc: Document, data: dict, power: dict, nummap: dict, *, double_spaced: bool = False) -> None:
    add_heading(
        doc,
        "Methods and Materials" if double_spaced else "Methods",
        level=1,
        double_spaced=double_spaced,
    )
    sections = methods_en_brs(data) + methods_supplement_en(data, power)
    for subsection, paragraphs in sections:
        add_heading(doc, subsection, level=2, double_spaced=double_spaced)
        for para in paragraphs:
            add_p(doc, brs_cite(render_citations(para, nummap)), double_spaced=double_spaced, body=True)


def add_english_back_matter(doc: Document, data: dict, *, double_spaced: bool = False) -> None:
    add_heading(doc, "Acknowledgments", level=1, double_spaced=double_spaced)
    add_p(doc, "The authors thank the clinical staff and participants at Henan Psychiatric Hospital.", double_spaced=double_spaced)
    add_heading(doc, "Financial Disclosures", level=1, double_spaced=double_spaced)
    add_p(
        doc,
        "The authors report no biomedical financial interests or potential conflicts of interest.",
        double_spaced=double_spaced,
    )
    add_p(doc, "Funding: see Author Completion Checklist.", double_spaced=double_spaced)
    add_p(doc, "Author contributions: see Author Completion Checklist.", double_spaced=double_spaced)
    add_p(
        doc,
        "Data availability: De-identified individual-level data are not publicly available because of "
        "institutional privacy restrictions and the absence of participant consent for open sharing. "
        "Summary statistics and structured analysis outputs sufficient to reproduce the reported inferential "
        "results are provided in the Supplementary Material. De-identified data may be available from the "
        "corresponding author upon reasonable request and with appropriate institutional agreements.",
        double_spaced=double_spaced,
    )
    add_p(
        doc,
        "Code availability: Statistical analysis and manuscript generation scripts are available from the "
        "corresponding author upon request. A permanent public repository URL has not yet been assigned.",
        double_spaced=double_spaced,
    )
    add_p(
        doc,
        "Trial registration: This study was not prospectively registered before participant enrollment "
        "(2019). Retrospective registration with disclosure of the registration date is planned before "
        "journal submission.",
        double_spaced=double_spaced,
    )
    add_heading(doc, "Declaration of Generative AI and AI-Assisted Technologies", level=1, double_spaced=double_spaced)
    add_p(
        doc,
        "During preparation of this work, the authors used generative AI tools for literature monitoring, "
        "language editing, and manuscript formatting. After using these tools, the authors reviewed and edited the "
        "content as needed and take full responsibility for the published article.",
        double_spaced=double_spaced,
    )


def add_table_legends(doc: Document, data: dict, *, double_spaced: bool = False) -> None:
    n = data["n"]
    add_heading(doc, "Figure and Table Legends", level=1, double_spaced=double_spaced)
    legends = [
        "Figure 1. CONSORT participant flow diagram for the three-arm randomized trial "
        "(85 randomized; 77 analyzed: control n=29; active HD-tDCS n=21; sham HD-tDCS n=27).",
        "Figure 2. Bilateral 4×1 HD-tDCS montage (F3 anode / F4 cathode with listed return electrodes).",
        "Figure 3. Post-intervention primary outcomes (TMT-B, VFT, HVLT-R) with individual observations.",
        "Figure 4. Baseline-to-post trajectories for primary outcomes (mean±95% CI).",
        f"Table 1. Sociodemographic and drinking characteristics (mean±SD) in active (n={n['1']}), "
        f"sham (n={n['2']}), and control (n={n['3']}) groups.",
        "Table 2. Cognitive scores at baseline and post-intervention (mean±SD) across groups.",
        "Table 3. Baseline-to-post-intervention change scores (mean change±SD) across groups.",
        "Table 4. Baseline-adjusted (ANCOVA) post-intervention comparisons: adjusted group means "
        "with F, P, and partial eta-squared for each outcome.",
        "Table 5. Safety monitoring and session completion.",
        "Supplementary Table S1. Analysis-set row reconciliation (cleaning and reassignment).",
    ]
    for legend in legends:
        add_p(doc, legend, double_spaced=double_spaced)



def assemble_english_manuscript(doc: Document, data: dict, power: dict, *, double_spaced: bool, brs_title_page: bool) -> None:
    nummap = build_nummap(data, power)
    if brs_title_page:
        add_title_page_brs(doc)
    else:
        add_title(doc, TITLE_EN, east="Times New Roman")

    abstract_sections = structured_abstract_en(data)
    wc = abstract_word_count(abstract_sections)
    if wc > 250:
        print(f"Warning: structured abstract is {wc} words (limit 250).")
    add_structured_abstract(doc, abstract_sections, double_spaced=double_spaced)
    add_highlights(doc, HIGHLIGHTS_EN, double_spaced=double_spaced)

    add_heading(doc, "Introduction", level=1, double_spaced=double_spaced)
    for para in INTRO_EN:
        add_p(doc, brs_cite(render_citations(para, nummap)), double_spaced=double_spaced, body=True)

    add_english_methods(doc, data, power, nummap, double_spaced=double_spaced)
    add_consort_figure(doc, "en", double_spaced=double_spaced)
    add_result_figures(doc, "en", double_spaced=double_spaced)
    results_section(doc, data, "en", double_spaced=double_spaced)
    add_safety_results(doc, data, "en", nummap, double_spaced=double_spaced)

    add_heading(doc, "Discussion", level=1, double_spaced=double_spaced)
    for para in DISCUSSION_EN:
        add_p(doc, brs_cite(render_citations(para, nummap)), double_spaced=double_spaced, body=True)

    add_english_back_matter(doc, data, double_spaced=double_spaced)
    add_references_vancouver(doc, nummap, double_spaced=double_spaced)
    add_table_legends(doc, data, double_spaced=double_spaced)
    add_analysis_set_supplement(doc, "en", double_spaced=double_spaced)
    add_author_completion_checklist(doc, "en", double_spaced=double_spaced)


def methods_en_brs(data: dict) -> list[tuple[str, list[str]]]:
    n = data["n"]
    total = data["total_n"]
    mde = data.get("sensitivity", {}).get("minimum_detectable_effect", {}).get(
        "min_detectable_d_80power_21v27", 0.83
    )
    return [
        (
            "Study design",
            [
                "This single-blind, single-center randomized trial compared routine inpatient care, "
                "routine care plus active HD-tDCS, and routine care plus sham HD-tDCS in male inpatients "
                "with DSM-5 alcohol dependence.",
            ],
        ),
        (
            "Participants",
            [
                "Male inpatients were recruited from Henan Psychiatric Hospital (The Second Affiliated Hospital "
                "of Xinxiang Medical University) between October 2019 and October 2020. Inclusion criteria were "
                "age 18–55 years, right-handedness, junior high school education or above, and written informed "
                "consent. Exclusion criteria included major neurological disease, prior cognitive behavioral therapy, "
                "intracranial surgery, pacemaker or metallic implants, and severe medical comorbidity.",
                "Eighty-five participants were randomized 1:1:1 to routine care (n=29), active HD-tDCS (n=28), or "
                "sham HD-tDCS (n=28). Three participants discontinued the intervention early because of hospital "
                "discharge (1 active, 2 sham), leaving 82 completing records on file. After removal of five "
                f"erroneous active-arm records and reassignment of one retained record to sham, {total} participants "
                f"entered the primary complete-case analysis (active n={n['1']}, sham n={n['2']}, control n={n['3']}). "
                "The study was approved by the institutional ethics committee (approval number: see Author Completion Checklist). All "
                "participants and their guardians provided written informed consent. Participant flow is summarized "
                "in Figure 1.",
            ]
            + data_verification_en(),
        ),
        (
            "Randomization and allocation",
            [
                "The randomization sequence was generated with a random number table at a 1:1:1 allocation ratio. "
                "The mechanism used to conceal the upcoming assignment from enrollment staff was not documented in "
                "the archived protocol and is listed in the Author Completion Checklist.",
            ],
        ),
        (
            "Procedures",
            [
                "After benzodiazepine-assisted detoxification, participants received routine inpatient care alone "
                "(control), routine care plus active HD-tDCS, or routine care plus sham HD-tDCS. The control arm did "
                "not receive equivalent device-contact time; active and sham arms each completed 14 additional "
                "20-min stimulation sessions. Stimulation was delivered with a Soterix MXN-9 HD-tES device using "
                "bilateral 4×1 ring montages: F3 anode with return electrodes at FP1, F7, C1, and C3; F4 cathode "
                "with return electrodes at FP2, F8, C2, and C4 (Figure 2). Intensity was ±1.0 mA for 20 min once "
                "daily for 14 sessions [@fregni2021,@kuo2013]. Electrode size/material and session-adherence "
                "verification details are listed in the Author Completion Checklist. Active stimulation used continuous direct current "
                "throughout each session. Sham stimulation applied current only during the first and last 30 s of "
                "each session.",
            ],
        ),
        (
            "Outcome measures",
            [
                "Cognitive outcomes were assessed at baseline (week 2 of hospitalization) and post-intervention "
                "(week 4). The MMSE [@zhangmy2015] was used for global cognitive screening. TMT-B measured attention and set "
                "shifting; DST assessed working memory; VFT assessed executive control; and HVLT-R assessed verbal "
                "learning and memory.",
            ],
        ),
        (
            "Statistical analysis",
            [
                "Analyses were performed on the cleaned individual-level complete-case dataset. The original analyses "
                "used SPSS 22.0; all statistics reported here were independently recomputed from the raw records in "
                "Python (SciPy and statsmodels) and cross-checked against the original output. Data are presented as "
                "mean±SD. Within-group pre-to-post changes were evaluated with paired t-tests. Between-group "
                "comparisons at each time point used one-way ANOVA when variances were homogeneous (Levene's test) "
                "and Welch's ANOVA otherwise, with LSD or Games-Howell pairwise comparisons accordingly. To separate "
                "stimulation-specific benefit from partial recovery during inpatient care, baseline-to-post-intervention "
                "change scores were compared across groups using the same procedures, and analysis of covariance "
                "(ANCOVA) tested group differences in each post-intervention score with its baseline value as "
                "covariate; homogeneity of regression slopes was verified through group-by-baseline interaction "
                "terms. TMT-B, VFT, and HVLT-R were prespecified primary outcomes and MMSE and DST secondary "
                "outcomes. Benjamini-Hochberg false discovery rate (FDR) correction was applied across the five "
                "post-intervention omnibus tests. Change-score and ANCOVA families were reported with nominal P "
                "values as confirmatory parallel analyses. For primary contrasts, Cohen's d was reported with "
                "large-sample 95% confidence intervals together with Hedges g correction [@cohen1988]. Pearson "
                "correlation examined all twenty associations between four drinking characteristics and five "
                "baseline cognitive measures as an exploratory analysis, with Benjamini-Hochberg correction applied "
                "across the twenty tests. Sensitivity analyses reassigned the reclassified participant to the active "
                "arm, excluded that participant, or analyzed all five ambiguous participants in the active arm as "
                f"recorded; the minimum detectable d at 80% power for n=21 versus n=27 is {mde:.2f}. Two-sided "
                "P<0.05 denoted nominal significance; multiplicity-adjusted inference used q<0.05. No interim "
                "analyses were conducted.",
            ],
        ),
        (
            "Integrity of the blind",
            [
                "Blinding applied to participants in the active and sham arms through an identical electrode setup "
                "and a sham protocol that delivered brief current at the beginning and end of each session. The "
                "routine-care arm was unblinded by design because no device sessions were delivered. Outcome "
                "assessors were not blinded. Sham credibility and expectancy ratings were not collected.",
            ],
        ),
    ]


def collect_cite_texts(data: dict, power: dict) -> list[str]:
    """All EN paragraphs in document order — drives canonical Vancouver numbering."""
    texts = list(INTRO_EN)
    for _, paras in methods_en_brs(data):
        texts.extend(paras)
    for _, paras in methods_supplement_en(data, power):
        texts.extend(paras)
    texts.extend(DISCUSSION_EN)
    return texts


def build_nummap(data: dict, power: dict) -> dict[str, int]:
    nummap = assign_reference_numbers(collect_cite_texts(data, power))
    missing = sorted(unassigned_keys(nummap))
    if missing:
        raise RuntimeError(f"Reference keys defined but never cited: {missing}")
    return nummap


def add_references_vancouver(doc: Document, nummap: dict, *, double_spaced: bool = False) -> None:
    add_heading(doc, "References", level=1, double_spaced=double_spaced)
    for i, ref in enumerate(reference_list(nummap), start=1):
        add_p(doc, f"{i}. {ref}", double_spaced=double_spaced)


def add_table_legends_brs(doc: Document, data: dict) -> None:
    add_table_legends(doc, data, double_spaced=True)


def add_brs_back_matter(doc: Document) -> None:
    add_english_back_matter(doc, {}, double_spaced=True)


def add_consort_figure(doc: Document, lang: str, *, double_spaced: bool = False) -> None:
    if not CONSORT_FIG.exists():
        return
    if lang == "zh":
        add_heading(doc, "图1  CONSORT参与者流程图", level=2)
        doc.add_picture(str(CONSORT_FIG), width=Inches(6.5))
        add_p(
            doc,
            "图1  CONSORT参与者流程图。随机分配85例（对照29、真刺激28、伪刺激28）；干预期间脱落3例；"
            "删除5条真刺激组误录/重复行并将1例保留记录按随机化名单重归伪刺激组后，主要分析纳入77例"
            "（真刺激21、伪刺激27、对照29）。",
        )
    else:
        add_heading(doc, "Figure 1", level=2, double_spaced=double_spaced)
        doc.add_picture(str(CONSORT_FIG), width=Inches(6.5))
        add_p(
            doc,
            "Figure 1. CONSORT participant flow diagram. Eighty-five participants were randomized "
            "(control 29, active 28, sham 28). Three discontinued before completing the intervention. "
            "Five erroneous active-arm records were removed and one retained record was reassigned to the "
            "sham arm per the randomization log, leaving 77 participants in the primary analysis "
            "(active 21, sham 27, control 29).",
            double_spaced=double_spaced,
        )


def add_result_figures(doc: Document, lang: str, *, double_spaced: bool = False) -> None:
    figures = [
        (
            MONTAGE_FIG,
            "图2  HD-tDCS 4×1 电极布局示意图",
            "Figure 2",
            "图2  双侧4×1环形蒙太奇：F3阳极（返回电极FP1、F7、C1、C3）与F4阴极（返回电极FP2、F8、C2、C4）。",
            "Figure 2. Bilateral 4×1 HD-tDCS montage. F3 anode with returns at FP1, F7, C1, and C3; "
            "F4 cathode with returns at FP2, F8, C2, and C4.",
        ),
        (
            OUTCOMES_FIG,
            "图3  干预后主要结局",
            "Figure 3",
            "图3  干预后TMT-B、VFT与HVLT-R个体分布（箱线+散点）。TMT-B越低越好。",
            "Figure 3. Post-intervention primary outcomes (box plots with individual points) for TMT-B, "
            "VFT, and HVLT-R. Lower TMT-B scores indicate better performance.",
        ),
        (
            CHANGE_FIG,
            "图4  基线至干预后变化轨迹",
            "Figure 4",
            "图4  三项主要结局的基线→干预后均数±95%CI及个体轨迹。",
            "Figure 4. Baseline-to-post trajectories for the three primary outcomes (mean±95% CI with "
            "individual paths).",
        ),
    ]
    for path, zh_h, en_h, zh_leg, en_leg in figures:
        if not path.exists():
            continue
        if lang == "zh":
            add_heading(doc, zh_h, level=2)
            doc.add_picture(str(path), width=Inches(6.5))
            add_p(doc, zh_leg)
        else:
            add_heading(doc, en_h, level=2, double_spaced=double_spaced)
            doc.add_picture(str(path), width=Inches(6.5))
            add_p(doc, en_leg, double_spaced=double_spaced)


def methods_zh(data: dict) -> list[str]:
    n = data["n"]
    total = data["total_n"]
    mde = data.get("sensitivity", {}).get("minimum_detectable_effect", {}).get(
        "min_detectable_d_80power_21v27", 0.83
    )
    return [
        "1.1 对象  本研究于2019年10月至2020年10月在河南省精神病医院（新乡医学院第二附属医院）招募男性酒依赖住院患者。入组标准：(1)符合DSM-5酒依赖诊断标准；(2)初中及以上文化程度；(3)除尼古丁外无其他精神活性物质使用史；(4)男性，18–55岁；(5)右利手；(6)本人及监护人知情同意。排除标准：脑器质性或神经系统疾病；既往认知行为治疗；颅脑外伤或手术史；昏迷或卒中史；严重躯体疾病；心脏起搏器、金属植入物等。脱落标准：拒绝tDCS或量表评估不配合。",
        f"共随机入组85例（对照组29例、真刺激组28例、伪刺激组28例）；干预期间因提前出院脱落3例（真刺激1例、伪刺激2例）。"
        f"数据核查删除5条真刺激组误录/重复行，并将1例保留记录按随机化名单重归伪刺激组后，主要完整病例分析纳入{total}例"
        f"（真刺激组{n['1']}例、伪刺激组{n['2']}例、对照组{n['3']}例）。参与者流程见图1。本研究经新乡医学院第二附属医院伦理委员会批准，所有受试者签署知情同意书。",
        data_verification_zh(),
        "1.2.1 一般情况调查表  收集年龄、受教育年限、首次饮酒年龄、累计使用时间、成瘾时间和戒酒次数。",
        "1.2.2 认知功能评估  (1)MMSE：评估定向、记忆、注意、计算、语言等，总分30分[@zhangmy2015]。(2)TMT-B：记录完成时间，反映注意与认知转换。(3)DST：顺背与倒背分数之和，反映工作记忆。(4)VFT：1 min内正确动物名称数，反映执行控制。(5)HVLT-R：3次学习总分，反映言语学习记忆。",
        "1.3.1 干预流程  三组在完成苯二氮䓬类替代治疗后分别接受常规治疗、常规+HD-tDCS真刺激或常规+HD-tDCS伪刺激。对照组未接受等时设备接触；真/伪刺激组额外完成14次、每次20 min刺激。设备为Soterix MXN-9 HD-tES刺激器，波形DC。阳极中央电极F3（返回FP1、F7、C1、C3），阴极中央电极F4（返回FP2、F8、C2、C4），强度±1.0 mA，20 min/次，1次/日，共14次（图2）[@fregni2021,@kuo2013]。真刺激组为ACTIVE STIM；伪刺激组为SHAM STIM（仅首尾30 s有电流）。电极尺寸/材料及依从性核查细节见文末作者待填清单。",
        "1.3.2 评估流程  基线评估于入院第2周末，干预后评估于第4周末。",
        "1.3.3 随机与盲法  随机序列由随机数字表按1:1:1生成；分配隐藏具体机制见文末作者待填清单。真/伪刺激组受试者盲，常规治疗组因无设备接触而不盲；评估者未盲；未采集伪刺激可信度或期待评分。",
        "1.4 统计学处理  原始分析采用SPSS 22.0；本稿全部统计量基于清洗后的个体数据在Python（SciPy与statsmodels）中独立重算并与原输出交叉核对。计量资料以x̄±s表示。组内前后比较采用配对t检验；干预后组间比较采用单因素方差分析（方差齐，Levene检验）或Welch法（方差不齐），两两比较相应采用LSD或Games-Howell检验。进一步比较三组基线到干预后的变化值（change score），并采用协方差分析（ANCOVA）以基线得分为协变量校正干预后组间差异，回归斜率同质性经组×基线交互项检验。主要对比报告Cohen's d及其95%置信区间（同时给出Hedges g）。预先设定主要结局为TMT-B、VFT和HVLT-R，次要结局为MMSE和DST；对5项干预后组间比较采用Benjamini-Hochberg FDR校正；变化值与ANCOVA报告名义P值。"
        f"敏感性分析包括将重归记录改回真刺激、剔除该记录，或将5例争议受试者全部按录入臂归入真刺激；n=21 vs 27、80%效能下最小可检测d约为{mde:.2f}。"
        "基线认知指标与饮酒特征的全部20个组合行探索性Pearson相关分析并进行FDR校正。P<0.05为名义显著；多重比较校正后以q<0.05为准。未实施期中分析。",
    ]


def methods_en(data: dict) -> list[str]:
    n = data["n"]
    total = data["total_n"]
    return [
        "Participants: Male inpatients with DSM-5 alcohol dependence were recruited from Henan Psychiatric Hospital (The Second Affiliated Hospital of Xinxiang Medical University) between October 2019 and October 2020. Inclusion criteria were age 18–55 years, right-handedness, junior high school education or above, and informed consent. Exclusion criteria included major neurological disease, prior cognitive behavioral therapy, intracranial surgery, pacemaker or metallic implants, and severe medical comorbidity.",
        f"Of 85 screened patients, 3 dropped out before completion and 5 duplicate or misclassified records were excluded after data verification, leaving {total} analyzable participants (active n={n['1']}, sham n={n['2']}, control n={n['3']}). The study was approved by the institutional ethics committee and all participants provided written informed consent.",
        "Procedures: After benzodiazepine-assisted detoxification, participants received routine inpatient care alone (control), routine care plus active HD-tDCS, or routine care plus sham HD-tDCS. Stimulation was delivered with a Soterix MXN-9 device using bilateral 4×1 ring montages over F3 (anode) and F4 (cathode) at ±1.0 mA for 20 min once daily for 14 sessions. Sham stimulation applied current only during the first and last 30 s of each session.",
        "Outcomes: MMSE, TMT-B, DST, VFT, and HVLT-R were assessed at baseline (week 2 of hospitalization) and post-intervention (week 4).",
        "Statistics: Data are presented as mean±SD. Within-group changes were analyzed with paired t-tests. Between-group comparisons used one-way ANOVA or Welch's ANOVA with LSD or Games-Howell post hoc tests. Cohen's d was calculated for significant outcomes. Pearson correlation was used to examine associations between baseline cognition and drinking characteristics. Two-sided P<0.05 was considered significant.",
    ]



def add_author_completion_checklist(doc: Document, lang: str, *, double_spaced: bool = False) -> None:
    """Appendix-style checklist; keep manuscript body free of bracket placeholders."""
    if lang == "zh":
        add_heading(doc, "作者待填清单（投稿前完成）", level=1)
        items = [
            "作者姓名、学位、单位与通讯作者联系方式",
            "伦理批件号",
            "基金与 CRediT 作者贡献",
            "回顾性试验注册号及注册日期（2019 年入组，未前瞻注册）",
            "分配隐藏具体机制（如密封信封/中心分配等）",
            "电极直径/材料、电流密度、ramp 参数与依从性核查方式",
            "不良事件病历回顾计数（若可获得）",
        ]
        for it in items:
            add_p(doc, f"• {it}")
    else:
        add_heading(doc, "Author Completion Checklist", level=1, double_spaced=double_spaced)
        items = [
            "Author names, degrees, affiliations, and corresponding-author contact details",
            "Ethics approval number",
            "Funding statement and CRediT contributions",
            "Retrospective trial registration number and registration date (enrolled 2019; not prospectively registered)",
            "Allocation-concealment mechanism (e.g., sealed envelopes or central allocation)",
            "Electrode diameter/material, current density, ramp parameters, and adherence verification",
            "Adverse-event counts from chart review, if available",
        ]
        for it in items:
            add_p(doc, f"• {it}", double_spaced=double_spaced)


def add_analysis_set_supplement(doc: Document, lang: str, *, double_spaced: bool = False) -> None:
    """Supplementary Table S1: row-level cleaning / reassignment accounting."""
    log = json.loads((ROOT / "04_data" / "processed" / "data_cleaning_log.json").read_text(encoding="utf-8"))
    before = log["sample_size_before"]
    after = log["sample_size_after"]
    rows = [
        ["Randomized (allocation)", "29", "28", "28", "85"],
        ["Discontinued intervention (early discharge)", "0", "1", "2", "3"],
        ["Completing records on file", str(before["control"]), str(before["active"]), str(before["sham"]), str(before["total"])],
        ["Erroneous active-arm rows removed", "0", "5", "0", "5"],
        ["Retained record reassigned active→sham", "0", "−1", "+1", "0"],
        ["Primary analysis set", str(after["control"]), str(after["active"]), str(after["sham"]), str(after["total"])],
    ]
    note_rows = [[r["participant_id"], r["note"]] for r in log["rows_removed_from_active"]]
    if lang == "zh":
        add_heading(doc, "补充材料", level=1)
        add_heading(doc, "补充表S1  分析集行级核查与重归对账", level=2)
        add_p(doc, "依据原始随机化名单与清洗日志，将误录/重复行删除并将保留记录重归后形成主要分析集。")
        add_table(
            doc,
            "补充表S1a  样本量流转（对照 / 真刺激 / 伪刺激）",
            ["步骤", "对照", "真刺激", "伪刺激", "合计"],
            rows,
        )
        add_table(
            doc,
            "补充表S1b  自真刺激臂移除或重归的记录",
            ["受试者", "处理说明"],
            note_rows,
        )
    else:
        add_heading(doc, "Supplementary Material", level=1, double_spaced=double_spaced)
        add_heading(doc, "Supplementary Table S1. Analysis-set row reconciliation", level=2, double_spaced=double_spaced)
        add_p(
            doc,
            "Row-level cleaning followed the original randomization log: erroneous active-arm entries were "
            "removed, and one retained record was reassigned to the sham arm before the primary complete-case analysis.",
            double_spaced=double_spaced,
        )
        add_table(
            doc,
            "Supplementary Table S1a. Sample-size flow (control / active / sham)",
            ["Step", "Control", "Active", "Sham", "Total"],
            rows,
        )
        add_table(
            doc,
            "Supplementary Table S1b. Records removed from or reassigned out of the active arm",
            ["Participant ID (de-identified label)", "Action"],
            note_rows,
        )


def add_within_group_results(doc: Document, data: dict, lang: str, n: dict, *, double_spaced: bool = False) -> None:
    rows = data.get("within_group_change", [])
    if not rows:
        return
    if lang == "zh":
        add_heading(doc, "2.3 组内干预前后变化", level=2)
        add_p(
            doc,
            "配对t检验显示，三组在住院期间均出现一定认知改善，符合戒断早期部分自然恢复；真刺激组在TMT-B、VFT和HVLT-R上的组内改善幅度最大。"
            "基线与干预后均数见表2，组间变化值比较见表3。",
            body=True,
        )
    else:
        add_heading(doc, "Within-group pre-post changes", level=2, double_spaced=double_spaced)
        add_p(
            doc,
            "Paired t-tests indicated cognitive improvement in all three groups during inpatient care, consistent with "
            "partial recovery during early abstinence. Gains were most pronounced in the active group for TMT-B, "
            "VFT, and HVLT-R. Baseline and post-intervention means are shown in Table 2; between-group change scores "
            "are reported in Table 3.",
            double_spaced=double_spaced,
            body=True,
        )


def add_change_score_results(doc: Document, data: dict, lang: str, n: dict, *, double_spaced: bool = False) -> None:
    rows = {r["outcome"]: r for r in data.get("change_score_anova", [])}
    cpw = data.get("change_score_pairwise", {})
    sd_rows = {r["outcome"]: r for r in data.get("enhanced", {}).get("change_sd", [])}
    if not rows:
        return

    def ms(outcome: str, grp: str) -> str:
        cell = sd_rows.get(outcome, {}).get(grp)
        return f"{cell['mean']}±{cell['sd']}" if cell else "—"

    table_rows = []
    for outcome in ["MMSE", "TMT-B", "DST", "VFT", "HVLT-R"]:
        r = rows.get(outcome)
        if not r:
            continue
        table_rows.append(
            [
                outcome,
                ms(outcome, "active"),
                ms(outcome, "sham"),
                ms(outcome, "control"),
                fmt_f(r["F"]),
                fmt_p(r["P"]),
            ]
        )
    if lang == "zh":
        add_heading(doc, "2.4 变化值组间比较", level=2)
        add_p(
            doc,
            f"变化值分析显示，TMT-B（{stat_str('F', rows['TMT-B']['F'])}, {stat_str('P', rows['TMT-B']['P'])}）、"
            f"VFT（{stat_str('F', rows['VFT']['F'])}, {stat_str('P', rows['VFT']['P'])}）和HVLT-R（{stat_str('F', rows['HVLT-R']['F'])}, {stat_str('P', rows['HVLT-R']['P'])}）"
            f"三组改善幅度差异有统计学意义；真刺激组改善大于伪刺激组及对照组（均P<0.05）。MMSE与DST变化值组间差异无统计学意义（表3）。",
        )
        add_table(
            doc,
            "表3  三组认知功能变化值比较（Δ值，x̄±s）",
            ["指标", f"真刺激组(n={n['1']})", f"伪刺激组(n={n['2']})", f"对照组(n={n['3']})", "F", "P"],
            table_rows,
        )
    else:
        ds = double_spaced
        add_heading(doc, "Change-score analyses", level=2, double_spaced=ds)
        add_p(
            doc,
            "Because all groups improved during inpatient care, change-score analyses were used to determine whether "
            f"the magnitude of improvement differed across groups. Significant differences were observed for TMT-B ({stat_str('F', rows['TMT-B']['F'])}, "
            f"{stat_str('P', rows['TMT-B']['P'])}), VFT ({stat_str('F', rows['VFT']['F'])}, {stat_str('P', rows['VFT']['P'])}), and "
            f"HVLT-R ({stat_str('F', rows['HVLT-R']['F'])}, {stat_str('P', rows['HVLT-R']['P'])}), but not for MMSE or DST (Table 3). "
            "Post hoc comparisons confirmed greater change in the active group than in the sham group on TMT-B "
            f"({stat_str('P', cpw['TMT-B']['active_vs_sham'])}), VFT ({stat_str('P', cpw['VFT']['active_vs_sham'])}), and HVLT-R "
            f"({stat_str('P', cpw['HVLT-R']['active_vs_sham'])}), and greater change than the control group on the same measures "
            f"({stat_str('P', cpw['TMT-B']['active_vs_control'])}, {stat_str('P', cpw['VFT']['active_vs_control'])}, and "
            f"{stat_str('P', cpw['HVLT-R']['active_vs_control'])}, respectively). Sham and control groups did not differ on change scores.",
            double_spaced=ds,
            body=True,
        )
        add_table(
            doc,
            "Table 3. Baseline-to-post-intervention change scores (mean Δ±SD) across groups",
            ["Measure", f"Active (n={n['1']})", f"Sham (n={n['2']})", f"Control (n={n['3']})", "F", "P"],
            table_rows,
        )


def _ancova_f_str(r: dict) -> str:
    return f"F({int(r['df1'])},{int(r['df2'])})={fmt_f(r['F'])}"


def add_ancova_results(doc: Document, data: dict, lang: str, n: dict, *, double_spaced: bool = False) -> None:
    """Baseline-adjusted ANCOVA subsection (EN + ZH)."""
    enh = data.get("enhanced", {})
    anc = enh.get("ancova", {})
    if not anc:
        raise RuntimeError("enhanced_analyses.json missing 'ancova'; run enhanced_analyses.py first")
    outcomes = ["MMSE", "TMT-B", "DST", "VFT", "HVLT-R"]

    def am(outcome: str, key: str) -> str:
        return f"{anc[outcome]['adjusted_means'][key]:.2f}"

    table_rows = []
    for o in outcomes:
        r = anc[o]
        table_rows.append(
            [o, am(o, "active"), am(o, "sham"), am(o, "control"),
             f"{fmt_f(r['F'])} ({int(r['df1'])},{int(r['df2'])})", fmt_p(r["P"]),
             f"{r['partial_eta_sq']:.2f}"]
        )

    if lang == "zh":
        add_heading(doc, "2.6 协方差分析（基线校正）", level=2)
        add_p(
            doc,
            f"以基线评分为协变量对干预后评分进行ANCOVA：TMT-B（{_ancova_f_str(anc['TMT-B'])}, {stat_str('P', anc['TMT-B']['P'])}，偏η²="
            f"{anc['TMT-B']['partial_eta_sq']:.2f}）、VFT（{_ancova_f_str(anc['VFT'])}, {stat_str('P', anc['VFT']['P'])}，偏η²="
            f"{anc['VFT']['partial_eta_sq']:.2f}）和HVLT-R（{_ancova_f_str(anc['HVLT-R'])}, {stat_str('P', anc['HVLT-R']['P'])}，偏η²="
            f"{anc['HVLT-R']['partial_eta_sq']:.2f}）组间效应显著；MMSE与DST仍无统计学意义（表4）。校正后均数显示真刺激组在三项主要结局上均优于"
            "其余两组，且各模型回归斜率同质（交互项均P>0.11）。ANCOVA、变化值分析与未调整的干预后比较结论一致，提示结果不能归因于基线不平衡。",
        )
        add_table(
            doc,
            "表4  干预后认知评分的协方差分析（校正均数）",
            ["指标", f"真刺激组(n={n['1']})", f"伪刺激组(n={n['2']})", f"对照组(n={n['3']})", "F (df)", "P", "偏η²"],
            table_rows,
        )
    else:
        ds = double_spaced
        add_heading(doc, "Baseline-adjusted between-group comparisons (ANCOVA)", level=2, double_spaced=ds)
        add_p(
            doc,
            "As a confirmatory analysis, ANCOVA on post-intervention scores with baseline performance as covariate showed "
            f"group effects for TMT-B ({_ancova_f_str(anc['TMT-B'])}, {stat_str('P', anc['TMT-B']['P'])}, partial η²="
            f"{anc['TMT-B']['partial_eta_sq']:.2f}), VFT ({_ancova_f_str(anc['VFT'])}, {stat_str('P', anc['VFT']['P'])}, partial η²="
            f"{anc['VFT']['partial_eta_sq']:.2f}), and HVLT-R ({_ancova_f_str(anc['HVLT-R'])}, {stat_str('P', anc['HVLT-R']['P'])}, partial η²="
            f"{anc['HVLT-R']['partial_eta_sq']:.2f}); MMSE and DST remained nonsignificant (Table 4). Adjusted means placed the "
            "active group ahead of both comparison arms on all three primary outcomes, and homogeneity of regression slopes held "
            "for every model (all interaction P>0.11). The convergence of ANCOVA, change-score, and unadjusted post-intervention "
            "analyses indicates that the findings are not attributable to baseline imbalances.",
            double_spaced=ds,
            body=True,
        )
        add_table(
            doc,
            "Table 4. Analysis of covariance on post-intervention cognitive scores (adjusted means)",
            ["Measure", f"Active (n={n['1']})", f"Sham (n={n['2']})", f"Control (n={n['3']})", "F (df)", "P", "Partial η²"],
            table_rows,
        )


CORR_NAME_EN = {
    "age_first_drink": "age at first drink",
    "duration_use": "cumulative alcohol-use duration",
    "dependence_duration": "dependence duration",
    "abstinence_attempts": "prior abstinence attempts",
}
CORR_NAME_ZH = {
    "age_first_drink": "首次饮酒年龄",
    "duration_use": "累计使用时间",
    "dependence_duration": "成瘾时间",
    "abstinence_attempts": "戒酒次数",
}
COG_NAME_ZH = {"MMSE": "MMSE", "TMT-B": "TMT-B", "DST": "DST", "VFT": "VFT", "HVLT-R": "HVLT-R"}
COG_NAME_EN = {"MMSE": "MMSE", "TMT-B": "TMT-B completion time", "DST": "digit span", "VFT": "verbal fluency", "HVLT-R": "HVLT-R recall"}


def _corr_sentences(cm: dict, lang: str) -> tuple[str, str]:
    """Build (main_text, survivors_desc) for the full 20-cell correlation grid."""
    cells = cm["cells"]
    survivors = [c for c in cells if c["q_bh"] < 0.05]
    nominal_only = [c for c in cells if c["P"] < 0.05 and c["q_bh"] >= 0.05]

    def fmt_cell(c):
        return (CORR_NAME_EN[c["drinking"]], CORR_NAME_EN[c["cognition"]], c["r"], c["q_bh"])

    if lang == "zh":
        surv_txt = "；".join(
            f"{CORR_NAME_ZH[c['drinking']]}与{COG_NAME_ZH[c['cognition']]}呈{'正' if c['r']>0 else '负'}相关（r={abs(c['r']):.2f}，q={c['q_bh']:.3f}）"
            for c in survivors
        )
        main = (
            f"Pearson相关分析共检验{cm['n_tests']}个组合（4项饮酒特征×5项基线认知指标），其中{cm['nominal_sig_count']}个达到名义显著水平"
            f"（均P<0.05），经Benjamini-Hochberg校正后有{cm['fdr_sig_count']}个保持显著：{surv_txt}。"
            "TMT-B完成时间越长提示认知转换越慢，即饮酒暴露时间越长者基线认知损害越重。"
        )
        if nominal_only:
            rest = "；".join(
                f"{CORR_NAME_ZH[c['drinking']]}与{COG_NAME_ZH[c['cognition']]}（r={c['r']:.2f}）" for c in nominal_only
            )
            main += f"其余名义显著的关联——{rest}——校正后不再显著（q均>0.10）；其他组合未达名义显著。"
        return main, ""
    else:
        surv_txt = "; ".join(
            f"{'longer ' if c['drinking'] != 'age_first_drink' else 'older '}{CORR_NAME_EN[c['drinking']]} was "
            f"{'positively' if c['r'] > 0 else 'negatively'} correlated with baseline {COG_NAME_EN[c['cognition']]} "
            f"(r={c['r']:.2f}, q={c['q_bh']:.3f})"
            for c in survivors
        )
        main = (
            f"Pearson correlation analyses examined {cm['n_tests']} drinking-cognition pairs (four drinking characteristics "
            f"x five baseline cognitive measures). Of these, {cm['nominal_sig_count']} were nominally significant "
            f"(all P<0.05), and {cm['fdr_sig_count']} survived Benjamini-Hochberg correction across the twenty tests: "
            f"{surv_txt}. Because higher TMT-B completion times indicate slower set shifting, both associations reflect "
            "greater cognitive impairment with longer alcohol exposure."
        )
        if nominal_only:
            rest = "; ".join(
                f"{CORR_NAME_EN[c['drinking']]} with {COG_NAME_EN[c['cognition']]} (r={c['r']:.2f})" for c in nominal_only
            )
            qs = ", ".join(f"q={c['q_bh']:.3f}" for c in nominal_only)
            main += f" The remaining nominally significant associations ({rest}) did not survive correction ({qs})."
        return main, ""


def results_section(doc: Document, data: dict, lang: str, *, double_spaced: bool = False) -> None:
    n = data["n"]
    t1 = {r["variable"]: r for r in data["tables"]["table1"]}
    base = {r["outcome"]: r for r in data["tables"]["table_baseline"]}
    post = {r["outcome"]: r for r in data["tables"]["table_post"]}
    pw = data["pairwise_post"]
    fdr = data.get("multiplicity", {}).get("fdr_adjusted_p", {})
    enh = data.get("enhanced", {})
    es_ci = enh.get("effect_sizes_post_ci", {})

    def table_row(measure_label, row, *, unit=""):
        return [
            measure_label,
            row["active"], row["sham"], row["control"],
            fmt_f(row["F"]), fmt_p(row["P"]),
        ]

    if lang == "zh":
        add_heading(doc, "2 结果", level=1)
        add_p(doc, f"最终{data['total_n']}例纳入主要分析（图1）。", double_spaced=double_spaced, body=True)
        add_heading(doc, "2.1 三组一般资料比较", level=2)
        add_p(doc, "三组年龄、教育年限及饮酒史指标差异均无统计学意义（P>0.05），具有可比性（表1）。", body=True)
        add_table(
            doc,
            "表1  三组社会特征比较（x̄±s）",
            ["变量", f"真刺激组(n={n['1']})", f"伪刺激组(n={n['2']})", f"对照组(n={n['3']})", "F", "P"],
            [
                table_row("年龄（岁）", t1["年龄"]),
                table_row("教育年限（年）", t1["教育年限"]),
                table_row("首次饮酒年龄（岁）", t1["首次饮酒年龄"]),
                table_row("累计使用时间（年）", t1["累计使用时间"]),
                table_row("成瘾时间（年）", t1["成瘾时间"]),
                table_row("戒酒次数（次）", t1["戒酒次数"]),
            ],
        )
        add_heading(doc, "2.2 认知功能评分（基线与干预后）", level=2)
        add_p(doc, "三组基线MMSE、TMT-B、DST、VFT和HVLT-R差异均无统计学意义（P>0.05）。基线与干预后均数见表2。", body=True)
        add_table(
            doc,
            "表2  三组认知功能评分（基线与干预后，x̄±s）",
            ["指标", "时点", "真刺激组", "伪刺激组", "对照组", "F", "P"],
            [
                ["MMSE", "基线", base["MMSE"]["active"], base["MMSE"]["sham"], base["MMSE"]["control"], fmt_f(base["MMSE"]["F"]), fmt_p(base["MMSE"]["P"])],
                ["", "干预后", post["MMSE"]["active"], post["MMSE"]["sham"], post["MMSE"]["control"], fmt_f(post["MMSE"]["F"]), fmt_p(post["MMSE"]["P"])],
                ["TMT-B（s）", "基线", base["TMT-B"]["active"], base["TMT-B"]["sham"], base["TMT-B"]["control"], fmt_f(base["TMT-B"]["F"]), fmt_p(base["TMT-B"]["P"])],
                ["", "干预后", post["TMT-B"]["active"], post["TMT-B"]["sham"], post["TMT-B"]["control"], fmt_f(post["TMT-B"]["F"]), fmt_p(post["TMT-B"]["P"])],
                ["DST", "基线", base["DST"]["active"], base["DST"]["sham"], base["DST"]["control"], fmt_f(base["DST"]["F"]), fmt_p(base["DST"]["P"])],
                ["", "干预后", post["DST"]["active"], post["DST"]["sham"], post["DST"]["control"], fmt_f(post["DST"]["F"]), fmt_p(post["DST"]["P"])],
                ["VFT", "基线", base["VFT"]["active"], base["VFT"]["sham"], base["VFT"]["control"], fmt_f(base["VFT"]["F"]), fmt_p(base["VFT"]["P"])],
                ["", "干预后", post["VFT"]["active"], post["VFT"]["sham"], post["VFT"]["control"], fmt_f(post["VFT"]["F"]), fmt_p(post["VFT"]["P"])],
                ["HVLT-R", "基线", base["HVLT-R"]["active"], base["HVLT-R"]["sham"], base["HVLT-R"]["control"], fmt_f(base["HVLT-R"]["F"]), fmt_p(base["HVLT-R"]["P"])],
                ["", "干预后", post["HVLT-R"]["active"], post["HVLT-R"]["sham"], post["HVLT-R"]["control"], fmt_f(post["HVLT-R"]["F"]), fmt_p(post["HVLT-R"]["P"])],
            ],
        )
        add_within_group_results(doc, data, lang, n, double_spaced=double_spaced)
        add_change_score_results(doc, data, lang, n, double_spaced=double_spaced)
        add_heading(doc, "2.5 干预后组间认知功能", level=2)
        add_p(
            doc,
            f"干预后TMT-B（{stat_str('F', post['TMT-B']['F'])}, {stat_str('P', post['TMT-B']['P'])}, {stat_str('q', fdr.get('TMT-B'))}）、"
            f"VFT（{stat_str('F', post['VFT']['F'])}, {stat_str('P', post['VFT']['P'])}, {stat_str('q', fdr.get('VFT'))}）和"
            f"HVLT-R（{stat_str('F', post['HVLT-R']['F'])}, {stat_str('P', post['HVLT-R']['P'])}, {stat_str('q', fdr.get('HVLT-R'))}）"
            f"组间差异有统计学意义；MMSE（{stat_str('P', post['MMSE']['P'])}, {stat_str('q', fdr.get('MMSE'))}）和"
            f"DST（{stat_str('P', post['DST']['P'])}, {stat_str('q', fdr.get('DST'))}）差异无统计学意义（表2）。"
            "三项预先设定的主要结局经Benjamini-Hochberg FDR校正后仍保持显著。",
        )
        add_p(
            doc,
            f"两两比较：TMT-B真刺激组优于伪刺激组（{stat_str('P', pw['TMT-B']['active_vs_sham'])}）及对照组（{stat_str('P', pw['TMT-B']['active_vs_control'])}）；"
            f"VFT（{stat_str('P', pw['VFT']['active_vs_sham'])}/{stat_str('P', pw['VFT']['active_vs_control'])}）；"
            f"HVLT-R（{stat_str('P', pw['HVLT-R']['active_vs_sham'])}/{stat_str('P', pw['HVLT-R']['active_vs_control'])}）。"
            f"效应量方面，真刺激组对伪刺激组：TMT-B d={ci_str(es_ci['TMT-B']['active_vs_sham'])}，VFT d={ci_str(es_ci['VFT']['active_vs_sham'])}，"
            f"HVLT-R d={ci_str(es_ci['HVLT-R']['active_vs_sham'])}（Hedges g="
            f"{g_str(es_ci['TMT-B']['active_vs_sham'])}/{g_str(es_ci['VFT']['active_vs_sham'])}/"
            f"{g_str(es_ci['HVLT-R']['active_vs_sham'])}）；对对照组：d={ci_str(es_ci['TMT-B']['active_vs_control'])}、"
            f"{ci_str(es_ci['VFT']['active_vs_control'])}、{ci_str(es_ci['HVLT-R']['active_vs_control'])}（方括号为95%置信区间）。"
            "伪刺激组与对照组差异均无统计学意义（P>0.05）。干预后主要结局Levene检验均P>0.10；Kruskal-Wallis检验显著性模式一致。",
        )
        sens = data.get("sensitivity", {})
        if sens:
            s4 = sens.get("S4_as_recorded", {})
            add_p(
                doc,
                "敏感性分析（重归记录改回真刺激、剔除该记录，或5例争议受试者全部按录入臂归入真刺激）均保持三项主要结局"
                f"ANCOVA结论（均P≤0.0003；真–伪刺激|d|跨集合0.89–1.46）。最保守的按录入臂集合（真刺激n={s4.get('n_active', 26)}，"
                f"伪刺激n={s4.get('n_sham', 22)}）中，真–伪刺激d仍为"
                f"{s4.get('TMT-B', {}).get('cohen_d', -1.20):.2f}/"
                f"{s4.get('VFT', {}).get('cohen_d', 1.14):.2f}/"
                f"{s4.get('HVLT-R', {}).get('cohen_d', 0.89):.2f}。",
            )
        add_ancova_results(doc, data, lang, n, double_spaced=double_spaced)
        add_heading(doc, "2.7 基线认知与饮酒特征的相关性", level=2)
        main_txt, _ = _corr_sentences(enh.get("correlation_matrix", {}), lang)
        add_p(doc, main_txt)
    else:
        ds = double_spaced
        add_heading(doc, "Results", level=1, double_spaced=ds)
        add_p(
            doc,
            f"A total of {data['total_n']} participants were included in the primary analysis (Figure 1).",
            double_spaced=ds,
            body=True,
        )
        add_heading(doc, "Baseline characteristics", level=2, double_spaced=ds)
        add_p(doc, "The three groups did not differ in age, education, or drinking history (all P>0.05; Table 1).", double_spaced=ds, body=True)
        add_table(
            doc,
            "Table 1. Sociodemographic and drinking characteristics (mean±SD)",
            ["Variable", f"Active (n={n['1']})", f"Sham (n={n['2']})", f"Control (n={n['3']})", "F", "P"],
            [
                table_row("Age (years)", t1["年龄"]),
                table_row("Education (years)", t1["教育年限"]),
                table_row("Age at first drink", t1["首次饮酒年龄"]),
                table_row("Duration of use (years)", t1["累计使用时间"]),
                table_row("Dependence duration (years)", t1["成瘾时间"]),
                table_row("Prior abstinence attempts", t1["戒酒次数"]),
            ],
        )
        add_heading(doc, "Cognitive scores at baseline and post-intervention", level=2, double_spaced=ds)
        add_p(
            doc,
            "No baseline differences were observed across groups for MMSE, TMT-B, DST, VFT, or HVLT-R (all P>0.05). "
            "Baseline and post-intervention means are summarized in Table 2.",
            double_spaced=ds,
            body=True,
        )
        add_table(
            doc,
            "Table 2. Cognitive scores at baseline and post-intervention (mean±SD)",
            ["Measure", "Time", "Active", "Sham", "Control", "F", "P"],
            [
                ["MMSE", "Baseline", base["MMSE"]["active"], base["MMSE"]["sham"], base["MMSE"]["control"], fmt_f(base["MMSE"]["F"]), fmt_p(base["MMSE"]["P"])],
                ["", "Post", post["MMSE"]["active"], post["MMSE"]["sham"], post["MMSE"]["control"], fmt_f(post["MMSE"]["F"]), fmt_p(post["MMSE"]["P"])],
                ["TMT-B (s)", "Baseline", base["TMT-B"]["active"], base["TMT-B"]["sham"], base["TMT-B"]["control"], fmt_f(base["TMT-B"]["F"]), fmt_p(base["TMT-B"]["P"])],
                ["", "Post", post["TMT-B"]["active"], post["TMT-B"]["sham"], post["TMT-B"]["control"], fmt_f(post["TMT-B"]["F"]), fmt_p(post["TMT-B"]["P"])],
                ["DST", "Baseline", base["DST"]["active"], base["DST"]["sham"], base["DST"]["control"], fmt_f(base["DST"]["F"]), fmt_p(base["DST"]["P"])],
                ["", "Post", post["DST"]["active"], post["DST"]["sham"], post["DST"]["control"], fmt_f(post["DST"]["F"]), fmt_p(post["DST"]["P"])],
                ["VFT", "Baseline", base["VFT"]["active"], base["VFT"]["sham"], base["VFT"]["control"], fmt_f(base["VFT"]["F"]), fmt_p(base["VFT"]["P"])],
                ["", "Post", post["VFT"]["active"], post["VFT"]["sham"], post["VFT"]["control"], fmt_f(post["VFT"]["F"]), fmt_p(post["VFT"]["P"])],
                ["HVLT-R", "Baseline", base["HVLT-R"]["active"], base["HVLT-R"]["sham"], base["HVLT-R"]["control"], fmt_f(base["HVLT-R"]["F"]), fmt_p(base["HVLT-R"]["P"])],
                ["", "Post", post["HVLT-R"]["active"], post["HVLT-R"]["sham"], post["HVLT-R"]["control"], fmt_f(post["HVLT-R"]["F"]), fmt_p(post["HVLT-R"]["P"])],
            ],
        )
        add_within_group_results(doc, data, lang, n, double_spaced=ds)
        add_change_score_results(doc, data, lang, n, double_spaced=ds)
        add_heading(doc, "Post-intervention between-group comparisons", level=2, double_spaced=ds)
        add_p(
            doc,
            f"Significant group differences were observed for TMT-B ({stat_str('F', post['TMT-B']['F'])}, {stat_str('P', post['TMT-B']['P'])}, "
            f"{stat_str('q', fdr.get('TMT-B'))}), VFT ({stat_str('F', post['VFT']['F'])}, {stat_str('P', post['VFT']['P'])}, "
            f"{stat_str('q', fdr.get('VFT'))}), and HVLT-R ({stat_str('F', post['HVLT-R']['F'])}, {stat_str('P', post['HVLT-R']['P'])}, "
            f"{stat_str('q', fdr.get('HVLT-R'))}), but not for MMSE ({stat_str('P', post['MMSE']['P'])}, {stat_str('q', fdr.get('MMSE'))}) "
            f"or DST ({stat_str('P', post['DST']['P'])}, {stat_str('q', fdr.get('DST'))}; Table 2). "
            "All three prespecified primary outcomes remained significant after Benjamini-Hochberg FDR correction.",
            double_spaced=ds,
            body=True,
        )
        add_p(
            doc,
            f"Post hoc comparisons showed that the active group outperformed the sham group on TMT-B ({stat_str('P', pw['TMT-B']['active_vs_sham'])}), "
            f"VFT ({stat_str('P', pw['VFT']['active_vs_sham'])}), and HVLT-R ({stat_str('P', pw['HVLT-R']['active_vs_sham'])}), and surpassed the control group on the same measures "
            f"({stat_str('P', pw['TMT-B']['active_vs_control'])}, {stat_str('P', pw['VFT']['active_vs_control'])}, and {stat_str('P', pw['HVLT-R']['active_vs_control'])}, respectively). "
            f"For active versus sham, Cohen's d was {ci_str(es_ci['TMT-B']['active_vs_sham'])} for TMT-B, {ci_str(es_ci['VFT']['active_vs_sham'])} for VFT, and "
            f"{ci_str(es_ci['HVLT-R']['active_vs_sham'])} for HVLT-R (Hedges g="
            f"{g_str(es_ci['TMT-B']['active_vs_sham'])}, {g_str(es_ci['VFT']['active_vs_sham'])}, and "
            f"{g_str(es_ci['HVLT-R']['active_vs_sham'])}); versus control, d was {ci_str(es_ci['TMT-B']['active_vs_control'])}, "
            f"{ci_str(es_ci['VFT']['active_vs_control'])}, and {ci_str(es_ci['HVLT-R']['active_vs_control'])}, respectively (95% confidence intervals in brackets). "
            "Sham and control groups did not differ on any outcome. Levene tests for post-intervention primary "
            "outcomes were nonsignificant (all P>0.10); Kruskal-Wallis tests yielded the same significance pattern.",
            double_spaced=ds,
            body=True,
        )
        sens = data.get("sensitivity", {})
        if sens:
            s4 = sens.get("S4_as_recorded", {})
            add_p(
                doc,
                "Sensitivity analyses that reassigned the reclassified participant to the active arm, excluded that "
                "participant, or analyzed all five ambiguous participants in the active arm as recorded preserved the "
                "primary ANCOVA conclusions for TMT-B, VFT, and HVLT-R (all ANCOVA P≤0.0003; active-versus-sham "
                f"|d| range across sets 0.89–1.46). In the most conservative as-recorded set (active n={s4.get('n_active', 26)}, "
                f"sham n={s4.get('n_sham', 22)}), active-versus-sham d remained "
                f"{s4.get('TMT-B', {}).get('cohen_d', -1.20):.2f}/"
                f"{s4.get('VFT', {}).get('cohen_d', 1.14):.2f}/"
                f"{s4.get('HVLT-R', {}).get('cohen_d', 0.89):.2f}.",
                double_spaced=ds,
            )
        add_ancova_results(doc, data, lang, n, double_spaced=ds)
        add_heading(doc, "Correlation analyses", level=2, double_spaced=ds)
        main_txt, _ = _corr_sentences(enh.get("correlation_matrix", {}), lang)
        add_p(doc, main_txt, double_spaced=ds, body=True)


def add_references(doc: Document, lang: str, nummap: dict) -> None:
    add_heading(doc, "参考文献" if lang == "zh" else "References", level=1)
    for i, ref in enumerate(reference_list(nummap), start=1):
        add_p(doc, f"[{i}] {ref}")


def add_declarations(doc: Document, lang: str) -> None:
    add_heading(doc, "伦理与声明" if lang == "zh" else "Ethics and Declarations", level=1)
    if lang == "zh":
        add_p(doc, "伦理批准：新乡医学院第二附属医院伦理委员会（批件号见文末作者待填清单）。")
        add_p(doc, "知情同意：所有受试者及其监护人签署书面知情同意书。")
        add_p(doc, "利益冲突：作者声明不存在利益冲突。")
        add_p(doc, "作者贡献：见文末作者待填清单。")
        add_p(doc, "资助信息：见文末作者待填清单。")
        add_p(doc, "数据可用性：去标识化个体数据因机构隐私保护及受试者未授权公开共享，暂不对外公开；可复现报告推断性结果的汇总统计与结构化分析输出见项目补充文件。去标识化数据可向通讯作者合理申请，并需履行机构审批程序。")
        add_p(doc, "代码可用性：统计分析与稿件生成脚本位于 03_analysis/ 目录（主入口：clean_and_analyze.py、build_full_manuscript.py）。尚未分配永久公开仓库地址；脚本可向通讯作者申请获取。")
        add_p(doc, "试验注册：本研究未在入组前前瞻注册（2019年入组）；投稿前拟进行回顾性注册并披露注册日期。")
    else:
        add_p(doc, "Ethics approval: Institutional Ethics Committee of the Second Affiliated Hospital of Xinxiang Medical University (approval number: see Author Completion Checklist).")
        add_p(doc, "Consent: Written informed consent was obtained from all participants and their guardians.")
        add_p(doc, "Competing interests: The authors declare no competing interests.")
        add_p(doc, "Author contributions: see Author Completion Checklist.")
        add_p(doc, "Funding: see Author Completion Checklist.")
        add_p(doc, "Data availability: De-identified individual-level data are not publicly available because of institutional privacy restrictions and the absence of participant consent for open sharing. Summary statistics and structured analysis outputs sufficient to reproduce the reported inferential results are maintained in the study supplement files. De-identified data may be available from the corresponding author upon reasonable request and with appropriate institutional agreements.")
        add_p(doc, "Code availability: Statistical analysis and manuscript generation scripts are maintained in the study analysis folder (03_analysis/). Key entry points are clean_and_analyze.py and build_full_manuscript.py. A permanent public repository URL has not yet been assigned; scripts are available from the corresponding author upon request.")


def build_zh(data: dict, power: dict) -> None:
    doc = Document()
    set_doc_font(doc)
    nummap = build_nummap(data, power)  # ZH reuses the EN-driven canonical numbering
    zh_texts = list(INTRO_ZH) + methods_zh(data) + methods_supplement_zh(data, power) + DISCUSSION_ZH
    zh_keys = {
        tok.lstrip("@")
        for t in zh_texts
        for m in re.findall(r"\[@([^\]]+)\]", t)
        for tok in m.split(",")
        if tok.strip()
    }
    unknown = zh_keys - set(nummap)
    if unknown:
        raise RuntimeError(f"ZH text cites keys absent from EN nummap: {sorted(unknown)}")
    render = lambda t: render_citations(t, nummap)

    add_title(doc, TITLE_ZH)
    add_p(doc, "【摘要】" + abstract_zh(data))
    add_p(doc, "【关键词】酒依赖；高精度经颅直流电刺激；认知功能；随机对照试验")
    add_p(doc, "Abstract: " + abstract_en(data))
    add_p(doc, f"Keywords: {KEYWORDS_EN}")

    add_heading(doc, "引言", level=1)
    for para in INTRO_ZH:
        add_p(doc, render(para), body=True)

    add_heading(doc, "1 对象与方法", level=1)
    for para in methods_zh(data):
        add_p(doc, render(para), body=True)
    for para in methods_supplement_zh(data, power):
        add_p(doc, render(para), body=True)

    add_consort_figure(doc, "zh")
    add_result_figures(doc, "zh")
    results_section(doc, data, "zh")
    add_safety_results(doc, data, "zh", nummap)

    add_heading(doc, "3 讨论", level=1)
    for para in DISCUSSION_ZH:
        add_p(doc, render(para), body=True)

    add_declarations(doc, "zh")
    add_references(doc, "zh", nummap)
    add_analysis_set_supplement(doc, "zh")
    add_author_completion_checklist(doc, "zh")
    doc.save(OUT_ZH)
    print("Wrote", OUT_ZH)


def build_en(data: dict, power: dict) -> None:
    """English manuscript synced with BRS content (single-spaced working layout)."""
    doc = Document()
    set_doc_font(doc, western="Times New Roman", east="Times New Roman")
    assemble_english_manuscript(doc, data, power, double_spaced=False, brs_title_page=False)
    doc.save(OUT_EN)
    print("Wrote", OUT_EN)


def build_en_brs(data: dict, power: dict) -> None:
    """Brain Stimulation submission layout (double-spaced, title page)."""
    doc = Document()
    set_doc_font(doc, western="Times New Roman", east="Times New Roman")
    assemble_english_manuscript(doc, data, power, double_spaced=True, brs_title_page=True)
    doc.save(OUT_EN_BRS)
    print("Wrote", OUT_EN_BRS)


def count_main_text_words(doc: Document, *, brs: bool) -> int:
    """Count words in Introduction through Discussion (BRS body limit ~4000)."""
    start_markers = {"Introduction", "1 引言", "引言"}
    stop_markers = {
        "Acknowledgments",
        "Financial Disclosures",
        "References",
        "Figure and Table Legends",
        "伦理与声明",
        "参考文献",
    }
    counting = False
    words = 0
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name if para.style else ""
        if not counting:
            if text in start_markers or style.startswith("Heading 1") and text in start_markers:
                counting = True
                continue
            if text == "Introduction":
                counting = True
                continue
        if counting:
            if text in stop_markers or (style.startswith("Heading 1") and text in stop_markers):
                break
            if text.startswith("Table ") or text.startswith("表"):
                continue
            words += len(text.split())
    return words


def report_word_counts(data: dict) -> None:
    from docx import Document as Doc

    for path, brs in [(OUT_EN_BRS, True), (OUT_EN, False)]:
        if not path.exists():
            continue
        doc = Doc(str(path))
        wc = count_main_text_words(doc, brs=brs)
        label = "BRS" if brs else "EN"
        status = "OK" if wc <= 4000 else "REVIEW"
        print(f"Main text word count ({label}): {wc} [{status} vs BRS ~4000 limit]")
        if brs:
            out = ROOT / "04_data" / "processed" / "manuscript_word_count.json"
            out.write_text(
                json.dumps({"brs_main_text_words": wc, "limit": 4000, "status": status}, indent=2) + "\n",
                encoding="utf-8",
            )


def main() -> None:
    from consort_flow import main as generate_consort_figure
    from build_submission_materials import main as build_submission_materials

    check_highlights(HIGHLIGHTS_EN)
    generate_consort_figure()
    for fig_mod in ("figure_montage", "figure_outcomes", "figure_change"):
        try:
            mod = __import__(fig_mod)
            if hasattr(mod, "main"):
                mod.main()
        except Exception as exc:  # figures are best-effort if scripts not yet present
            print(f"Figure module {fig_mod} skipped: {exc}")
    data = load_stats()
    if ENHANCED_PATH.exists():
        data["enhanced"] = json.loads(ENHANCED_PATH.read_text(encoding="utf-8"))
    else:
        from enhanced_analyses import main as run_enhanced

        run_enhanced()
        data["enhanced"] = json.loads(ENHANCED_PATH.read_text(encoding="utf-8"))
    if SENSITIVITY_PATH.exists():
        data["sensitivity"] = json.loads(SENSITIVITY_PATH.read_text(encoding="utf-8"))
    else:
        data["sensitivity"] = {}
    power = compute_post_hoc_power(data)
    build_zh(data, power)
    build_en(data, power)
    build_en_brs(data, power)
    report_word_counts(data)
    build_submission_materials()


if __name__ == "__main__":
    main()
