#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""构建 Frontiers 修回材料：修订稿三版 + 回复信 + 补充材料 + 待作者补清单。

产出（01_稿件/）：
  Manuscript_Frontiers_修订稿_含修订痕迹.docx   投稿用（Word 原生修订痕迹）
  Manuscript_Frontiers_修订稿_干净版.docx       干净版
  Manuscript_Frontiers_修订稿_高亮版.docx       内部审阅版（黄色高亮 + 删除线）
  修订记录/Response_to_Reviewers.docx          逐条回复信
  修订记录/Supplementary_Material.docx         补充材料（Table S1–S6）
  修订记录/待作者补清单.docx                    待作者补的信息

用法：python3 03_scripts/build_revision.py [--repo-public|--repo-private]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_edit import Docx  # noqa: E402
import revision_content as RC  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "01_稿件" / "Manuscript_Frontiers_投稿版_1016.docx"
RESULTS = json.loads((ROOT / "06_results" / "复算结果_全量.json").read_text(encoding="utf-8"))
SENS = json.loads((ROOT / "04_raw_data" / "upstream_repo" / "04_data" / "processed" /
                   "sensitivity_analyses.json").read_text(encoding="utf-8"))
AUTHOR = "Authors"
DATE = "2026-10-27T00:00:00Z"

MODES = {
    "tracked": ("Manuscript_Frontiers_修订稿_含修订痕迹.docx", "tracked"),
    "clean": ("Manuscript_Frontiers_修订稿_干净版.docx", "clean"),
    "highlight": ("Manuscript_Frontiers_修订稿_高亮版.docx", "highlight"),
}


# ------------------------------------------------------------------ 稿件

def pick_heading_template(src: Docx, style: str = "Heading2"):
    """挑一个同级别的既有标题段��做模板，让新标题继承原稿的加粗等 run 格式。"""
    from docx_edit import qn
    for p in src.paragraphs():
        pPr = p.find(qn("pPr"))
        if pPr is None:
            continue
        ps = pPr.find(qn("pStyle"))
        if ps is not None and ps.get(qn("val")) == style:
            return p
    return None


def pick_body_template(src: Docx):
    """挑一个无 run 级粗体的正文段落作为模板，避免新段落继承粗体。"""
    from docx_edit import qn
    for p in src.paragraphs():
        pPr = p.find(qn("pPr"))
        if pPr is not None and pPr.find(qn("pStyle")) is not None:
            continue
        runs = p.findall(qn("r"))
        if not runs:
            continue
        ok = True
        for r in runs:
            rpr = r.find(qn("rPr"))
            if rpr is not None and rpr.find(qn("b")) is not None:
                ok = False
                break
        if ok:
            return p
    return src.paragraphs()[10]


def build_manuscript(mode: str, out_name: str, repo_public: bool) -> Path:
    d = Docx(SRC)
    template_table = d.tables()[0]          # Table 1，作为新表样式模板
    src = Docx(SRC)
    body_tpl = pick_body_template(src)

    # 1) 段落改写
    for anchor, new_text in RC.edits(RESULTS):
        p = d.find_paragraph(anchor)
        d.set_paragraph_text(p, new_text, mode, AUTHOR, DATE)

    # 2) Data / Code availability / Trial registration
    data_av, code_av = RC.data_availability_text(repo_public)
    d.set_paragraph_text(d.find_paragraph("Data availability:"), data_av, mode, AUTHOR, DATE)
    d.set_paragraph_text(d.find_paragraph("Code availability:"), code_av, mode, AUTHOR, DATE)
    d.set_paragraph_text(d.find_paragraph("Trial registration:"), RC.trial_registration_text(),
                         mode, AUTHOR, DATE)

    # 3) 新增小节
    head_tpl = pick_heading_template(src, "Heading2")
    for sec in RC.new_sections(RESULTS):
        anchor_p = d.find_paragraph(sec["anchor"])
        h = d.insert_paragraph_after(anchor_p, sec["title"], style="Heading2", mode=mode,
                                     author=AUTHOR, date=DATE, template=head_tpl or body_tpl,
                                     inherit_rpr=True if head_tpl is not None else None)
        d.insert_paragraph_after(h, sec["text"], mode=mode, author=AUTHOR, date=DATE,
                                 template=body_tpl)

    out = ROOT / "01_稿件" / out_name
    d.save(out)
    return out


# ------------------------------------------------------------------ 补充材料

def build_supplementary(repo_public: bool) -> Path:
    d = Docx(SRC)
    d.clear_body()
    src = Docx(SRC)
    tpl_table = src.tables()[0]
    body_tpl = pick_body_template(src)

    def add_p(text, style=None, bold=False):
        p = d._new_paragraph(text, style, "clean", AUTHOR, DATE, template=body_tpl)
        if bold and p.find(dq("r")) is not None:
            rpr = p.find(dq("r")).find(dq("rPr"))
            if rpr is None:
                rpr = etree_sub(p, "rPr")
            etree_sub(rpr, "b")
        d.append(p)
        return p

    def dq(tag):
        from docx_edit import qn
        return qn(tag)

    def etree_sub(parent, tag):
        from lxml import etree
        from docx_edit import qn
        return etree.SubElement(parent, qn(tag))

    def add_table(caption, header, rows, widths=None):
        add_p(caption)
        tbl = d.build_table(tpl_table, header, rows, mode="clean", author=AUTHOR, date=DATE,
                            widths=widths)
        d.append(tbl)
        add_p("")

    add_p("Supplementary Material", style="Heading1")
    add_p(
        "Effects of High-Definition Transcranial Direct Current Stimulation on Cognitive Function in "
        "Male Inpatients with Alcohol Dependence: A Randomized Sham-Controlled Trial")
    add_p(
        "This file contains the supplementary analyses referenced in the main text. Every value was "
        "recomputed from the de-identified individual-level dataset with the scripts deposited in the "
        "repository cited under Data availability; the derivation of each table is given in the notes "
        "below the tables and in the corresponding script.")

    # ---- Table S1
    flow = RESULTS["sample_flow"]
    add_table(
        "Table S1. Participant flow and derivation of the primary analysis set.",
        ["Stage", "n", "Note"],
        [
            ["Randomized", str(flow["randomized"]),
             "1:1:1 allocation: routine care 29, active HD-tDCS 28, sham HD-tDCS 28"],
            ["Discontinued before the week-4 assessment", "3",
             "Early hospital discharge (1 active, 2 sham); no week-4 outcome available"],
            ["Records with a week-4 assessment", str(sum(v for k, v in flow["records_with_week4_assessment"].items() if k != "total")),
             "Active 27, sham 26, control 29"],
            ["Erroneous active-arm records removed", "5",
             "Four cross-arm misentries and one duplicate row; each duplicated a participant already "
             "represented in the sham arm, so no participant outcome was discarded"],
            ["Retained record reassigned to the sham arm", "1",
             "Applied according to the original randomization log"],
            ["Primary analysis set", str(sum(v for k, v in flow["unique_participants_in_analysis"].items() if k != "total")),
             "Active 21, sham 27, control 29; one record per participant"],
            ["Arithmetic check", "—",
             "85 randomized = 3 discontinued + 82 records with a week-4 assessment; 77 = 82 − 5 "
             "erroneous records"],
        ], widths=[3.0, 0.7, 6.3])
    add_p(
        "Note: the five removed rows were a data-quality correction, not a loss of outcome data. The "
        "three participants discharged early are the only randomized participants without a week-4 "
        "assessment; they are addressed in the worst-case analysis of Table S3.")

    # ---- Table S2
    rows = []
    set_labels = {
        "S1_primary": "S1: primary set (reclassified record analysed in the sham arm)",
        "S2_lei_as_active": "S2: reclassified record analysed in the active arm",
        "S3_lei_excluded": "S3: reclassified record excluded",
        "S4_as_recorded": "S4: all five ambiguous records analysed in the active arm as recorded",
    }
    for key, label in set_labels.items():
        s = SENS[key]
        ns = f"{s['n_active']}/{s['n_sham']}/{s['n_control']}"
        for oc in ("TMT-B", "VFT", "HVLT-R"):
            o = s[oc]
            rows.append([label, ns, oc, f"{o['ancova_F']} ({o['ancova_df']})", str(o["ancova_P"]),
                         f"{o['cohen_d']}", f"{o['hedges_g']}"])
    add_table(
        "Table S2. Sensitivity analyses for the post-randomization record-verification step.",
        ["Analysis set", "n (active/sham/control)", "Outcome", "ANCOVA F (df)", "ANCOVA P",
         "Cohen d", "Hedges g"],
        rows, widths=[4.2, 1.2, 0.9, 1.4, 0.9, 0.7, 0.7])
    add_p(
        "Note: d and g are the active-versus-sham contrasts on post-intervention scores. All four sets "
        "give the same qualitative conclusion for the three primary outcomes.")

    # ---- Table S3
    tp = RESULTS["tipping_point"]["outcomes"]
    rows = []
    for oc in ("TMT-B", "VFT", "HVLT-R"):
        t = tp[oc]
        rows.append([
            oc,
            RC.p3(t["observed_active_vs_sham_P"]),
            str(t["observed_d"]),
            f"{t['tipping']['delta_sd']} SD ({t['tipping']['delta_units']} units)",
            f"{RC.f2(t['tipping']['imputed_active_post'])} / {RC.f2(t['tipping']['imputed_sham_post'])}",
            RC.p3(t["extreme_worst_case_P"]),
        ])
    add_table(
        "Table S3. Worst-case analysis for the three participants without a week-4 assessment.",
        ["Outcome", "Observed active-vs-sham P", "Observed d",
         "Offset at which significance is lost", "Implied active / sham score at that offset",
         "Extreme-case P"],
        rows, widths=[0.9, 1.3, 0.7, 2.7, 2.3, 1.1])
    add_p(
        "Note: the offset is applied in the direction unfavourable to the active arm (the missing "
        "active-arm participant downwards and the two missing sham-arm participants upwards, each "
        "scaled to that arm's standard deviation). The extreme case assigns the missing active-arm "
        "participant the worst observed active-arm score and the two missing sham-arm participants the "
        "best observed sham-arm scores. The implied scores at the tipping point lie outside the range "
        "the instruments can take (a negative TMT-B completion time and an HVLT-R score above the "
        "36-point maximum are unattainable), so no physically possible outcome for these three "
        "participants overturns the conclusion. Multiple imputation was not possible because the "
        "baseline covariates of the three participants are absent from the de-identified dataset.")

    # ---- Table S4
    perm = RESULTS["permutation"]["outcomes"]
    boot = RESULTS["bootstrap"]["outcomes"]
    rows = []
    for oc in ("TMT-B", "VFT", "HVLT-R"):
        rows.append([
            oc,
            RC.p3(perm[oc]["perm_p_omnibus"]),
            RC.p4(perm[oc]["perm_p_active_vs_sham"]),
            f"{RC.f2(boot[oc]['change']['observed_diff'])} "
            f"[{RC.f2(boot[oc]['change']['boot_ci_lo'])}, {RC.f2(boot[oc]['change']['boot_ci_hi'])}]",
            "Yes" if boot[oc]["change"]["excludes_zero"] else "No",
        ])
    add_table(
        "Table S4. Permutation tests and bootstrap confidence intervals for the primary contrasts.",
        ["Outcome", "Permutation P (omnibus, 10,000 permutations)",
         "Permutation P (active vs sham, restricted)",
         "Bootstrap 95% CI, active-minus-sham change score", "Excludes zero"],
        rows, widths=[0.9, 2.0, 2.0, 2.7, 1.0])
    add_p(
        "Note: the omnibus permutation test permutes all 77 group labels; the active-versus-sham "
        "contrast uses a restricted permutation within the 48 active and sham participants, with the "
        "control arm held fixed. Bootstrap intervals are percentile intervals from 10,000 resamples "
        "(seed 20261027) of the active-minus-sham difference in baseline-to-post change scores.")

    # ---- Table S5
    diag = RESULTS["diagnostics"]["outcomes"]
    rows = []
    for oc in ("MMSE", "TMT-B", "DST", "VFT", "HVLT-R"):
        g = diag[oc]
        shap = g["shapiro_W_P_by_arm"]
        rows.append([
            oc,
            f"{shap['active']} / {shap['sham']} / {shap['control']}",
            str(g["levene_P"]),
            RC.prange(g['loo_active_vs_sham_P_range'][0], g['loo_active_vs_sham_P_range'][1]),
            "Yes" if g["all_loo_P_below_0_05"] else "No",
        ])
    add_table(
        "Table S5. Distribution diagnostics and leave-one-participant-out stability.",
        ["Outcome", "Shapiro-Wilk P (active/sham/control)", "Levene P",
         "Leave-one-out active-vs-sham P range", "All leave-one-out P < 0.05"],
        rows, widths=[0.9, 2.6, 0.9, 2.2, 1.2])
    add_p(
        "Note: the leave-one-out analysis removes each of the 77 participants in turn and recomputes "
        "the active-versus-sham comparison. The range column gives the smallest and largest P value "
        "obtained across the 77 re-analyses.")

    # ---- Table S6
    repo_rows = [
        ["data/deid/比较_deid.xlsx", "De-identified analysis dataset, 77 records × 24 variables, "
         "one record per participant"],
        ["data/deid/真刺激组_deid.xlsx", "Active arm (n=21)"],
        ["data/deid/伪刺激组_deid.xlsx", "Sham arm (n=27)"],
        ["data/deid/酒依赖组_deid.xlsx", "Routine-care arm (n=29)"],
        ["data/deid/ambiguous_participants.json", "Codes for the five participants affected by the "
         "record-verification step"],
        ["data/deid/DEID_REPORT.json", "De-identification report: source file, sheet, row and column "
         "counts, and the action applied to each field"],
        ["03_scripts/recompute_all.py", "Recomputation of every reported statistic and the added "
         "robustness analyses (permutation, bootstrap, worst-case, diagnostics)"],
        ["06_results/", "Derived result tables in CSV and JSON, including the upstream-consistency check"],
    ]
    add_table("Table S6. Contents of the deposited reproducibility package.", ["File", "Description"],
              repo_rows, widths=[3.2, 6.8])
    add_p(
        "De-identification: participant names were removed at source and replaced with hashed study "
        "identifiers; the mapping between names and codes was kept locally and is not part of the "
        "package. No direct identifier, date of birth, or free-text clinical note is present in the "
        "de-identified tables.")

    out = ROOT / "01_稿件" / "修订记录" / "Supplementary_Material.docx"
    d.save(out)
    return out


# ------------------------------------------------------------------ 回复信

def build_response(repo_public: bool) -> Path:
    d = Docx(SRC)
    d.clear_body()
    src = Docx(SRC)
    body_tpl = pick_body_template(src)
    tpl_table = src.tables()[0]

    def add(text, style=None):
        p = d._new_paragraph(text, style, "clean", AUTHOR, DATE, template=body_tpl)
        d.append(p)
        return p

    add("Response to Reviewers", style="Heading1")
    add("Manuscript: Effects of High-Definition Transcranial Direct Current Stimulation on Cognitive "
        "Function in Male Inpatients with Alcohol Dependence: A Randomized Sham-Controlled Trial")
    add("Journal: Frontiers in Psychology — Manuscript ID 1308143")
    add("")
    add("We thank the reviewers for the careful reading of our manuscript and for the constructiveness "
        "of the comments. The revision below answers each point, and every change is marked in the "
        "revised manuscript (Word tracked changes). In summary: we have made the allocation-concealment "
        "gap explicit instead of implicit; we have documented exactly how the analysis set was derived "
        "and what the record-verification step did and did not change; we have added a worst-case "
        "analysis for the three participants without a week-4 assessment, together with permutation "
        "tests, bootstrap confidence intervals and a leave-one-participant-out analysis; we have "
        "expanded the technical description of the stimulation; and we have added a Methods subsection "
        "specifying the routine inpatient treatment that all three arms shared. All reported statistics "
        "were recomputed from the individual-level records, and the recomputation reproduces every "
        "checked value of the previous version exactly.")

    # ---------------- Reviewer 2
    add("Reviewer 2", style="Heading2")
    add("We thank Reviewer 2 for the detailed methodological review. The comment raises five specific "
        "areas — the integrity of randomization and post-randomization data handling, the choice of "
        "analysis population, trial registration and prespecification, statistical methodology, and the "
        "technical description of the HD-tDCS intervention. We answer each of them below, and the table "
        "at the end of this letter lists where each change appears in the revised manuscript. The "
        "comment also refers to a detailed report attached to the original submission; the corresponding "
        "entry (“Review supporting file”, reference 1308143) is present in the review system, but the "
        "attachment itself was not delivered to us and could not be opened from our side. We therefore "
        "ask the editorial office to forward it if it contains further points, and we will respond to "
        "each of them without delay.")

    r2 = [
        ("Point 1. Integrity of randomization and post-randomization data handling",
         "We agree that this needed to be stated rather than implied, and we have made three changes. "
         "First, the Methods now state explicitly that the mechanism used to conceal the upcoming "
         "assignment from enrolment staff was not documented in the archived protocol and could not be "
         "reconstructed, and that the sequence was therefore not verifiably concealed; this is listed as "
         "a trial limitation rather than left to inference. Second, the record-verification procedure "
         "is now described in terms of what it did and did not change: each of the five removed rows "
         "duplicated a participant who was already represented in the sham arm, so removing them "
         "discarded no participant outcome and changed no participant's allocation; the correction "
         "changed the number of records per participant, not the set of participants with data. Third, "
         "a new supplementary table (Table S1) reconciles the arithmetic of the analysis set step by "
         "step: 85 randomized, 3 participants discharged early without a week-4 assessment, 82 records "
         "with a week-4 assessment, 5 erroneous active-arm records removed, 1 retained record "
         "reassigned to the sham arm, 77 participants in the primary analysis, one record per "
         "participant. The same reconciliation is now stated in the Methods and in the Results, and the "
         "flow diagram (Figure 1) is consistent with it."),

        ("Point 2. Choice of analysis population",
         "The analysis set is now defined precisely: it contains every randomized participant who has a "
         "week-4 assessment, with one record per participant. The three participants who were "
         "discharged early are the only randomized participants without a week-4 assessment, and we no "
         "longer describe the analysis set as if it were an intention-to-treat set. To show that these "
         "three participants cannot overturn the findings, we added a worst-case analysis: their "
         "missing week-4 scores are imputed at progressively larger offsets in the direction "
         "unfavourable to the active arm, and the offset at which the active-versus-sham comparison "
         "would lose significance is identified (Table S3). That offset is 3.55 standard deviations for "
         "TMT-B, 4.0 for VFT and 5.5 for HVLT-R, and at those points the implied scores are outside the "
         "range the instruments can take (a negative TMT-B completion time, an HVLT-R score above the "
         "36-point maximum), so no physically possible outcome for these three participants overturns "
         "the conclusion. In a deterministic extreme case, assigning the missing active-arm participant "
         "the worst observed active-arm score and the two missing sham-arm participants the best "
         "observed sham-arm scores, the comparison remains significant for all three primary outcomes "
         "(TMT-B P=0.010, VFT P=0.013, HVLT-R P=0.001). We also state why multiple imputation was not "
         "used: the three participants' baseline covariates are absent from the de-identified dataset, "
         "so no imputation model could be fitted."),

        ("Point 3. Trial registration and prespecification",
         "We have stated the position plainly rather than leaving it to the reader. The trial was not "
         "prospectively registered before enrolment in 2019, no analysis plan was registered or "
         "published, and all analyses reported in the manuscript are therefore post hoc with respect "
         "to registration. The Methods now state that the outcome hierarchy (TMT-B, VFT and HVLT-R as "
         "primary outcomes; MMSE and DST as secondary outcomes) and the analysis framework were fixed "
         "before the present re-analysis of the de-identified dataset was run, and that the "
         "permutation, bootstrap and worst-case analyses were added during this revision in response "
         "to reviewer comments and are identified as such. The retrospective registration identifier "
         "and date will be inserted in the manuscript and in the submission system once assigned."),

        ("Point 4. Statistical methodology",
         "Every statistic in the manuscript was recomputed from the individual-level records in Python "
         "(SciPy 1.17, statsmodels 0.14) and cross-checked against the original SPSS 22.0 output; the "
         "recomputation reproduces all 109 checked values of the previous version exactly, and the "
         "check is deposited with the scripts. The analysis section now describes the family structure "
         "explicitly: which P values are FDR-adjusted across the five post-intervention omnibus tests, "
         "which families are reported with nominal P values as confirmatory parallel analyses, and "
         "which analyses are exploratory. Four robustness analyses were added at revision and are "
         "labelled as reviewer-requested: permutation tests (10,000 permutations; the omnibus test "
         "permutes all 77 labels, the active-versus-sham contrast uses a restricted permutation within "
         "the 48 active and sham participants), bootstrap percentile confidence intervals (10,000 "
         "resamples) for the active-minus-sham difference in both post-intervention scores and change "
         "scores, a leave-one-participant-out re-analysis over all 77 participants, and the "
         "worst-case analysis described above. All agree with the parametric results: permutation "
         "P=0.0016, 0.001 and 0.0002 for the active-versus-sham contrast on TMT-B, VFT and HVLT-R, and "
         "the bootstrap intervals for the change-score differences exclude zero for all three primary "
         "outcomes while including zero for MMSE and DST. A diagnostics table (Table S5) reports "
         "Shapiro-Wilk and Levene results and the leave-one-out P ranges."),

        ("Point 5. Technical description of the HD-tDCS intervention",
         "The Procedures section has been expanded with the parameters requested: a Soterix MXN-9 "
         "HD-tES device in direct-current mode; bilateral 4×1 ring montages with the F3 anode and the "
         "F4 cathode and the four return electrodes over each hemisphere, named with the standard "
         "10-20 scalp nomenclature; conductive gel rather than saline-soaked sponge pads at the "
         "electrode-scalp interface; ±1.0 mA for 20 min once daily for 14 consecutive sessions; "
         "continuous direct current throughout each active session; and a sham protocol in which "
         "current is applied only during the first and last 30 s of each session, with the ramp "
         "periods not disclosed to participants. The electrode size and material and the "
         "session-adherence verification procedure are to be completed from the device supply and "
         "session records, and are flagged in the Author Completion Checklist rather than guessed."),
    ]
    for title, body in r2:
        add(title, style="Heading3")
        add(body)
        add("")

    # ---------------- Reviewer 3
    add("Reviewer 3", style="Heading2")
    add("We thank Reviewer 3 for the positive assessment of the analyses and for the suggestion to "
        "specify the routine procedures, which is a fair and important point: all three arms improved "
        "on the cognitive battery during hospitalization, and the routine background care is the "
        "context in which that improvement occurred.")
    add("Change 1. New Methods subsection “Routine inpatient care”.", style="Heading3")
    add("The new subsection states that all three arms received the same background treatment for "
        "alcohol dependence, namely benzodiazepine-assisted withdrawal management followed by "
        "supportive inpatient care; that no arm received structured cognitive training, neurofeedback "
        "or any other neuromodulation during the study period; that the routine-care arm was never "
        "exposed to the device and therefore did not receive the additional 20-min device-contact time "
        "that the active and sham arms completed on 14 occasions; and that because all three arms "
        "improved during hospitalization, the between-arm contrast rather than the within-arm change "
        "is the interpretable quantity, with improvement common to all arms reflecting the natural "
        "course of early abstinence under routine care. The specific withdrawal medications, dosing "
        "schedule and any psychological interventions delivered during hospitalization are to be "
        "completed from the ward records and are flagged in the Author Completion Checklist.")
    add("Change 2. Discussion.", style="Heading3")
    add("The Discussion now states this explicitly and, in addition, is more careful about what the "
        "sham-versus-control comparison can support. Because the control arm did not receive matched "
        "device-contact time, that comparison cannot separate the effect of the delivered current from "
        "any effect of daily device contact; the active-versus-sham contrast, in which both arms "
        "received identical device contact, is the estimate that isolates the delivered current. We "
        "agree with the reviewer's emphasis that the technique is an adjunct to, and not a replacement "
        "for, routine inpatient care, and we have kept that framing in the conclusion.")
    add("Change 3. Value to the field.", style="Heading3")
    add("We have kept the conclusion scoped to what the trial can support: the adjunctive use of "
        "focal bilateral 4×1 HD-tDCS during inpatient detoxification, with preregistered multicenter "
        "trials, matched attention controls and relapse follow-up required before the cognitive gains "
        "can be said to translate into durable clinical benefit.")

    # ---------------- 改动位置对照
    add("Summary of changes and their locations", style="Heading2")
    add("Location | Change", style="Heading3")
    rows = [
        ["Abstract, Methods and Results sentences; Highlights",
         "Analysis set defined as every randomized participant with a week-4 assessment; permutation, "
         "bootstrap and worst-case analyses added"],
        ["Methods — Participants; Supplementary Table S1",
         "Step-by-step derivation of the analysis set (85 → 82 → 77) and the statement that removing "
         "the five erroneous rows discarded no participant outcome"],
        ["Methods — Randomization and allocation",
         "Allocation-concealment gap stated explicitly and listed as a limitation"],
        ["Methods — Procedures",
         "Expanded technical description: device, DC mode, 4×1 montages, 10-20 nomenclature, gel "
         "electrode interface, current, duration, session number, sham ramp protocol"],
        ["Methods — Routine inpatient care (new)",
         "Common background treatment, absence of device contact in the routine-care arm, and the "
         "interpretation of the common within-arm improvement"],
        ["Methods — Statistical analysis; Methods — Outcome hierarchy, sample size, and multiple testing",
         "Descriptions of the added permutation, bootstrap and worst-case analyses; explicit statement "
         "of the registration status and of which analyses were added at revision"],
        ["Results — Robustness of the primary contrasts (new)",
         "Permutation, bootstrap, leave-one-out and worst-case results; Supplementary Tables S3–S5"],
        ["Results — first paragraph",
         "Analysis set description and reference to Supplementary Table S1"],
        ["Discussion — first and third paragraphs",
         "The three participants without a week-4 assessment; what the sham-versus-control comparison "
         "can and cannot support"],
        ["Discussion — limitations paragraph",
         "Allocation concealment, complete-case accounting, and the worst-case analysis added"],
        ["Financial Disclosures — Data availability, Code availability, Trial registration",
         "Repository and script availability; registration status stated explicitly"],
    ]
    d.append(d.build_table(tpl_table, ["Location", "Change"], rows, mode="clean",
                           author=AUTHOR, date=DATE, widths=[3.4, 6.6]))
    add("")

    add("Request to the editorial office regarding the attached first-round report", style="Heading2")
    add("Reviewer 2’s comment refers to a detailed report attached to the original submission. The "
        "corresponding entry (“Review supporting file”, reference 1308143) is visible in the review "
        "system, but the attachment was not delivered to us and could not be opened from our side. The "
        "five areas listed in the comment are answered in full above. If the attached report contains "
        "further specific requests, we would be grateful if the editorial office could forward it to us; "
        "we will answer every item in it, and can do so quickly.")

    out = ROOT / "01_稿件" / "修订记录" / "Response_to_Reviewers.docx"
    d.save(out)
    return out


# ------------------------------------------------------------------ 待补清单

def build_todo(repo_public: bool) -> Path:
    d = Docx(SRC)
    d.clear_body()
    src = Docx(SRC)
    body_tpl = pick_body_template(src)
    tpl_table = src.tables()[0]

    def add(text, style=None):
        p = d._new_paragraph(text, style, "clean", AUTHOR, DATE, template=body_tpl)
        d.append(p)
        return p

    add("待作者补信息清单（Frontiers in Psychology 修回，稿件号 1308143）", style="Heading1")
    add("下列信息无法从数据或原始记录中确定，须由作者填写后替换稿件中的占位表述。每条均注明在稿件中的位置。")
    rows = [
        ["1", "伦理批件号", "Methods — Participants；Financial Disclosures",
         "机构：新乡医学院第二附属医院（河南省精神病医院）伦理委员会。需填批件号与批准日期。"],
        ["2", "作者名单、学位、通讯作者及联系方式", "稿件首页",
         "稿件现为“see Author Completion Checklist”占位。"],
        ["3", "基金资助与 CRediT 作者贡献", "Financial Disclosures",
         "现为占位；需按期刊要求分角色填写。"],
        ["4", "回顾性试验注册号与注册日期", "Methods — Outcome hierarchy…；Financial Disclosures",
         "稿件已声明未前瞻注册；需补注册标识与注册日期，并同步到投稿系统。"],
        ["5", "分配隐藏机制", "Methods — Randomization and allocation",
         "若确无记录，请确认按现稿表述“未记录且无法重建”定稿；若有随机号表执行细节，请补。"],
        ["6", "电极片尺寸与材料、导电介质", "Methods — Procedures",
         "现稿只写“电极膏导电、10-20 命名位点”。需按设备与耗材记录补具体型号/直径/材料。"],
        ["7", "每次治疗的接触与阻抗检查、会话依从性核对方式", "Methods — Procedures",
         "现为占位。需说明 14 次完成情况如何逐人核对。"],
        ["8", "住院常规治疗的具体内容", "Methods — Routine inpatient care（新增小节）",
         "已写苯二氮䓬类药物替代治疗与常规支持治疗；需补：具体药名与剂量方案、是否有心理干预、"
         "是否有其他非药物干预、住院期间是否有其他影响认知的药物。"],
        ["9", "不良事件是否有可提取的结构化记录", "Methods — Safety monitoring；Results — Safety and tolerability",
         "现稿如实写明未预设结构化采集。需确认病历回顾能否补充具体事件与发生率；不能则维持现表述。"],
        ["10", "公开仓库与永久标识", "Financial Disclosures — Data / Code availability",
         "GitHub 仓库 Cybing521/hd-tdcs-aud-cognition 现为 private。需确认是否转公开，并补 Zenodo "
         "等永久 DOI。回复信与稿件中的两段文字随此结论切换。"],
        ["11", "Reviewer 2 首轮附件的逐条回复", "回复信最后一节",
         "该附件在审稿系统内以 “Review supporting file” 存在但当前无法打开；取得后需逐条补写。"],
    ]
    d.append(d.build_table(tpl_table, ["#", "待补项", "稿件位置", "说明与建议"], rows,
                           mode="clean", author=AUTHOR, date=DATE,
                           widths=[0.4, 2.0, 2.8, 4.8]))
    add("")
    add("说明：第 10 项影响 Financial Disclosures 两段的写法（公开仓库版 / 仅按需提供版两套文本已备好，"
        "按确认结果切换即可）；第 11 项影响回复信完整性。除以上各条外，本轮修订所引用的全部统计量"
        "均已由脚本从去标识数据重算，并已与上一版结果表逐项核对一致（109/109）。")

    out = ROOT / "01_稿件" / "修订记录" / "待作者补清单.docx"
    d.save(out)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo-public", dest="repo_public", action="store_true", default=True)
    ap.add_argument("--repo-private", dest="repo_public", action="store_false")
    args = ap.parse_args()

    outs = []
    for mode, (name, _) in MODES.items():
        outs.append(build_manuscript(mode, name, args.repo_public))
    outs.append(build_supplementary(args.repo_public))
    outs.append(build_response(args.repo_public))
    outs.append(build_todo(args.repo_public))
    for o in outs:
        print("written:", o)


if __name__ == "__main__":
    main()
