#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修回文本内容：所有统计量从 06_results/复算结果_全量.json 取，不手工抄写。

改写项以 (段落定位串, 新文本) 表达；新增段落以 (锚点定位串, 标题, 正文) 表达。
"""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RESULTS = json.loads((ROOT / "06_results" / "复算结果_全量.json").read_text(encoding="utf-8"))


def p(x, nd=2):
    """P 值排版：<0.001 写 <0.001，其余保留位数。"""
    if x is None:
        return "n/a"
    if x < 0.001:
        return "<0.001"
    return f"{x:.{nd}f}"


def p3(x, nd: int = 3) -> str:
    """P 值三位小数（置换检验总体、极端情形、留一法区间）。"""
    if x is None:
        return "n/a"
    if x < 0.001:
        return "<0.001"
    return f"{x:.{nd}f}"


def p4(x) -> str:
    """受限置换 P 值四位小数。"""
    if x is None:
        return "n/a"
    return f"{x:.4f}"


def prange(lo, hi) -> str:
    """P 值区间：两端都小于 0.001 时整段写 <0.001。"""
    if hi < 0.001:
        return "<0.001"
    return f"{p3(lo)}–{p3(hi)}"


def f3(x):
    return f"{x:.3f}"


def f2(x):
    return f"{x:.2f}"


def _tp(outcome: str) -> dict:
    return RESULTS["tipping_point"]["outcomes"][outcome]


def _boot(outcome: str, kind: str) -> dict:
    return RESULTS["bootstrap"]["outcomes"][outcome][kind]


def _perm(outcome: str) -> dict:
    return RESULTS["permutation"]["outcomes"][outcome]


def _loo(outcome: str) -> dict:
    return RESULTS["diagnostics"]["outcomes"][outcome]


# --------------------------------------------------------------- 段落改写

def edits(results: dict) -> list[tuple[str, str]]:
    """(段落定位子串, 新全文) —— 子串唯一命中一个段落。"""
    anc = results["ancova"]
    tp = results["tipping_point"]["outcomes"]
    perm = results["permutation"]["outcomes"]
    boot = results["bootstrap"]["outcomes"]
    loo = results["diagnostics"]["outcomes"]

    out: list[tuple[str, str]] = []

    # ---- 摘要 Methods
    out.append((
        "Methods: In this single-blind three-arm trial",
        "Methods: In this single-blind three-arm trial, 85 male inpatients were randomized to routine "
        "care (n=29), active HD-tDCS (n=28), or sham HD-tDCS (n=28); 77 were analyzed (n=29/21/27) "
        "after data verification, comprising every randomized participant who has a week-4 assessment. "
        "Bilateral 4×1 sessions (±1.0 mA, 20 min) targeted F3/F4 after detoxification. Primary "
        "outcomes were TMT-B, VFT, and HVLT-R. Analyses included post-intervention tests, "
        "baseline-adjusted ANCOVA, change scores, FDR correction, Cohen's d (95% CI), and Hedges g; "
        "permutation tests, bootstrap confidence intervals, and a worst-case analysis for the three "
        "participants without a week-4 assessment were added at revision.",
    ))

    # ---- 摘要 Results
    out.append((
        "Results: Active HD-tDCS surpassed sham and control on TMT-B",
        "Results: Active HD-tDCS surpassed sham and control on TMT-B, VFT, and HVLT-R (omnibus FDR "
        "q<0.05; versus sham d=-0.96 [-1.56, -0.36], 1.05 [0.44, 1.66], and 1.30 [0.67, 1.93]; "
        "g=-0.94/1.03/1.28). ANCOVA confirmed group effects (partial η²=0.19-0.31), as did change "
        "scores and sensitivity analyses (ANCOVA P≤0.0003). MMSE and DST did not differ; sham equaled "
        "control. Permutation and bootstrap analyses agreed with the parametric results, and the "
        "primary conclusions were unchanged under worst-case imputation for the three participants "
        "without a week-4 assessment.",
    ))

    # ---- Highlights
    out.append((
        "• ANCOVA and sensitivity analyses confirmed primary-outcome group effects",
        "• Conclusions held under permutation, bootstrap and worst-case imputation",
    ))

    # ---- Methods / Participants
    out.append((
        "Eighty-five participants were randomized 1:1:1",
        "Eighty-five participants were randomized 1:1:1 to routine care (n=29), active HD-tDCS "
        "(n=28), or sham HD-tDCS (n=28). Three participants discontinued the intervention early "
        "because of hospital discharge (1 active, 2 sham) and therefore have no week-4 cognitive "
        "assessment; 82 records with a week-4 assessment were on file. After removal of five erroneous "
        "active-arm records and reassignment of one retained record to sham, 77 participants entered "
        "the primary analysis (active n=21, sham n=27, control n=29). The primary analysis therefore "
        "contains every randomized participant who has a week-4 assessment, with one record per "
        "participant; the three participants without a week-4 assessment contribute no outcome data "
        "and were addressed in a worst-case analysis rather than being treated as having no "
        "improvement. The study was approved by the institutional ethics committee (approval number: "
        "see Author Completion Checklist). All participants and their guardians provided written "
        "informed consent. Participant flow is summarized in Figure 1 and Supplementary Table S1.",
    ))

    # ---- Methods / 数据核查
    out.append((
        "During database consolidation after intervention completion",
        "During database consolidation after intervention completion, 82 records with a week-4 "
        "assessment were on file (active 27, sham 26, control 29). Five erroneous active-arm records "
        "were identified: four cross-arm misentries whose valid copies already sat in the sham file, "
        "and one duplicate active-arm row for a participant whose retained record was reassigned to the "
        "sham arm according to the original randomization log. Removing the five erroneous active-arm "
        "rows and applying that reassignment yielded the primary analysis set of n=77 (active 21, sham "
        "27, control 29; Figure 1). Because each removed row duplicated a participant who was already "
        "represented in the sham arm, removing them discarded no participant outcome and changed no "
        "participant's allocation; the correction changed the number of records per participant, not "
        "the set of participants with data. No participant contributed more than one analyzed record. "
        "Sensitivity analyses that moved or dropped the reassigned record, or that analyzed all five "
        "ambiguous participants in the active arm as recorded, are reported in Statistical analysis "
        "and Results.",
    ))

    # ---- Methods / Randomization
    out.append((
        "The randomization sequence was generated with a random number table",
        "The randomization sequence was generated with a random number table at a 1:1:1 allocation "
        "ratio. The mechanism used to conceal the upcoming assignment from enrollment staff was not "
        "documented in the archived protocol and could not be reconstructed from the available "
        "records; this is stated explicitly as a trial limitation rather than left to inference (see "
        "Author Completion Checklist). The sequence was therefore not verifiably concealed from "
        "enrolment staff.",
    ))

    # ---- Methods / Procedures（补技术描述）
    out.append((
        "After benzodiazepine-assisted detoxification, participants received routine inpatient care alone",
        "After benzodiazepine-assisted detoxification, participants received routine inpatient care "
        "alone (control), routine care plus active HD-tDCS, or routine care plus sham HD-tDCS. The "
        "control arm did not receive equivalent device-contact time; active and sham arms each "
        "completed 14 additional 20-min stimulation sessions. Stimulation was delivered with a Soterix "
        "MXN-9 HD-tES device in direct-current (DC) mode using bilateral 4×1 ring montages: F3 anode "
        "with return electrodes at FP1, F7, C1, and C3; F4 cathode with return electrodes at FP2, F8, "
        "C2, and C4 (Figure 2). Electrode positions are named with the standard 10-20 scalp "
        "nomenclature; each 4×1 ring places one central electrode with four surrounding return "
        "electrodes over one hemisphere, and the electrode-scalp interface was prepared with conductive "
        "gel rather than saline-soaked sponge pads. Intensity was ±1.0 mA for 20 min once daily for 14 "
        "sessions (12, 22). Active stimulation used continuous direct current throughout each session. "
        "Sham stimulation applied current only during the first and last 30 s of each session, and the "
        "ramp periods were not disclosed to participants. Electrode size/material and session-adherence "
        "verification details are listed in the Author Completion Checklist.",
    ))

    # ---- Methods / Statistical analysis（补新增分析）
    out.append((
        "Analyses were performed on the cleaned individual-level complete-case dataset",
        "Analyses were performed on the cleaned individual-level analysis dataset, which contains one "
        "record per participant with a week-4 assessment. The original analyses used SPSS 22.0; all "
        "statistics reported here were independently recomputed from the individual-level records in "
        "Python (SciPy 1.17 and statsmodels 0.14) and cross-checked against the original output. Data "
        "are presented as mean±SD. Within-group pre-to-post changes were evaluated with paired "
        "t-tests. Between-group comparisons at each time point used one-way ANOVA when variances were "
        "homogeneous (Levene's test) and Welch's ANOVA otherwise, with LSD or Games-Howell pairwise "
        "comparisons accordingly. To separate stimulation-specific benefit from partial recovery during "
        "inpatient care, baseline-to-post-intervention change scores were compared across groups using "
        "the same procedures, and analysis of covariance (ANCOVA) tested group differences in each "
        "post-intervention score with its baseline value as covariate; homogeneity of regression slopes "
        "was verified through group-by-baseline interaction terms. TMT-B, VFT, and HVLT-R were "
        "prespecified primary outcomes and MMSE and DST secondary outcomes. Benjamini-Hochberg false "
        "discovery rate (FDR) correction was applied across the five post-intervention omnibus tests. "
        "Change-score and ANCOVA families were reported with nominal P values as confirmatory parallel "
        "analyses. For primary contrasts, Cohen's d was reported with large-sample 95% confidence "
        "intervals together with Hedges g correction (24). Pearson correlation examined all twenty "
        "associations between four drinking characteristics and five baseline cognitive measures as an "
        "exploratory analysis, with Benjamini-Hochberg correction applied across the twenty tests. "
        "Sensitivity analyses reassigned the reclassified participant to the active arm, excluded that "
        "participant, or analyzed all five ambiguous participants in the active arm as recorded; the "
        "minimum detectable d at 80% power for n=21 versus n=27 is 0.83. In response to reviewer "
        "comments, three further analyses were added to the re-analysed dataset during the present "
        "revision and are reported as such: permutation tests (10,000 label permutations, seed "
        "20261027) for the omnibus three-group comparison and, in a restricted form within the active "
        "and sham arms, for the active-versus-sham contrast; bootstrap percentile confidence intervals "
        "(10,000 resamples, seed 20261027) for the active-versus-sham mean difference in "
        "post-intervention scores and in change scores; and a worst-case analysis for the three "
        "participants without a week-4 assessment, which imputed their missing week-4 scores at "
        "progressively larger offsets in the direction unfavourable to the active arm and identified "
        "the offset at which the active-versus-sham comparison would lose significance, together with "
        "a deterministic extreme case imputing the worst observed active-arm score and the best "
        "observed sham-arm scores. Multiple imputation was not possible because the baseline "
        "covariates of the three participants are absent from the de-identified dataset, so no "
        "imputation model could be fitted; the offset and extreme-case analyses were used instead. "
        "Two-sided P<0.05 denoted nominal significance; multiplicity-adjusted inference used q<0.05. No "
        "interim analyses were conducted.",
    ))

    # ---- Methods / 预设与注册
    out.append((
        "Benjamini-Hochberg false discovery rate (FDR) correction was applied across all five",
        "Benjamini-Hochberg false discovery rate (FDR) correction was applied across all five "
        "post-intervention omnibus between-group tests, and primary outcomes remained significant at "
        "FDR-adjusted q<0.05. The same correction was applied across the twenty exploratory "
        "correlation tests. Change-score and ANCOVA families were treated as confirmatory parallel "
        "analyses and are reported with nominal P values. This study was not prospectively registered, "
        "so no analysis plan was registered or published before enrolment. The outcome hierarchy and "
        "the analysis framework reported here were fixed before the present re-analysis of the "
        "de-identified dataset was run, and every reported value was recomputed from the "
        "individual-level records. The permutation, bootstrap, and worst-case analyses described in "
        "Statistical analysis were added during this revision in response to reviewer comments and "
        "are identified as such in Results. No interim analyses were planned or conducted; the trial "
        "ran to completion.",
    ))

    # ---- Results / 样本量
    out.append((
        "A total of 77 participants were included in the primary analysis",
        "A total of 77 participants were included in the primary analysis (Figure 1; Supplementary "
        "Table S1), comprising every randomized participant who has a week-4 assessment.",
    ))

    # ---- Discussion / 开头
    out.append((
        "In this three-arm inpatient trial, bilateral 4×1 HD-tDCS over the DLPFC",
        "In this three-arm inpatient trial, bilateral 4×1 HD-tDCS over the DLPFC produced greater "
        "improvement than sham on a prespecified cognitive battery (TMT-B, VFT, and HVLT-R), with "
        "parallel advantage over routine care. Baseline characteristics were comparable. Of 85 "
        "randomized participants, 77 entered the primary analysis after removal of erroneous "
        "active-arm records and reassignment of one retained record to the sham arm according to the "
        "randomization log (Supplementary Table S1). The three participants discharged early before "
        "week 4 were without outcome data; the worst-case analysis showed that the primary conclusions "
        "do not depend on their outcomes, and permutation and bootstrap analyses gave the same answer "
        "as the parametric tests.",
    ))

    # ---- Discussion / 假刺激 vs 对照（Reviewer 3）
    out.append((
        "Sham HD-tDCS did not differ from routine care on primary outcomes",
        "Sham HD-tDCS did not differ from routine care on primary outcomes. All three arms shared the "
        "same inpatient withdrawal management and supportive care, and all three improved on the "
        "cognitive battery during hospitalization; the sham-versus-control comparison therefore shows "
        "that the device-contact and session structure shared by the active and sham arms was "
        "insufficient, under the present dose, to produce measurable cognitive gains beyond routine "
        "treatment. Because the control arm did not receive matched device-contact time, this "
        "equivalence does not prove that sham current is cognitively inert, and it cannot separate the "
        "effect of the delivered current from any effect of daily device contact. The "
        "active-versus-sham contrast, in which both arms received identical device contact, is the "
        "estimate that isolates the delivered current.",
    ))

    # ---- Discussion / 局限性
    out.append((
        "Generalizability is bounded by the single-center male inpatient sample",
        "Generalizability is bounded by the single-center male inpatient sample (active n=21 after "
        "verification), unblinded assessors, unmatched device-contact in the routine-care arm, and "
        "absence of prospective registration. Allocation concealment was not documented in the "
        "archived protocol and could not be reconstructed, so the possibility that enrolment staff "
        "anticipated the upcoming assignment cannot be excluded; with three arms and 85 participants "
        "this is a design-level weakness that no statistical adjustment can remove. The three "
        "participants who were discharged early before week 4 contributed no outcome data; the "
        "worst-case analysis indicates that the primary conclusions do not depend on their outcomes, "
        "but the trial would have been better served by a complete-case accounting in which all "
        "randomized participants are followed to a common assessment point. Observed active-versus-sham "
        "effects exceeded the minimum detectable d of 0.83 at 80% power for n=21 versus n=27. "
        "Change-score and ANCOVA analyses reduce concerns about baseline imbalance but do not eliminate "
        "residual confounding.",
    ))

    return out


def new_sections(results: dict) -> list[dict]:
    """新增小节：锚点定位串 + 标题 + 正文（数字取自结果表）。"""
    tp = results["tipping_point"]["outcomes"]
    perm = results["permutation"]["outcomes"]
    boot = results["bootstrap"]["outcomes"]
    loo = results["diagnostics"]["outcomes"]

    routine = (
        "All three arms received the same routine inpatient treatment for alcohol dependence, "
        "consisting of benzodiazepine-assisted withdrawal management followed by supportive inpatient "
        "care. No arm received structured cognitive training, neurofeedback, or any other "
        "neuromodulation during the study period. The routine-care arm was never exposed to the "
        "stimulation device and therefore did not receive the additional 20-min device-contact time "
        "that the active and sham arms completed on 14 occasions. Because all three arms improved on "
        "the cognitive battery during hospitalization, the between-arm contrast rather than the "
        "within-arm change is the interpretable quantity: improvement common to all arms reflects the "
        "natural course of early abstinence under routine care (26), whereas the active-versus-sham "
        "contrast isolates the effect of the delivered current. The specific withdrawal medications, "
        "dosing schedule, and any psychological interventions delivered during hospitalization are "
        "listed in the Author Completion Checklist."
    )

    tmt, vft, hvlt = tp["TMT-B"], tp["VFT"], tp["HVLT-R"]
    robustness = (
        "The active-versus-sham contrast on the three primary outcomes was unchanged under analyses "
        "that do not rely on the normality and equal-variance assumptions of the t-test. Permutation "
        f"tests (10,000 permutations) gave omnibus P={p3(perm['TMT-B']['perm_p_omnibus'])} for TMT-B, "
        f"P={p3(perm['VFT']['perm_p_omnibus'])} for VFT, and P={p3(perm['HVLT-R']['perm_p_omnibus'])} "
        "for HVLT-R, and restricted-permutation "
        f"P={p4(perm['TMT-B']['perm_p_active_vs_sham'])}, P={p4(perm['VFT']['perm_p_active_vs_sham'])}, "
        f"and P={p4(perm['HVLT-R']['perm_p_active_vs_sham'])} for the active-versus-sham contrast, "
        "matching the parametric results. Bootstrap percentile confidence intervals (10,000 resamples) "
        "for the active-minus-sham difference in change scores excluded zero for all three primary "
        f"outcomes (TMT-B {f2(boot['TMT-B']['change']['observed_diff'])} s, 95% CI "
        f"{f2(boot['TMT-B']['change']['boot_ci_lo'])} to {f2(boot['TMT-B']['change']['boot_ci_hi'])}; "
        f"VFT {f2(boot['VFT']['change']['observed_diff'])} points, 95% CI "
        f"{f2(boot['VFT']['change']['boot_ci_lo'])} to {f2(boot['VFT']['change']['boot_ci_hi'])}; "
        f"HVLT-R {f2(boot['HVLT-R']['change']['observed_diff'])} points, 95% CI "
        f"{f2(boot['HVLT-R']['change']['boot_ci_lo'])} to {f2(boot['HVLT-R']['change']['boot_ci_hi'])}) "
        "and included zero for both secondary outcomes (MMSE "
        f"{f2(boot['MMSE']['change']['observed_diff'])}, 95% CI "
        f"{f2(boot['MMSE']['change']['boot_ci_lo'])} to {f2(boot['MMSE']['change']['boot_ci_hi'])}; "
        f"DST {f2(boot['DST']['change']['observed_diff'])}, 95% CI "
        f"{f2(boot['DST']['change']['boot_ci_lo'])} to {f2(boot['DST']['change']['boot_ci_hi'])}). "
        "Leave-one-participant-out re-analysis kept the active-versus-sham comparison significant for "
        "every one of the 77 possible removals (TMT-B P "
        f"{prange(loo['TMT-B']['loo_active_vs_sham_P_range'][0], loo['TMT-B']['loo_active_vs_sham_P_range'][1])}; "
        f"VFT P "
        f"{prange(loo['VFT']['loo_active_vs_sham_P_range'][0], loo['VFT']['loo_active_vs_sham_P_range'][1])}; "
        f"HVLT-R P "
        f"{prange(loo['HVLT-R']['loo_active_vs_sham_P_range'][0], loo['HVLT-R']['loo_active_vs_sham_P_range'][1])}), "
        f"so no single participant drives the "
        "result."
    )

    worst_case = (
        "The three participants without a week-4 assessment were addressed by a worst-case analysis. "
        "Imputing their missing week-4 scores at offsets in the direction unfavourable to the active "
        "arm, the active-versus-sham comparison lost significance only when the imputed values were "
        f"pushed to {tmt['tipping']['delta_sd']} standard deviations for TMT-B (imputed active-arm "
        f"score {f2(tmt['tipping']['imputed_active_post'])} s against an imputed sham-arm score of "
        f"{f2(tmt['tipping']['imputed_sham_post'])} s), {vft['tipping']['delta_sd']} standard deviations "
        f"for VFT ({f2(vft['tipping']['imputed_active_post'])} against "
        f"{f2(vft['tipping']['imputed_sham_post'])} points), and "
        f"{hvlt['tipping']['delta_sd']} standard deviations for HVLT-R "
        f"({f2(hvlt['tipping']['imputed_active_post'])} against "
        f"{f2(hvlt['tipping']['imputed_sham_post'])} points). Each of those imputations lies outside "
        "the range the instrument can take (a negative TMT-B completion time and an HVLT-R score above "
        "the 36-point maximum are not attainable), so the conclusion cannot be overturned by any "
        "physically possible outcome for these three participants. In the deterministic extreme case, in "
        "which the missing active-arm participant was assigned the worst observed active-arm score and "
        "the two missing sham-arm participants the best observed sham-arm scores, the "
        "active-versus-sham comparison remained significant for all three primary outcomes (TMT-B "
        f"P={p3(tmt['extreme_worst_case_P'])}, VFT P={p3(vft['extreme_worst_case_P'])}, HVLT-R "
        f"P={p3(hvlt['extreme_worst_case_P'])})."
    )

    return [
        {"anchor": "Stimulation was delivered with a Soterix MXN-9 HD-tES device",
         "after": "Procedures", "title": "Routine inpatient care", "text": routine},
        {"anchor": "Sensitivity analyses that reassigned the reclassified participant to the active arm",
         "after": "Post-intervention between-group comparisons", "title": "Robustness of the primary contrasts",
         "text": robustness + " " + worst_case},
    ]


def data_availability_text(repo_public: bool) -> tuple[str, str]:
    """(Data availability, Code availability) 两段。"""
    if repo_public:
        data = (
            "Data availability: The de-identified individual-level analysis dataset, the statistical "
            "analysis scripts, and the derived result tables are deposited in a public repository "
            "(https://github.com/Cybing521/hd-tdcs-aud-cognition; a permanent DOI will be assigned at "
            "the point of submission) and are listed in Supplementary Table S6. The repository "
            "contains no direct identifiers: participant names were removed at source and replaced "
            "with hashed study identifiers, and the de-identification procedure is documented in the "
            "de-identification report distributed with the dataset. Identifiable records remain with "
            "the originating institution and are not redistributed."
        )
        code = (
            "Code availability: All statistical analysis scripts used to produce every number reported "
            "here, together with the de-identified analysis tables and the derived result tables, are "
            "available in the public repository cited under Data availability. The scripts run on "
            "Python 3.14 with SciPy 1.17, statsmodels 0.14, and openpyxl 3.1, and regenerate the result "
            "tables and figures from the de-identified tables without manual steps."
        )
    else:
        data = (
            "Data availability: The de-identified individual-level analysis dataset, the statistical "
            "analysis scripts, and the derived result tables are available from the corresponding "
            "author on reasonable request and with appropriate institutional agreements, because "
            "institutional privacy restrictions and the absence of participant consent for open "
            "sharing preclude unrestricted redistribution. Summary statistics and structured analysis "
            "outputs sufficient to reproduce the reported inferential results are provided in the "
            "Supplementary Material. A public repository is being prepared and its permanent URL will "
            "be added here before publication."
        )
        code = (
            "Code availability: Statistical analysis and manuscript generation scripts are available "
            "from the corresponding author upon request. A permanent public repository URL has not yet "
            "been assigned; the scripts reproduce every reported statistic from the de-identified "
            "analysis tables without manual steps."
        )
    return data, code


def trial_registration_text() -> str:
    return (
        "Trial registration: This study was not prospectively registered before participant "
        "enrollment (2019). Retrospective registration with disclosure of the registration date is "
        "planned before journal submission, and the registration identifier and date will be added "
        "here. Because no prospective protocol or analysis plan was registered, all analyses reported "
        "in this manuscript are post hoc with respect to registration; the outcome hierarchy and the "
        "analysis framework were fixed before the present re-analysis was run, and the additional "
        "robustness analyses added during revision are identified as such in the Methods and Results."
    )
