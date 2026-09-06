# -*- coding: utf-8 -*-
"""Journal-aligned manuscript text and references (symbolic citation system).

Citations are written as [@key,...] tokens. At build time,
``assign_reference_numbers`` numbers keys by first appearance in document
order (Vancouver), ``render_citations`` substitutes condensed numeric
citations (e.g. [18-24]) and ``reference_list`` produces the ordered list.
All bibliographic metadata below was verified against Crossref on 2026-08-24
(see ccfa-review-reports/ref_validation_2026-08-24.md).
"""

from __future__ import annotations

import re

# ---------------------------------------------------------------------------
# Reference database (key -> formatted Vancouver string)
# ---------------------------------------------------------------------------

REF_DB: dict[str, str] = {
    "stavro2013": (
        "Stavro K, Pelletier J, Potvin S. Widespread and sustained cognitive deficits in "
        "alcoholism: a meta-analysis. Addict Biol. 2013;18(2):203-213. "
        "doi:10.1111/j.1369-1600.2011.00418.x."
    ),
    "ellingson2014": (
        "Ellingson JM, Fleming KA, Vergés A, et al. Working memory as a moderator of impulsivity "
        "and alcohol involvement. Addict Behav. 2014;39(11):1622-1631. "
        "doi:10.1016/j.addbeh.2014.01.004."
    ),
    "christiansen2014": (
        "Christiansen P, Bloor JF. Individualised but not general alcohol Stroop predicts "
        "alcohol use. Drug Alcohol Depend. 2014;134:410-413. "
        "doi:10.1016/j.drugalcdep.2013.10.021."
    ),
    "weiland2014": (
        "Weiland BJ, Sabbineni A, Calhoun VD, et al. Reduced left executive control network "
        "functional connectivity is associated with alcohol use disorders. Alcohol Clin Exp Res. "
        "2014;38(9):2445-2453. doi:10.1111/acer.12505."
    ),
    "bernardin2014": (
        "Bernardin F, Maheut-Bosser A, Paille F. Cognitive impairments in alcohol-dependent "
        "subjects. Front Psychiatry. 2014;5:78. doi:10.3389/fpsyt.2014.00078."
    ),
    "chenj2020": (
        "陈静, 钱敏才, 林敏, 等. 酒依赖脱瘾维持期患者复饮相关因素分析. 临床精神医学杂志. 2020;30(3):189-191."
    ),
    "boggio2008": (
        "Boggio PS, Sultani N, Fecteau S, et al. Prefrontal cortex modulation using transcranial "
        "DC stimulation reduces alcohol craving: a double-blind, sham-controlled study. Drug "
        "Alcohol Depend. 2008;92(1-3):55-60. doi:10.1016/j.drugalcdep.2007.06.011."
    ),
    "stagg2011": (
        "Stagg CJ, Nitsche MA. Physiological basis of transcranial direct current stimulation. "
        "Neuroscientist. 2011;17(1):37-53. doi:10.1177/1073858410386614."
    ),
    "park2013": (
        "Park CH, Chang WH, Park JY, et al. Transcranial direct current stimulation increases "
        "resting state interhemispheric connectivity. Neurosci Lett. 2013;539:7-10. "
        "doi:10.1016/j.neulet.2013.01.047."
    ),
    "fregni2021": (
        "Fregni F, El-Hagrassy MM, Pacheco-Barrios K, et al. Evidence-based guidelines and "
        "secondary meta-analysis for the use of transcranial direct current stimulation in "
        "neurological and psychiatric disorders. Int J Neuropsychopharmacol. 2021;24(4):256-313. "
        "doi:10.1093/ijnp/pyaa051."
    ),
    "denuyl2015": (
        "den Uyl TE, Gladwin TE, Wiers RW. Transcranial direct current stimulation, implicit "
        "alcohol associations and craving. Biol Psychol. 2015;105:37-42. "
        "doi:10.1016/j.biopsycho.2014.12.004."
    ),
    "klauss2014": (
        "Klauss J, Penido Pinheiro LC, Silva Merlo BL, et al. A randomized controlled trial of "
        "targeted prefrontal cortex modulation with tDCS in patients with alcohol dependence. "
        "Int J Neuropsychopharmacol. 2014;17(11):1793-1803. doi:10.1017/S1461145714000984."
    ),
    "klauss2018": (
        "Klauss J, Anders QS, Felippe LV, et al. Multiple sessions of transcranial direct "
        "current stimulation reduced craving and relapses for alcohol use: a randomized "
        "placebo-controlled trial in alcohol use disorder. Front Pharmacol. 2018;9:716. "
        "doi:10.3389/fphar.2018.00716."
    ),
    "dubuson2021": (
        "Dubuson M, Kornreich C, Vanderhasselt MA, et al. Transcranial direct current "
        "stimulation combined with alcohol cue inhibitory control training reduces the risk of "
        "early alcohol relapse: a randomized placebo-controlled clinical trial. Brain Stimul. "
        "2021;14(6):1531-1543. doi:10.1016/j.brs.2021.10.386."
    ),
    "camchong2023": (
        "Camchong J, Roediger D, Fiecas M, et al. Frontal tDCS reduces alcohol relapse rates by "
        "increasing connections from left dorsolateral prefrontal cortex to addiction networks. "
        "Brain Stimul. 2023;16(4):1032-1040. doi:10.1016/j.brs.2023.06.011."
    ),
    "bollen2022": (
        "Bollen Z, Dormal G, Maurage P. How should transcranial direct current stimulation be "
        "used in populations with severe alcohol use disorder? A clinically oriented systematic "
        "review. Clin EEG Neurosci. 2022;53(4):287-301. doi:10.1177/15500594211001212."
    ),
    "zhang2024": (
        "Zhang X, Huang M, Yu Y, et al. Is transcranial direct current stimulation effective for "
        "cognitive dysfunction in substance use disorders? A systematic review. Brain Sci. "
        "2024;14(8):754. doi:10.3390/brainsci14080754."
    ),
    "dayal2023": (
        "Dayal P, Kaloiya G, Verma R, et al. Effects of transcranial direct-current stimulation "
        "(tDCS) treatment sessions on neuropsychological functioning among detoxified inpatients "
        "with alcohol dependence: a randomized sham-controlled clinical trial. Alcohol Treat Q. "
        "2023;41(4):474-487. doi:10.1080/07347324.2023.2243237."
    ),
    "dasilva2013": (
        "da Silva MC, Conti CL, Klauss J, et al. Behavioral effects of transcranial direct "
        "current stimulation induced dorsolateral prefrontal cortex plasticity in alcohol "
        "dependence. J Physiol Paris. 2013;107(6):493-502. doi:10.1016/j.jphysparis.2013.07.003."
    ),
    "nakamura2012": (
        "Nakamura-Palacios EM, de Almeida Benevides MC, da Penha Zago-Gomes M, et al. Auditory "
        "event-related potentials (P3) and cognitive changes induced by frontal direct current "
        "stimulation in alcoholics according to Lesch alcoholism typology. Int J "
        "Neuropsychopharmacol. 2012;15(5):601-616. doi:10.1017/S1461145711001040."
    ),
    "kuo2013": (
        "Kuo HI, Bikson M, Datta A, et al. Comparing cortical plasticity induced by conventional "
        "and high-definition 4×1 ring tDCS: a neurophysiological study. Brain Stimul. "
        "2013;6(4):644-648. doi:10.1016/j.brs.2013.04.006."
    ),
    "kumar2026": (
        "Kumar S, Khanra S, Roy C, Munda SK. Effect of adjunctive high-definition transcranial "
        "direct current stimulation (HD-tDCS) to oral naltrexone on craving, brain "
        "glutamate-glutamine (Glx) at left dorsolateral prefrontal cortex (DLPFC) in alcohol use "
        "disorder: a randomized sham-controlled HD-tDCS-fMRS study. J ECT. 2026. "
        "doi:10.1097/YCT.0000000000001257."
    ),
    "zhangmy2015": (
        "张明园, 何燕玲. 精神科评定量表手册. 长沙: 湖南科学技术出版社; 2015:280-283."
    ),
    "cohen1988": (
        "Cohen J. Statistical power analysis for the behavioral sciences. 2nd ed. Hillsdale "
        "(NJ): Lawrence Erlbaum Associates; 1988."
    ),
    "brunoni2011": (
        "Brunoni AR, Amadera J, Berbel B, et al. A systematic review on reporting and assessment "
        "of adverse effects associated with transcranial direct current stimulation. Int J "
        "Neuropsychopharmacol. 2011;14(8):1133-1145. doi:10.1017/S1461145710001690."
    ),
    "wangsj2020": (
        "王淑静, 王传升, 张慧, 等. 酒依赖患者规范化治疗后认知功能改善分析. 精神医学杂志. 2020;33(3):214-216."
    )
}

# ---------------------------------------------------------------------------
# Citation resolution helpers
# ---------------------------------------------------------------------------

_CITE_RE = re.compile(r"\[@([^\]]+)\]")


def _iter_keys(text: str):
    for m in _CITE_RE.finditer(text):
        for tok in m.group(1).split(","):
            tok = tok.strip().lstrip("@")
            if tok:
                yield tok


def assign_reference_numbers(texts) -> dict[str, int]:
    """Number citation keys by first appearance across `texts` (in order)."""
    order: dict[str, int] = {}
    for t in texts:
        for k in _iter_keys(t):
            if k not in order:
                if k not in REF_DB:
                    raise KeyError(f"Unknown reference key in text: {k}")
                order[k] = len(order) + 1
    return order


def unassigned_keys(nummap: dict[str, int]) -> set[str]:
    """Reference-database entries never cited anywhere in `nummap`."""
    return set(REF_DB) - set(nummap)


def _condense(nums) -> str:
    nums = sorted(set(nums))
    parts: list[str] = []
    i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        parts.append(str(nums[i]) if i == j else f"{nums[i]}-{nums[j]}")
        i = j + 1
    return ",".join(parts)


def render_citations(text: str, nummap: dict[str, int]) -> str:
    def repl(m):
        keys = [t.strip().lstrip("@") for t in m.group(1).split(",") if t.strip()]
        return "[" + _condense(nummap[k] for k in keys) + "]"

    return _CITE_RE.sub(repl, text)


def reference_list(nummap: dict[str, int]) -> list[str]:
    return [REF_DB[k] for k, _ in sorted(nummap.items(), key=lambda kv: kv[1])]


# ---------------------------------------------------------------------------
# Introduction (EN)
# ---------------------------------------------------------------------------

INTRO_EN = [
    (
        "Alcohol dependence is marked by impaired control over drinking and by enduring deficits in "
        "executive function, attention, and learning and memory "
        "[@stavro2013,@ellingson2014,@christiansen2014,@weiland2014,@bernardin2014]. These deficits "
        "can undermine treatment engagement and abstinence [@stavro2013,@chenj2020]. Because prefrontal "
        "circuits that support cognitive control are implicated in alcohol use disorder "
        "[@weiland2014,@dasilva2013,@camchong2023], noninvasive stimulation of the dorsolateral "
        "prefrontal cortex (DLPFC) is a candidate strategy for restoring frontally mediated cognition "
        "during early recovery."
    ),
    (
        "Transcranial direct current stimulation (tDCS) can modulate cortical excitability at low cost "
        "and with a favorable safety profile [@boggio2008,@stagg2011,@park2013,@fregni2021]. In alcohol "
        "use disorder, however, most randomized trials have prioritized craving, relapse, or cue "
        "reactivity "
        "[@denuyl2015,@klauss2014,@klauss2018,@dubuson2021,@camchong2023,@bollen2022,@kumar2026]. Cognitive "
        "endpoints remain heterogeneous across substance-use studies and are sensitive to montage, "
        "dose, and outcome choice [@zhang2024]. After ten sessions of conventional left-DLPFC "
        "sponge-electrode tDCS, Dayal et al. [@dayal2023] observed no tDCS-by-time interaction on "
        "neuropsychological performance in recently detoxified inpatients, whereas earlier DLPFC work "
        "reported behavioral, FAB, or cognitive improvement [@nakamura2012,@dasilva2013]."
    ),
    (
        "Methodologically, electrode montage determines how much current reaches the intended cortex. "
        "Conventional bipolar sponge electrodes (typically ~25–35 cm²) distribute current over broad "
        "frontocortical territories. High-definition tDCS (HD-tDCS) instead uses a 4×1 ring of small "
        "gel electrodes to constrain the electric field beneath the center electrode and increase "
        "focality over the DLPFC [@kuo2013]. In the present protocol, bilateral rings placed the anode "
        "at F3 (returns FP1, F7, C1, C3) and the cathode at F4 (returns FP2, F8, C2, C4), delivering "
        "±1.0 mA for 20 min once daily across 14 consecutive sessions (Figure 2). That combination—focal "
        "4×1 geometry, bilateral DLPFC targeting, and a two-week inpatient dose—defines the stimulation "
        "contribution tested here. Prior HD-tDCS work in alcohol use disorder has emphasized craving "
        "or metabolic readouts rather than a full standardized neuropsychological battery [@kumar2026]."
    ),
    (
        "We conducted a single-blind, three-arm randomized trial in male inpatients with alcohol "
        "dependence to test whether 14 daily sessions of bilateral 4×1 HD-tDCS over the DLPFC "
        "produce greater cognitive improvement than sham HD-tDCS or routine inpatient care alone. "
        "Prespecified primary outcomes were Trail Making Test Part B (TMT-B), verbal fluency (VFT), "
        "and Hopkins Verbal Learning Test-Revised (HVLT-R). The Mini-Mental State Examination (MMSE) "
        "and Digit Span Test (DST) were secondary outcomes."
    ),
]

# ---------------------------------------------------------------------------
# Discussion (EN)
# ---------------------------------------------------------------------------

DISCUSSION_EN = [
    (
        "In this three-arm inpatient trial, bilateral 4×1 HD-tDCS over the DLPFC produced greater "
        "improvement than sham on a prespecified cognitive battery (TMT-B, VFT, and HVLT-R), with "
        "parallel advantage over routine care. Baseline characteristics were comparable. Of 85 "
        "randomized participants, 77 entered the primary complete-case analysis after removal of "
        "erroneous active-arm records and reassignment of one retained record to the sham arm "
        "according to the randomization log (Supplementary Table S1)."
    ),
    (
        "The active-versus-sham contrast was large and consistent across analytic frameworks. "
        "Change-score, post-intervention, and baseline-adjusted ANCOVA models all showed group "
        "effects on the three primary outcomes (partial η²=0.19-0.31), and sensitivity analyses that "
        "altered assignment of ambiguous records preserved those ANCOVA conclusions. All arms "
        "improved to some degree during detoxification and supportive care [@wangsj2020]; the "
        "sham-controlled excess improvement therefore isolates stimulation-specific benefit from "
        "nonspecific inpatient recovery."
    ),
    (
        "Sham HD-tDCS did not differ from routine care on primary outcomes. Because the control arm "
        "did not receive matched device-contact time, this equivalence indicates that the added "
        "session structure shared by the active and sham arms was insufficient, under the present "
        "dose, to produce measurable cognitive gains beyond routine treatment—rather than proving "
        "that sham current is cognitively inert."
    ),
    (
        "Improvement concentrated on TMT-B, VFT, and HVLT-R, whereas MMSE and DST showed no group "
        "differences. We interpret this as a domain-selective behavioral profile under focal "
        "bilateral DLPFC HD dosing [@kuo2013,@dasilva2013,@camchong2023]. Near-ceiling baseline MMSE "
        "scores (mean 26-27 of 30) also left little room to detect further change on that secondary "
        "screen. Psychometric endpoints were chosen a priori; the trial was not designed to localize "
        "mechanism with imaging or spectroscopy."
    ),
    (
        "Neighboring HD-tDCS work in alcohol use disorder has focused on craving and prefrontal "
        "metabolites: adjunctive left-DLPFC HD-tDCS with naltrexone altered Glx without a "
        "between-group craving effect [@kumar2026]. Conventional sponge-electrode DLPFC tDCS trials "
        "have reported craving, relapse, P3/FAB screening, or mixed neuropsychological batteries "
        "[@nakamura2012,@dasilva2013,@dayal2023,@klauss2018,@bollen2022], but not a full "
        "sham-controlled standardized battery under bilateral 4×1 HD dosing. The present study "
        "therefore complements those endpoints by asking whether a focal, multi-session HD montage "
        "improves prespecified neuropsychological scores during inpatient detoxification."
    ),
    (
        "Exploratory correlations across twenty drinking-cognition pairs yielded two associations "
        "that survived Benjamini-Hochberg correction: longer cumulative alcohol-use duration "
        "(r=0.42, q=0.003) and longer dependence duration (r=0.33, q=0.037) tracked slower baseline "
        "TMT-B times. Cumulative exposure may therefore weigh particularly on set shifting; this "
        "pattern requires replication in larger cohorts."
    ),
    (
        "Generalizability is bounded by the single-center male inpatient sample (active n=21 after "
        "verification), unblinded assessors, unmatched device-contact in the routine-care arm, and "
        "absence of prospective registration. Observed active-versus-sham effects exceeded the "
        "minimum detectable d of 0.83 at 80% power for n=21 versus n=27. Change-score and ANCOVA "
        "analyses reduce concerns about baseline imbalance but do not eliminate residual confounding."
    ),
    (
        "Fourteen sessions of bilateral HD-tDCS over the DLPFC were associated with greater "
        "sham-controlled improvement on a standardized cognitive battery in male inpatients with "
        "alcohol dependence. Preregistered multicenter trials with matched attention controls, "
        "objective biomarkers, structured safety monitoring, and relapse follow-up will determine "
        "whether these cognitive gains translate into durable clinical benefit."
    ),
]

# ---------------------------------------------------------------------------
# Introduction / Discussion (ZH, kept in sync with EN numbering)
# ---------------------------------------------------------------------------

INTRO_ZH = [
    (
        "酒依赖常伴执行功能、注意及学习记忆等持续性认知损害"
        "[@stavro2013,@ellingson2014,@christiansen2014,@weiland2014,@bernardin2014]，并可能削弱治疗依从性与戒断维持"
        "[@stavro2013,@chenj2020]。鉴于支持认知控制的前额叶环路与酒精使用障碍相关"
        "[@weiland2014,@dasilva2013,@camchong2023]，对背外侧前额叶（DLPFC）的无创刺激是戒断早期改善额叶介导认知的候选策略。"
    ),
    (
        "经颅直流电刺激（tDCS）可在较低成本下调节皮层兴奋性，安全性较好[@boggio2008,@stagg2011,@park2013,@fregni2021]。"
        "但在酒精使用障碍中，多数随机试验仍以渴求、复饮或线索反应性为主"
        "[@denuyl2015,@klauss2014,@klauss2018,@dubuson2021,@camchong2023,@bollen2022,@kumar2026]。物质使用障碍中的认知终点异质性高，"
        "并受蒙太奇、剂量与结局选择影响[@zhang2024]。Dayal等[@dayal2023]在近期脱毒住院患者中采用10次传统左DLPFC海绵电极tDCS，"
        "未观察到刺激×时间交互；早期海绵电极DLPFC研究则报道行为、FAB或认知改善[@nakamura2012,@dasilva2013]。"
    ),
    (
        "方法学上，电极布局决定电流到达目标皮层的聚焦程度。传统双极海绵电极电流分布相对弥散；高精度tDCS（HD-tDCS）采用4×1环形小电极"
        "以约束电场并提高DLPFC聚焦性[@kuo2013]。本方案双侧环形布局为F3阳极（返回FP1、F7、C1、C3）与F4阴极（返回FP2、F8、C2、C4），"
        "±1.0 mA、20 min/次、连续14次（图2）。聚焦4×1几何、双侧DLPFC靶点与两周住院剂量，共同构成待检验的刺激学贡献。"
        "既往酒依赖HD-tDCS研究多关注渴求或代谢读出，而非完整标准化神经心理成套测验[@kumar2026]。"
    ),
    (
        "本研究在男性酒依赖住院患者中开展单盲三臂随机对照试验，检验14次每日双侧4×1 HD-tDCS较伪刺激或单纯常规住院治疗"
        "是否带来更大认知改善。预设主要结局为TMT-B、言语流畅性（VFT）和霍普金斯词语学习测验修订版（HVLT-R）；"
        "MMSE与数字广度测验（DST）为次要结局。"
    ),
]

DISCUSSION_ZH = [
    (
        "本三臂住院试验中，双侧4×1 DLPFC HD-tDCS在预设认知成套测验（TMT-B、VFT、HVLT-R）上相对伪刺激改善更大，"
        "并相对常规治疗具有平行优势。三组基线可比。85例随机入组，经删除真刺激组误录/重复行并将1例保留记录按随机化名单"
        "重归伪刺激组后，主要完整病例分析纳入77例（补充表S1）。"
    ),
    (
        "真–伪刺激对比在变化值、干预后比较与基线校正ANCOVA中一致，三项主要结局组间效应稳健（偏η²=0.19-0.31）；"
        "改变争议记录归属的敏感性分析仍保持ANCOVA结论。三组在脱毒与支持治疗期间均有一定改善[@wangsj2020]，"
        "伪刺激对照下的超额改善可将刺激特异性获益与非特异性住院恢复区分开。"
    ),
    (
        "伪刺激与常规治疗在主要结局上无显著差异。因对照组未匹配等时设备接触，该等价更说明真/伪刺激共享的额外疗程结构"
        "在本剂量下不足以产生超出常规治疗的可测认知增益，而非证明伪刺激电流认知惰性。"
    ),
    (
        "改善集中于TMT-B、VFT与HVLT-R，MMSE与DST无组间差异，宜解读为聚焦双侧DLPFC HD剂量下的域选择性行为剖面"
        "[@kuo2013,@dasilva2013,@camchong2023]。基线MMSE接近天花板（平均26-27/30）也限制了该次要筛查指标的进一步变化空间。"
        "心理测量终点为预设选择；本试验未以影像或波谱定位机制。"
    ),
    (
        "邻近HD-tDCS–酒依赖研究多聚焦渴求与前额叶代谢：联合纳曲酮的左DLPFC HD-tDCS可改变Glx，但渴求组间效应不显著[@kumar2026]。"
        "常规海绵电极DLPFC-tDCS则报告渴求、复饮、P3/FAB筛查或混合神经心理结局"
        "[@nakamura2012,@dasilva2013,@dayal2023,@klauss2018,@bollen2022]，而非双侧4×1 HD剂量下的完整伪刺激对照标准化成套测验。"
        "本研究以标准化神经心理评分为终点，与上述读出形成互补。"
    ),
    (
        "探索性相关共20对饮酒–认知组合，两项经Benjamini-Hochberg校正后仍显著：累计饮酒时间（r=0.42, q=0.003）与"
        "成瘾时间（r=0.33, q=0.037）越长，基线TMT-B越慢。累积暴露可能尤其加重认知转换负担，有待更大样本复核。"
    ),
    (
        "推广范围受限于单中心男性住院样本（核查后真刺激n=21）、评估者未盲、常规治疗组未匹配设备接触及未前瞻注册。"
        "观察到的真–伪刺激效应超过n=21 vs 27、80%效能下最小可检测d=0.83。变化值与ANCOVA可降低基线不平衡担忧，"
        "但不能消除残余混杂。"
    ),
    (
        "14次双侧DLPFC HD-tDCS与男性酒依赖住院患者标准化认知成套测验的伪刺激对照改善相关。"
        "预注册、多中心、匹配注意对照、客观生物标志物、结构化安全监测及复饮随访的试验，将决定这些认知获益能否转化为"
        "持久临床效益。"
    ),
]
