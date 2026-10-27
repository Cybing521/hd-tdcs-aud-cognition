#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HD-tDCS / AUD —— 复算全部已发表统计量 + 本轮新增分析（Frontiers 修稿用）。

输入：04_raw_data/upstream_repo/data/deid/*.xlsx（去标识分析表，GitHub 仓库
      Cybing521/hd-tdcs-aud-cognition 同步而来）
输出：06_results/*.csv、06_results/*.json、06_results/复算与新增分析报告.md

口径完全沿用上轮 03_analysis/ 脚本（Levene→ANOVA/Welch、LSD/Games-Howell、
BH-FDR、d 与 g 的 95%CI、ANCOVA 嵌套 F 检验），并对上游
04_data/processed/*.json 逐项做一致性断言。

本轮新增（回应 Reviewer 2 关于分析集选择 / 随机化后数据处理 / 统计方法的意见）：
  A. 退出者缺失结局的 tipping-point 分析（3 例提前出院者无第 4 周评估）
  B. 置换检验（标签置换，固定种子）
  C. Bootstrap 百分位 CI（固定种子）
  D. ANCOVA 稳健性：HC3 稳健协方差
  E. 残差正态性与方差齐性诊断
  F. 样本流向逐级核算（85 → 82 → 77）

运行：python3 03_scripts/recompute_all.py
"""

from __future__ import annotations

import csv
import json
import math
from pathlib import Path

import numpy as np
from scipy import stats
import statsmodels.api as sm

ROOT = Path(__file__).resolve().parents[1]


def _first_existing(*candidates: Path) -> Path:
    for c in candidates:
        if c.exists():
            return c
    return candidates[0]


# 目录布局兼容：项目/仓库布局（04_raw_data/upstream_repo/...）与独立可复现性包布局（data/...）
DEID = _first_existing(ROOT / "04_raw_data" / "upstream_repo" / "data" / "deid",
                        ROOT / "data" / "deid")
UPSTREAM = _first_existing(ROOT / "04_raw_data" / "upstream_repo" / "04_data" / "processed",
                           ROOT / "04_data" / "processed",
                           ROOT / "upstream" / "04_data" / "processed")
OUT = _first_existing(ROOT / "06_results", ROOT / "results")
OUT.mkdir(parents=True, exist_ok=True)

SEED = 20261027
N_BOOT = 10000
N_PERM = 10000

GROUP_LABEL = {1: "active", 2: "sham", 3: "control"}
COG = {
    "MMSE": ("MMSE1", "MMSE2"),
    "TMT-B": ("TMT-B1", "TMT-B2"),
    "DST": ("DST1", "DST2"),
    "VFT": ("VFT1", "VFT2"),
    "HVLT-R": ("HVLT-R1", "HVLT-R2"),
}
PRIMARY = ["TMT-B", "VFT", "HVLT-R"]
SECONDARY = ["MMSE", "DST"]
DEMO = ["年龄", "教育年限", "首次饮酒年龄", "累计使用时间", "成瘾时间", "戒酒次数"]
DRINK = {
    "首次饮酒年龄": "age_first_drink",
    "累计使用时间": "duration_use",
    "成瘾时间": "dependence_duration",
    "戒酒次数": "abstinence_attempts",
}
# 结局方向：+1 表示数值越大越好，-1 表示数值越小越好（用于构造最不利插补方向）
DIRECTION = {"MMSE": 1, "TMT-B": -1, "DST": 1, "VFT": 1, "HVLT-R": 1}
# 量表理论取值范围（用于判定插补值是否物理可能）
SCALE_RANGE = {
    "MMSE": (0, 30),
    "TMT-B": (0, None),      # 完成秒数，不能为负
    "DST": (0, 23),          # 顺背至 13 位 + 倒背至 10 位
    "VFT": (0, 60),          # 60 秒内正确说出数
    "HVLT-R": (0, 36),       # 3 次复述 × 12 个词
}


# ---------------------------------------------------------------- 基础工具

MASTER_CANDIDATES = ("比较_deid.xlsx", "comparison_deid.xlsx")


def resolve_master() -> Path:
    """主表文件名兼容：仓库里用原中文名，可复现性包里用 ASCII 名。"""
    for name in MASTER_CANDIDATES:
        if (DEID / name).exists():
            return DEID / name
    raise FileNotFoundError(f"在 {DEID} 未找到主表，候选：{MASTER_CANDIDATES}")


def load_rows() -> list[dict]:
    """读取已清洗主表（77 行 = 唯一受试者）。"""
    import openpyxl

    wb = openpyxl.load_workbook(resolve_master(), data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append(dict(zip(headers, row)))
    wb.close()
    return rows


def vals(rows, key, grp=None) -> np.ndarray:
    out = []
    for r in rows:
        if grp is not None and r.get("分组") != grp:
            continue
        v = r.get(key)
        if v is None or v == "":
            continue
        out.append(float(v))
    return np.asarray(out, dtype=float)


def paired(rows, c_pre, c_post, grp=None) -> np.ndarray:
    out = []
    for r in rows:
        if grp is not None and r.get("分组") != grp:
            continue
        a, b = r.get(c_pre), r.get(c_post)
        if a is None or b is None or a == "" or b == "":
            continue
        out.append(float(b) - float(a))
    return np.asarray(out, dtype=float)


def mean_sd(a) -> tuple[float, float, int]:
    a = np.asarray(a, dtype=float)
    if a.size == 0:
        return float("nan"), float("nan"), 0
    return float(a.mean()), float(a.std(ddof=1)), int(a.size)


def fmt_ms(a) -> str:
    m, s, _ = mean_sd(a)
    return f"{m:.2f}±{s:.2f}"


def oneway(by: dict[int, np.ndarray]) -> dict:
    """Levene → 经典 ANOVA 或 Welch；返回 F、P、方法。"""
    groups = [by[g] for g in sorted(by)]
    if any(len(g) < 2 for g in groups):
        return {"F": None, "P": None, "method": "insufficient n", "equal_var": True}
    lev = stats.levene(*groups)
    equal = lev.pvalue >= 0.05
    if equal:
        f, p = stats.f_oneway(*groups)
    else:
        res = sm.stats.anova_oneway(groups, use_var="unequal")
        f, p = res.statistic, res.pvalue
    return {
        "F": float(f),
        "P": float(p),
        "levene_P": float(lev.pvalue),
        "equal_var": bool(equal),
        "method": "ANOVA" if equal else "Welch",
    }


def pairwise_p(by: dict[int, np.ndarray], equal_var: bool) -> dict:
    out = {}
    for a, b in [(1, 2), (1, 3), (2, 3)]:
        if len(by[a]) < 2 or len(by[b]) < 2:
            out[f"{a}v{b}"] = float("nan")
            continue
        _, p = stats.ttest_ind(by[a], by[b], equal_var=bool(equal_var))
        out[f"{a}v{b}"] = float(p)
    return out


def fdr_bh(p_by: dict[str, float]) -> dict[str, float]:
    items = sorted(p_by.items(), key=lambda kv: kv[1])
    m = len(items)
    adj, prev = {}, 1.0
    for i in range(m - 1, -1, -1):
        k, p = items[i]
        val = min(prev, p * m / (i + 1))
        adj[k] = float(val)
        prev = val
    return adj


def cohens_d_ci(a, b) -> dict:
    a, b = np.asarray(a, float), np.asarray(b, float)
    na, nb = len(a), len(b)
    if na < 2 or nb < 2:
        return {}
    va, vb = a.var(ddof=1), b.var(ddof=1)
    sp = math.sqrt(((na - 1) * va + (nb - 1) * vb) / (na + nb - 2))
    d = float((a.mean() - b.mean()) / sp)
    se = math.sqrt((na + nb) / (na * nb) + d**2 / (2 * (na + nb)))
    j = 1 - 3 / (4 * (na + nb) - 9)
    g = d * j
    se_g = math.sqrt((na + nb) / (na * nb) + g**2 / (2 * (na + nb)))
    return {
        "d": round(d, 2),
        "d_ci_lo": round(d - 1.959964 * se, 2),
        "d_ci_hi": round(d + 1.959964 * se, 2),
        "g": round(g, 2),
        "g_ci_lo": round(g - 1.959964 * se_g, 2),
        "g_ci_hi": round(g + 1.959964 * se_g, 2),
        "mean_diff": round(float(a.mean() - b.mean()), 2),
    }


def ancova(rows, c_pre, c_post) -> dict:
    """OLS ANCOVA：post ~ group + 基线，嵌套 F 检验 + 斜率齐性 + HC3 稳健 p。"""
    y, g, x = [], [], []
    for r in rows:
        pre, post, grp = r.get(c_pre), r.get(c_post), r.get("分组")
        if pre in (None, "") or post in (None, "") or grp is None:
            continue
        y.append(float(post))
        g.append(int(grp))
        x.append(float(pre))
    y, g, x = np.asarray(y), np.asarray(g), np.asarray(x)
    n = len(y)
    xc = x - x.mean()
    d_a, d_s = (g == 1).astype(float), (g == 2).astype(float)
    X_full = np.column_stack([np.ones(n), d_a, d_s, xc])
    X_red = np.column_stack([np.ones(n), xc])
    fit_full = sm.OLS(y, X_full).fit()
    fit_red = sm.OLS(y, X_red).fit()

    df_num = int(fit_red.df_resid - fit_full.df_resid)
    df_den = int(fit_full.df_resid)
    f_val = ((fit_red.ssr - fit_full.ssr) / df_num) / (fit_full.ssr / df_den)
    p_val = float(stats.f.sf(f_val, df_num, df_den))
    ss_eff = float(fit_red.ssr - fit_full.ssr)
    ss_err = float(fit_full.ssr)
    peta = ss_eff / (ss_eff + ss_err)

    def adj_mean(gv: int) -> float:
        row = np.array([1.0, 1.0 if gv == 1 else 0.0, 1.0 if gv == 2 else 0.0, 0.0])
        return float(row @ fit_full.params)

    # 斜率齐性：加 group × 基线交互
    X_int = np.column_stack([X_full, d_a * xc, d_s * xc])
    fit_int = sm.OLS(y, X_int).fit()
    f_int = ((fit_full.ssr - fit_int.ssr) / 2) / (fit_int.ssr / fit_int.df_resid)
    p_int = float(stats.f.sf(f_int, 2, fit_int.df_resid))

    # HC3 稳健协方差下 group 联合检验（Wald）
    fit_hc3 = sm.OLS(y, X_full).fit(cov_type="HC3")
    b = fit_hc3.params[1:3]
    cov = fit_hc3.cov_params()[1:3, 1:3]
    wald = float(b @ np.linalg.solve(cov, b))
    p_hc3 = float(stats.chi2.sf(wald, 2))

    return {
        "F": round(float(f_val), 3),
        "df1": df_num,
        "df2": df_den,
        "P": round(p_val, 5) if p_val >= 1e-5 else 0.0,
        "partial_eta_sq": round(peta, 3),
        "slope_homogeneity_F": round(float(f_int), 3),
        "slope_homogeneity_P": round(p_int, 3),
        "adjusted_means": {
            "active": round(adj_mean(1), 2),
            "sham": round(adj_mean(2), 2),
            "control": round(adj_mean(3), 2),
        },
        "baseline_mean": round(float(x.mean()), 2),
        "n": n,
        "hc3_wald_chi2": round(wald, 3),
        "hc3_P": round(p_hc3, 5) if p_hc3 >= 1e-5 else 0.0,
    }


# ---------------------------------------------------------------- 主复算

def recompute(rows) -> dict:
    res: dict = {}
    n_by = {g: int(len(vals(rows, "年龄", g))) for g in (1, 2, 3)}
    res["n"] = {GROUP_LABEL[g]: n_by[g] for g in (1, 2, 3)}
    res["total_n"] = len(rows)

    # Table 1 人口学与饮酒特征
    t1 = []
    for c in DEMO:
        by = {g: vals(rows, c, g) for g in (1, 2, 3)}
        ow = oneway(by)
        t1.append({
            "variable": c,
            "active": fmt_ms(by[1]),
            "sham": fmt_ms(by[2]),
            "control": fmt_ms(by[3]),
            "F": round(ow["F"], 3) if ow.get("F") is not None else None,
            "P": round(ow["P"], 3) if ow.get("P") is not None else None,
            "method": ow["method"],
        })
    res["table1"] = t1

    # Table 2 基线与干预后
    tb, tp = [], []
    for label, (c_pre, c_post) in COG.items():
        for phase, colname, bucket in [("baseline", c_pre, tb), ("post", c_post, tp)]:
            by = {g: vals(rows, colname, g) for g in (1, 2, 3)}
            ow = oneway(by)
            bucket.append({
                "outcome": label,
                "time": phase,
                "active": fmt_ms(by[1]),
                "sham": fmt_ms(by[2]),
                "control": fmt_ms(by[3]),
                "F": round(ow["F"], 3) if ow.get("F") is not None else None,
                "P": round(ow["P"], 3) if ow.get("P") is not None else None,
                "method": ow["method"],
            })
    res["table_baseline"] = tb
    res["table_post"] = tp

    # 干预后两两比较 + 效应量
    post_pair, effect = {}, {}
    for label in COG:
        c_post = COG[label][1]
        by = {g: vals(rows, c_post, g) for g in (1, 2, 3)}
        ow = oneway(by)
        pw = pairwise_p(by, ow["equal_var"])
        post_pair[label] = {
            "overall_P": round(ow["P"], 3),
            "active_vs_sham_P": round(pw["1v2"], 3),
            "active_vs_control_P": round(pw["1v3"], 3),
            "sham_vs_control_P": round(pw["2v3"], 3),
            "method": ow["method"],
            "levene_P": round(ow["levene_P"], 4),
        }
        effect[label] = {
            "active_vs_sham": cohens_d_ci(by[1], by[2]),
            "active_vs_control": cohens_d_ci(by[1], by[3]),
        }
    res["pairwise_post"] = post_pair
    res["effect_sizes_post"] = effect

    # FDR（5 项干预后总体检验）
    raw_p = {r["outcome"]: r["P"] for r in tp if r["P"] is not None}
    adj = fdr_bh(raw_p)
    res["multiplicity"] = {
        "method": "Benjamini-Hochberg FDR across five post-intervention omnibus tests",
        "raw_p": raw_p,
        "fdr_q": {k: round(v, 4) for k, v in adj.items()},
        "primary_significant_at_fdr_0_05": {k: bool(adj.get(k, 1) < 0.05) for k in PRIMARY},
    }

    # 组内配对 t 检验
    within = []
    for label, (c_pre, c_post) in COG.items():
        for g in (1, 2, 3):
            pre, post = [], []
            for r in rows:
                if r.get("分组") != g:
                    continue
                a, b = r.get(c_pre), r.get(c_post)
                if a in (None, "") or b in (None, ""):
                    continue
                pre.append(float(a))
                post.append(float(b))
            if len(pre) < 3:
                continue
            t, p = stats.ttest_rel(pre, post)
            within.append({
                "outcome": label,
                "group": GROUP_LABEL[g],
                "pre": fmt_ms(pre),
                "post": fmt_ms(post),
                "delta": round(float(np.mean(post) - np.mean(pre)), 2),
                "t": round(float(t), 3),
                "P": round(float(p), 4),
                "n": len(pre),
            })
    res["within_group_change"] = within

    # Table 3 差值
    change_rows, change_pair, change_effect = [], {}, {}
    for label, (c_pre, c_post) in COG.items():
        by = {g: paired(rows, c_pre, c_post, g) for g in (1, 2, 3)}
        ow = oneway(by)
        change_rows.append({
            "outcome": label,
            "active": fmt_ms(by[1]),
            "sham": fmt_ms(by[2]),
            "control": fmt_ms(by[3]),
            "F": round(ow["F"], 3) if ow.get("F") is not None else None,
            "P": round(ow["P"], 4) if ow.get("P") is not None else None,
            "method": ow["method"],
        })
        pw = pairwise_p(by, ow["equal_var"])
        change_pair[label] = {
            "overall_P": round(ow["P"], 4),
            "active_vs_sham_P": round(pw["1v2"], 4),
            "active_vs_control_P": round(pw["1v3"], 4),
            "sham_vs_control_P": round(pw["2v3"], 4),
        }
        change_effect[label] = {
            "active_vs_sham": cohens_d_ci(by[1], by[2]),
            "active_vs_control": cohens_d_ci(by[1], by[3]),
        }
    res["change_score"] = change_rows
    res["change_score_pairwise"] = change_pair
    res["effect_sizes_change"] = change_effect

    # Table 4 ANCOVA
    res["ancova"] = {label: ancova(rows, c_pre, c_post) for label, (c_pre, c_post) in COG.items()}

    # 相关分析（20 格 + BH）
    cells, ps = [], []
    for dc, dname in DRINK.items():
        for cc, cname in COG.items():
            x, y = vals(rows, dc), vals(rows, COG[cc][0])
            m = min(len(x), len(y))
            x, y = x[:m], y[:m]
            ok = ~(np.isnan(x) | np.isnan(y))
            x, y = x[ok], y[ok]
            r, p = stats.pearsonr(x, y)
            cells.append({
                "drinking": dname,
                "cognition": cname,
                "r": round(float(r), 3),
                "P": round(float(p), 5) if p >= 1e-5 else 0.0,
                "n": int(m),
            })
            ps.append(float(p))
    order = np.argsort(ps)
    m_all = len(ps)
    q = [0.0] * m_all
    prev = 1.0
    for rank, idx in enumerate(reversed(order)):
        i = m_all - rank
        val = min(prev, ps[idx] * m_all / i)
        q[idx] = val
        prev = val
    for cell, qv in zip(cells, q):
        cell["q_bh"] = round(min(qv, 1.0), 3)
    res["correlations"] = {
        "n_tests": m_all,
        "nominal_sig_count": sum(1 for c in cells if c["P"] < 0.05),
        "fdr_sig_count": sum(1 for c in cells if c["q_bh"] < 0.05),
        "cells": cells,
    }

    # 检验力与最小可检测效应
    power = {}
    for label in PRIMARY:
        es = effect[label]["active_vs_sham"]
        d = abs(es["d"])
        n1, n2 = n_by[1], n_by[2]
        se = math.sqrt((n1 + n2) / (n1 * n2) + d * d / (2 * (n1 + n2)))
        z_a = stats.norm.ppf(1 - 0.05 / 2)
        z_b = stats.norm.ppf(0.80)
        ncp = d / se
        pw = float(stats.norm.cdf(-z_a - ncp) + stats.norm.cdf(z_a - ncp))
        # 最小可检测 d（80% power，双侧 0.05）
        d_min = (z_a + z_b) * math.sqrt((n1 + n2) / (n1 * n2))
        power[label] = {
            "cohens_d": d,
            "post_hoc_power": round(pw, 3),
            "min_detectable_d_80pct": round(d_min, 2),
        }
    res["power_and_mde"] = power
    return res


def rows_df_grp(rows, g):
    return [r for r in rows if r.get("分组") == g]


# ---------------------------------------------------------------- 新增分析

def tipping_point(rows) -> dict:
    """A. 3 例提前出院者（无第 4 周评估）的缺失结局 tipping-point 分析。

    3 名受试者因提前出院未完成第 4 周评估，结局缺失且去标识表内无其基线，
    因此不能做基于模型的插补（无协变量）——改为「最不利方向偏移插补」：
      · 真刺激组缺失者：向不利于真刺激的方向偏移 Δ
      · 伪刺激组缺失者（2 例）：向不利于真刺激的方向偏移 Δ
    Δ 自 0 起步增大，记录 P 越过 0.05 的临界 Δ（以单位与各自组标准差表示）。
    """
    out = {"design": "worst-direction offset imputation for 3 participants without week-4 assessment",
           "missing_by_arm": {"active": 1, "sham": 2, "control": 0},
           "note": "这些受试者的基线与结局均不在去标识分析表内，故无法做协变量插补（multiple imputation）；"
                   "本分析以偏移插补给出结论的脆弱性边界。",
           "outcomes": {}}
    for label in PRIMARY:
        c_post = COG[label][1]
        by = {g: vals(rows, c_post, g) for g in (1, 2, 3)}
        sd_a, sd_s = float(by[1].std(ddof=1)), float(by[2].std(ddof=1))
        mean_a, mean_s = float(by[1].mean()), float(by[2].mean())
        direction = DIRECTION[label]
        obs_p = float(stats.ttest_ind(by[1], by[2], equal_var=True).pvalue)
        obs_d = cohens_d_ci(by[1], by[2])["d"]

        # Δ 网格：0 → 8 SD
        step = 0.05
        max_delta = 8.0
        grid = np.arange(0.0, max_delta + step / 2, step)
        tipping = None
        curve = []
        for delta in grid:
            a = np.append(by[1], mean_a - direction * delta * sd_a)
            s = np.append(by[2], mean_s + direction * delta * sd_s)
            p = float(stats.ttest_ind(a, s, equal_var=True).pvalue)
            d = cohens_d_ci(a, s)["d"]
            curve.append({"delta_sd": round(float(delta), 2), "P": round(p, 5), "d": round(d, 2)})
            if tipping is None and p >= 0.05:
                tipping = {
                    "delta_sd": round(float(delta), 2),
                    "delta_units": round(float(delta * sd_a), 2),
                    "P_at_tipping": round(p, 5),
                    "d_at_tipping": round(d, 2),
                    "imputed_active_post": round(float(mean_a - direction * delta * sd_a), 2),
                    "imputed_sham_post": round(float(mean_s + direction * delta * sd_s), 2),
                }
        # 极端最不利：真刺激组缺失者取本组最差观测值；伪刺激组缺失者取本组最好观测值
        worst_active = float(by[1].min()) if direction == 1 else float(by[1].max())
        best_sham = float(by[2].max()) if direction == 1 else float(by[2].min())
        a_worst = np.append(by[1], worst_active)
        s_best = np.append(by[2], [best_sham, best_sham])
        p_worst = float(stats.ttest_ind(a_worst, s_best, equal_var=True).pvalue)
        out["outcomes"][label] = {
            "observed_active_vs_sham_P": round(obs_p, 5),
            "observed_d": obs_d,
            "tipping": tipping,
            "tipping_absent_within_8sd": tipping is None,
            "grid_max_delta_sd": max_delta,
            "extreme_worst_case_P": round(p_worst, 5),
            "extreme_worst_case_d": cohens_d_ci(a_worst, s_best)["d"],
            "extreme_worst_case_imputed": {
                "active_missing": round(worst_active, 2),
                "sham_missing": round(best_sham, 2),
            },
            "arm_sd": {"active": round(sd_a, 2), "sham": round(sd_s, 2)},
            "observed_range": {
                "active": [round(float(by[1].min()), 2), round(float(by[1].max()), 2)],
                "sham": [round(float(by[2].min()), 2), round(float(by[2].max()), 2)],
            },
            "scale_range": list(SCALE_RANGE[label]),
            "curve": curve[::4],
        }
        # 插补值是否落在「量表可能范围」与「本组观测范围」内
        lo_scale, hi_scale = SCALE_RANGE[label]
        def plausible(v, arm_range):
            if lo_scale is not None and v < lo_scale:
                return False
            if hi_scale is not None and v > hi_scale:
                return False
            return arm_range[0] <= v <= arm_range[1]
        arm_range = {
            "active": [float(by[1].min()), float(by[1].max())],
            "sham": [float(by[2].min()), float(by[2].max())],
        }
        if tipping is not None:
            out["outcomes"][label]["tipping_imputed_plausible"] = {
                "active_within_scale": plausible(tipping["imputed_active_post"], arm_range["active"]),
                "sham_within_scale": plausible(tipping["imputed_sham_post"], arm_range["sham"]),
                "note": "False 表示翻转结论所需的插补值已超出量表理论取值范围（物理不可能）",
            }
        out["outcomes"][label]["extreme_worst_case_plausible"] = {
            "active_within_scale": plausible(worst_active, arm_range["active"]),
            "sham_within_scale": plausible(best_sham, arm_range["sham"]),
        }
    return out


def permutation_tests(rows) -> dict:
    """B. 标签置换检验（固定种子）：三组总体与真刺激-伪刺激对比。

    总体检验：77 例标签整体置换；对比检验：仅在真刺激/伪刺激 48 例内受限置换
    （对照臂标签固定），这是成对比较的标准做法。
    """
    rng = np.random.default_rng(SEED)
    out = {"n_perm": N_PERM, "seed": SEED, "outcomes": {}}
    for label in PRIMARY:
        c_post = COG[label][1]
        y = vals(rows, c_post)
        g = np.array([r.get("分组") for r in rows if r.get(c_post) not in (None, "")], dtype=int)
        assert len(y) == len(g) == 77

        obs_f = float(stats.f_oneway(y[g == 1], y[g == 2], y[g == 3]).statistic)
        obs_p_as = float(stats.ttest_ind(y[g == 1], y[g == 2], equal_var=True).pvalue)
        obs_diff = float(y[g == 1].mean() - y[g == 2].mean())

        cnt_f = 0
        for _ in range(N_PERM):
            perm = rng.permutation(g)          # 只置换标签，观测值不动
            f = float(stats.f_oneway(y[perm == 1], y[perm == 2], y[perm == 3]).statistic)
            if f >= obs_f:
                cnt_f += 1

        # 受限置换：只在 active/sham 之间交换标签
        mask = (g == 1) | (g == 2)
        y_as, g_as = y[mask], g[mask]
        obs_diff_as = float(y_as[g_as == 1].mean() - y_as[g_as == 2].mean())
        cnt_as = 0
        for _ in range(N_PERM):
            swap = rng.permutation(g_as)       # 1/2 互换标签
            diff = float(y_as[swap == 1].mean() - y_as[swap == 2].mean())
            if abs(diff) >= abs(obs_diff_as):
                cnt_as += 1

        out["outcomes"][label] = {
            "obs_F": round(obs_f, 3),
            "perm_p_omnibus": round((cnt_f + 1) / (N_PERM + 1), 5),
            "obs_active_vs_sham_P": round(obs_p_as, 5),
            "obs_mean_diff": round(obs_diff, 2),
            "perm_p_active_vs_sham": round((cnt_as + 1) / (N_PERM + 1), 5),
            "perm_scheme": "omnibus: 77 labels permuted; contrast: restricted permutation within active+sham (48)",
        }
    return out


def bootstrap_cis(rows) -> dict:
    """C. Bootstrap 百分位 CI（固定种子）：真刺激-伪刺激差值。"""
    rng = np.random.default_rng(SEED)
    out = {"n_boot": N_BOOT, "seed": SEED, "outcomes": {}}
    for label in COG:
        c_pre, c_post = COG[label]
        for kind, getter in [
            ("post", lambda g: vals(rows, c_post, g)),
            ("change", lambda g: paired(rows, c_pre, c_post, g)),
        ]:
            a, b = getter(1), getter(2)
            obs = float(a.mean() - b.mean())
            diffs = np.empty(N_BOOT)
            for i in range(N_BOOT):
                ra = rng.choice(a, size=len(a), replace=True)
                rb = rng.choice(b, size=len(b), replace=True)
                diffs[i] = ra.mean() - rb.mean()
            lo, hi = np.percentile(diffs, [2.5, 97.5])
            # BCa 简化：只报百分位 + 正态近似对比
            se = float(diffs.std(ddof=1))
            out["outcomes"].setdefault(label, {})[kind] = {
                "observed_diff": round(obs, 2),
                "boot_ci_lo": round(float(lo), 2),
                "boot_ci_hi": round(float(hi), 2),
                "boot_se": round(se, 2),
                "excludes_zero": bool(lo > 0 or hi < 0),
            }
    return out


def diagnostics(rows) -> dict:
    """E. 诊断：残差正态性、方差齐性、极端值影响（留一法最大影响）。"""
    out = {"outcomes": {}}
    for label, (c_pre, c_post) in COG.items():
        # 干预后分数正态性
        by = {g: vals(rows, c_post, g) for g in (1, 2, 3)}
        shap = {GROUP_LABEL[g]: round(float(stats.shapiro(by[g]).pvalue), 4) for g in (1, 2, 3)}
        lev = float(stats.levene(by[1], by[2], by[3]).pvalue)
        # 留一法：剔除单个受试者后 active-vs-sham 的 P 与 d 变化
        deltas = []
        allrows = [r for r in rows if r.get(c_post) not in (None, "")]
        for r in allrows:
            sub = [x for x in allrows if x is not r]
            a = vals(sub, c_post, 1)
            s = vals(sub, c_post, 2)
            deltas.append({
                "P": round(float(stats.ttest_ind(a, s, equal_var=True).pvalue), 5),
                "d": cohens_d_ci(a, s)["d"],
            })
        ps = [d["P"] for d in deltas]
        ds = [d["d"] for d in deltas]
        out["outcomes"][label] = {
            "shapiro_W_P_by_arm": shap,
            "levene_P": round(lev, 4),
            "loo_active_vs_sham_P_range": [round(min(ps), 5), round(max(ps), 5)],
            "loo_active_vs_sham_d_range": [round(min(ds), 2), round(max(ds), 2)],
            "all_loo_P_below_0_05": bool(max(ps) < 0.05),
        }
    return out


def sample_flow() -> dict:
    """F. 样本流向逐级核算（依据 CONSORT 记录与数据清洗日志）。"""
    try:
        log = json.loads((UPSTREAM / "data_cleaning_log.json").read_text(encoding="utf-8"))
        flow = json.loads((UPSTREAM / "consort_flow.json").read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {"skipped": True,
                "reason": "缺少记录核查日志（data_cleaning_log.json / consort_flow.json）",
                "note": "参与者流向的逐级核算见补充材料 Table S1；本包不含上轮流程日志。"}
    randomized = flow["flow"]["randomized"]
    disc = flow["flow"]["discontinued_intervention"]
    before = log["sample_size_before"]
    after = log["sample_size_after"]
    return {
        "randomized": randomized,
        "allocated": flow["flow"]["allocated"],
        "discontinued_before_week4": {
            "reason": disc["reason"],
            "active": disc["active_hd_tdcs"],
            "sham": disc["sham_hd_tdcs"],
            "control": disc["control"],
            "total": disc["total"],
        },
        "records_with_week4_assessment": before,
        "unique_participants_in_analysis": after,
        "records_removed": {
            "count": len(log["rows_removed_from_active"]),
            "reason": "5 例真刺激组误录入/重复行（含 1 例按原始随机化名单改归伪刺激组）",
        },
        "identity_check": {
            "randomized_equals_records_plus_discontinued": randomized == sum(before.values()) + disc["total"],
            "analysis_plus_removed_plus_discontinued_equals_randomized":
                sum(after.values()) + len(log["rows_removed_from_active"]) + disc["total"] == randomized,
        },
        "outcomes_missing_by_arm": {"active": 1, "sham": 2, "control": 0},
    }


# ---------------------------------------------------------------- 上游一致性核对

def crosscheck(res: dict) -> list[dict]:
    """把复算值与上轮 04_data/processed/*.json 逐项比对。"""
    checks: list[dict] = []

    def add(name, upstream, mine, tol=0.006):
        if upstream is None or mine is None:
            checks.append({"item": name, "upstream": upstream, "recomputed": mine, "match": False})
            return
        if isinstance(upstream, str) or isinstance(mine, str):
            ok = str(upstream).strip() == str(mine).strip()
        else:
            ok = abs(float(upstream) - float(mine)) <= tol
        checks.append({"item": name, "upstream": upstream, "recomputed": mine, "match": bool(ok)})

    up = json.loads((UPSTREAM / "analysis_results.json").read_text(encoding="utf-8"))
    en = json.loads((UPSTREAM / "enhanced_analyses.json").read_text(encoding="utf-8"))

    add("n_active", up["n"]["1"], res["n"]["active"])
    add("n_sham", up["n"]["2"], res["n"]["sham"])
    add("n_control", up["n"]["3"], res["n"]["control"])

    for i, row in enumerate(res["table1"]):
        up_row = up["tables"]["table1"][i]
        for k in ("active", "sham", "control"):
            add(f"table1[{row['variable']}].{k}", up_row[k], row[k], tol=0.0)
        add(f"table1[{row['variable']}].P", up_row["P"], row["P"])

    for i, row in enumerate(res["table_post"]):
        up_row = up["tables"]["table_post"][i]
        for k in ("active", "sham", "control"):
            add(f"table_post[{row['outcome']}].{k}", up_row[k], row[k], tol=0.0)
        add(f"table_post[{row['outcome']}].P", up_row["P"], row["P"])

    for label in PRIMARY:
        add(f"post_P[{label}]", up["pairwise_post"][label]["overall_P"], res["pairwise_post"][label]["overall_P"])
        add(f"post_AS_P[{label}]", up["pairwise_post"][label]["active_vs_sham"],
            res["pairwise_post"][label]["active_vs_sham_P"])
        add(f"d_AS[{label}]", up["effect_sizes"][label]["active_vs_sham_d"],
            res["effect_sizes_post"][label]["active_vs_sham"]["d"])
        up_es = en["effect_sizes_post_ci"][label]["active_vs_sham"]
        mine_es = res["effect_sizes_post"][label]["active_vs_sham"]
        add(f"d_AS_CI_lo[{label}]", up_es["ci_lo"], mine_es["d_ci_lo"])
        add(f"d_AS_CI_hi[{label}]", up_es["ci_hi"], mine_es["d_ci_hi"])
        add(f"g_AS[{label}]", up_es["g"], mine_es["g"])
        up_an = en["ancova"][label]
        mine_an = res["ancova"][label]
        add(f"ancova_F[{label}]", up_an["F"], mine_an["F"])
        add(f"ancova_P[{label}]", up_an["P"], mine_an["P"], tol=0.0001)
        add(f"ancova_peta2[{label}]", up_an["partial_eta_sq"], mine_an["partial_eta_sq"])
        for g in ("active", "sham", "control"):
            add(f"ancova_adjmean[{label}][{g}]", up_an["adjusted_means"][g], mine_an["adjusted_means"][g])
        add(f"ancova_slopehomog_P[{label}]", up_an["slope_homogeneity_P"], mine_an["slope_homogeneity_P"])

    for i, row in enumerate(res["change_score"]):
        up_row = up["change_score_anova"][i]
        for k in ("active", "sham", "control"):
            mine_mean = float(str(row[k]).split("±")[0])
            add(f"change_mean[{row['outcome']}][{k}]", up_row[f"{k}_delta"], mine_mean, tol=0.006)
        add(f"change_P[{row['outcome']}]", up_row["P"], row["P"], tol=0.0002)

    add("corr_nominal_sig", en["correlation_matrix"]["nominal_sig_count"], res["correlations"]["nominal_sig_count"], tol=0)
    add("corr_fdr_sig", en["correlation_matrix"]["fdr_sig_count"], res["correlations"]["fdr_sig_count"], tol=0)
    add("corr_min_q", en["correlation_matrix"]["min_q"], min(c["q_bh"] for c in res["correlations"]["cells"]), tol=0.001)
    return checks


# ---------------------------------------------------------------- 输出

def write_csv(path: Path, rows: list[dict], fields: list[str] | None = None) -> None:
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fields = fields or list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def main() -> None:
    rows = load_rows()
    assert len(rows) == 77, f"主表行数异常：{len(rows)}（应为 77）"
    assert len({r["study_id"] for r in rows}) == 77, "study_id 存在重复，去标识表不是唯一受试者集合"

    res = recompute(rows)
    res["tipping_point"] = tipping_point(rows)
    res["permutation"] = permutation_tests(rows)
    res["bootstrap"] = bootstrap_cis(rows)
    res["diagnostics"] = diagnostics(rows)
    res["sample_flow"] = sample_flow()
    try:
        checks = crosscheck(res)
        res["upstream_consistency"] = {
            "total": len(checks),
            "matched": sum(1 for c in checks if c["match"]),
            "failed": [c for c in checks if not c["match"]],
        }
    except FileNotFoundError as exc:
        res["upstream_consistency"] = {
            "skipped": True,
            "reason": f"缺少上轮结果表（{exc.filename}），无法做逐项交叉核对",
        }

    (OUT / "复算结果_全量.json").write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(OUT / "表1_人口学与饮酒特征.csv", res["table1"])
    write_csv(OUT / "表2_认知评分_基线与干预后.csv", res["table_baseline"] + res["table_post"])
    write_csv(OUT / "表3_变化量.csv", res["change_score"])
    write_csv(OUT / "表4_ANCOVA.csv", [
        {"outcome": k, **v} for k, v in res["ancova"].items()
    ])
    write_csv(OUT / "新增A_tipping_point.csv", [
        {"outcome": k, **{kk: vv for kk, vv in v.items() if kk != "curve"}}
        for k, v in res["tipping_point"]["outcomes"].items()
    ])
    write_csv(OUT / "新增B_置换检验.csv", [
        {"outcome": k, **v} for k, v in res["permutation"]["outcomes"].items()
    ])
    write_csv(OUT / "新增C_bootstrap.csv", [
        {"outcome": k, "kind": kind, **v}
        for k, d in res["bootstrap"]["outcomes"].items() for kind, v in d.items()
    ])
    write_csv(OUT / "新增E_诊断.csv", [
        {"outcome": k, **{kk: json.dumps(vv, ensure_ascii=False) if isinstance(vv, (dict, list)) else vv
                          for kk, vv in v.items()}}
        for k, v in res["diagnostics"]["outcomes"].items()
    ])
    write_csv(OUT / "上游一致性核对.csv", checks)

    print("== 复算完成 ==")
    print("n:", res["n"], "total", res["total_n"])
    if res["upstream_consistency"].get("skipped"):
        print("上游一致性：跳过（", res["upstream_consistency"]["reason"], "）")
    else:
        print("上游一致性：", res["upstream_consistency"]["matched"], "/",
              res["upstream_consistency"]["total"], "一致")
        for c in res["upstream_consistency"]["failed"]:
            print("  不一致:", c)
    print("\n-- tipping point --")
    for k, v in res["tipping_point"]["outcomes"].items():
        print(f"  {k}: obs P={v['observed_active_vs_sham_P']}, d={v['observed_d']}, "
              f"tipping={v['tipping']}, 8SD 内不翻转={v['tipping_absent_within_8sd']}, "
              f"极端最不利 P={v['extreme_worst_case_P']}")
    print("\n-- 置换检验 --")
    for k, v in res["permutation"]["outcomes"].items():
        print(f"  {k}: 总体 perm P={v['perm_p_omnibus']}, 真-伪 perm P={v['perm_p_active_vs_sham']}")
    print("\n-- bootstrap (真刺激-伪刺激) --")
    for k, d in res["bootstrap"]["outcomes"].items():
        print(f"  {k}: post {d['post']} | change {d['change']}")


if __name__ == "__main__":
    main()
