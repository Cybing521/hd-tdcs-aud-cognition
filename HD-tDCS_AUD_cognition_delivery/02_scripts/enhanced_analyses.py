#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Enhanced SCI-level analyses on the cleaned study dataset.

Adds, on top of clean_and_analyze.py outputs:
1. Baseline-adjusted ANCOVA (post ~ group + baseline) per outcome,
   including homogeneity-of-slopes checks and adjusted group means.
2. Cohen's d with 95% CI (normal approximation) plus Hedges g correction
   for primary-outcome contrasts (post-intervention and change scores).
3. Mean±SD of change scores per group (Table 4 upgrade).
4. Full 20-cell Pearson correlation grid (4 drinking characteristics x 5
   baseline cognitive measures) with Benjamini-Hochberg q values -- replaces
   selective reporting of significant cells only.

Output: 04_data/processed/enhanced_analyses.json
"""

from __future__ import annotations

import json
import math
from pathlib import Path

import numpy as np
import openpyxl
from scipy import stats
import statsmodels.api as sm

ROOT = Path(__file__).resolve().parents[1]
sys_dir = Path(__file__).resolve().parent

import sys

sys.path.insert(0, str(sys_dir))
from clean_and_analyze import TABLE_DIR, COG_BASE, DEMO_COLS, load_compare, clean_compare_rows

OUT_PATH = ROOT / "04_data" / "processed" / "enhanced_analyses.json"

GROUPS = (1, 2, 3)
OUTCOMES = ["MMSE", "TMT-B", "DST", "VFT", "HVLT-R"]


def col(rows: list[dict], key: str, grp: int | None = None) -> np.ndarray:
    vals = []
    for r in rows:
        if grp is not None and r.get("分组") != grp:
            continue
        v = r.get(key)
        if v is None or v == "":
            continue
        vals.append(float(v))
    return np.asarray(vals, dtype=float)


def paired_col(rows: list[dict], k1: str, k2: str, grp: int | None = None) -> tuple[np.ndarray, np.ndarray]:
    """Pairwise-complete extraction of two columns (aligned by row)."""
    a, b = [], []
    for r in rows:
        if grp is not None and r.get("分组") != grp:
            continue
        v1, v2 = r.get(k1), r.get(k2)
        if v1 is None or v2 is None or v1 == "" or v2 == "":
            continue
        a.append(float(v1))
        b.append(float(v2))
    return np.asarray(a), np.asarray(b)


def cohens_d_ci(a: np.ndarray, b: np.ndarray) -> dict:
    """Cohen's d with large-sample 95% CI and Hedges g correction."""
    na, nb = len(a), len(b)
    va, vb = np.var(a, ddof=1), np.var(b, ddof=1)
    spool = math.sqrt(((na - 1) * va + (nb - 1) * vb) / (na + nb - 2))
    d = float((np.mean(a) - np.mean(b)) / spool)
    se = math.sqrt((na + nb) / (na * nb) + d**2 / (2 * (na + nb)))
    lo, hi = d - 1.959964 * se, d + 1.959964 * se
    j = 1 - 3 / (4 * (na + nb) - 9)
    g = d * j
    se_g = math.sqrt((na + nb) / (na * nb) + g**2 / (2 * (na + nb)))
    return {
        "d": round(d, 2),
        "ci_lo": round(lo, 2),
        "ci_hi": round(hi, 2),
        "g": round(g, 2),
        "g_ci_lo": round(g - 1.959964 * se_g, 2),
        "g_ci_hi": round(g + 1.959964 * se_g, 2),
    }


def nested_f_test(fit_restricted, fit_full) -> tuple[float, float, int]:
    """Manual nested-model F test returning (F, P, df_denominator)."""
    df_num = int(fit_restricted.df_resid - fit_full.df_resid)
    df_den = int(fit_full.df_resid)
    ss_diff = fit_restricted.ssr - fit_full.ssr
    f_val = (ss_diff / df_num) / (fit_full.ssr / df_den)
    p_val = float(stats.f.sf(f_val, df_num, df_den))
    return float(f_val), p_val, df_den


def ancova(rows: list[dict], pre_col: str, post_col: str) -> dict:
    """OLS ANCOVA: post ~ group + baseline, with slope-homogeneity probe."""
    y, g, x = [], [], []
    for r in rows:
        pre, post, grp = r.get(pre_col), r.get(post_col), r.get("分组")
        if None in (pre, post, grp) or pre == "" or post == "":
            continue
        y.append(float(post))
        g.append(int(grp))
        x.append(float(pre))
    y, g, x = np.asarray(y), np.asarray(g), np.asarray(x)
    n = len(y)

    # design: intercept + two group dummies (control=reference) + centered baseline
    base_centered = x - x.mean()
    d_active = (g == 1).astype(float)
    d_sham = (g == 2).astype(float)
    X_full = np.column_stack([np.ones(n), d_active, d_sham, base_centered])
    X_red = np.column_stack([np.ones(n), base_centered])

    fit_full = sm.OLS(y, X_full).fit()
    fit_red = sm.OLS(y, X_red).fit()
    f_stat, p_val, df_den = nested_f_test(fit_red, fit_full)
    ss_effect = float(fit_red.ssr - fit_full.ssr)
    ss_error = float(fit_full.ssr)
    partial_eta2 = ss_effect / (ss_effect + ss_error)

    # adjusted means at overall baseline mean
    def adj_mean(grp_val: int) -> float:
        row = np.array([1.0, 1.0 if grp_val == 1 else 0.0, 1.0 if grp_val == 2 else 0.0, 0.0])
        return float(row @ fit_full.params)

    # homogeneity of slopes: add group x baseline interactions
    X_int = np.column_stack([X_full, d_active * base_centered, d_sham * base_centered])
    fit_int = sm.OLS(y, X_int).fit()
    f_int, p_int, _ = nested_f_test(fit_full, fit_int)

    return {
        "F": round(float(f_stat), 3),
        "df1": 2,
        "df2": df_den,
        "P": round(float(p_val), 5) if p_val >= 1e-5 else 0.0,
        "partial_eta_sq": round(partial_eta2, 3),
        "slope_homogeneity_P": round(float(p_int), 3),
        "slope_homogeneity_F": round(float(f_int), 3),
        "adjusted_means": {"active": round(adj_mean(1), 2), "sham": round(adj_mean(2), 2), "control": round(adj_mean(3), 2)},
        "baseline_mean": round(float(x.mean()), 2),
        "n": n,
    }


def change_sd(rows: list[dict]) -> list[dict]:
    out = []
    for label, (c_pre, c_post, _) in COG_BASE.items():
        entry = {"outcome": label}
        for grp in GROUPS:
            pre_v, post_v = paired_col(rows, c_pre, c_post, grp)
            d = post_v - pre_v
            entry[{1: "active", 2: "sham", 3: "control"}[grp]] = {
                "mean": round(float(np.mean(d)), 2),
                "sd": round(float(np.std(d, ddof=1)), 2),
                "n": int(len(d)),
            }
        out.append(entry)
    return out


def contrast_d_by_outcome(rows: list[dict], use_change: bool) -> dict:
    """Cohen's d (+CI) for active-vs-sham and active-vs-control."""
    out: dict[str, dict] = {}
    for label, (c_pre, c_post, _) in COG_BASE.items():
        entry = {}
        pairs = [("active_vs_sham", 1, 2), ("active_vs_control", 1, 3)]
        for pname, ga, gb in pairs:
            if use_change:
                va = paired_col(rows, c_pre, c_post, ga)[1] - paired_col(rows, c_pre, c_post, ga)[0]
                vb = paired_col(rows, c_pre, c_post, gb)[1] - paired_col(rows, c_pre, c_post, gb)[0]
            else:
                va, vb = col(rows, c_post, ga), col(rows, c_post, gb)
            entry[pname] = cohens_d_ci(va, vb)
        out[label] = entry
    return out


def full_correlation(rows: list[dict]) -> dict:
    drink_cols = {
        c: k
        for c, k in zip(
            DEMO_COLS[2:],
            ["age_first_drink", "duration_use", "dependence_duration", "abstinence_attempts"],
        )
    }
    cog_cols = {COG_BASE[k][0]: k for k in OUTCOMES}
    cells = []
    raw_p: list[float] = []
    for dc, dname in drink_cols.items():
        for cc, cname in cog_cols.items():
            xs, ys = col(rows, dc), col(rows, cc)
            m = min(len(xs), len(ys))
            xs, ys = xs[:m], ys[:m]
            ok = ~(np.isnan(xs) | np.isnan(ys))
            xs, ys = xs[ok], ys[ok]
            r_val, p_val = stats.pearsonr(xs, ys)
            cells.append(
                {
                    "drinking": dname,
                    "cognition": cname,
                    "r": round(float(r_val), 3),
                    "P": round(float(p_val), 5) if p_val >= 1e-5 else 0.0,
                    "n": int(m),
                }
            )
            raw_p.append(float(p_val))
    order = np.argsort(raw_p)
    m_all = len(raw_p)
    q_vals = [0.0] * m_all
    prev = 1.0
    for rank, idx in enumerate(reversed(order)):
        i = m_all - rank
        val = min(prev, raw_p[idx] * m_all / i)
        q_vals[idx] = val
        prev = val
    for cell, qv in zip(cells, q_vals):
        cell["q_bh"] = round(min(qv, 1.0), 3)
    sig_nominal = sum(1 for c in cells if c["P"] < 0.05)
    sig_fdr = sum(1 for c in cells if c["q_bh"] < 0.05)
    return {
        "cells": cells,
        "n_tests": m_all,
        "nominal_sig_count": sig_nominal,
        "fdr_sig_count": sig_fdr,
        "min_q": round(min(q_vals), 3),
    }


def levene_for_post(rows: list[dict]) -> dict:
    out = {}
    for label in OUTCOMES:
        c_post = COG_BASE[label][1]
        groups = [col(rows, c_post, g) for g in GROUPS]
        lev = stats.levene(*groups)
        out[label] = round(float(lev.pvalue), 3)
    return out


def main() -> None:
    headers, rows = load_compare()
    cleaned, audit = clean_compare_rows(rows)

    result = {
        "ancova": {label: ancova(cleaned, COG_BASE[label][0], COG_BASE[label][1]) for label in OUTCOMES},
        "change_sd": change_sd(cleaned),
        "effect_sizes_post_ci": contrast_d_by_outcome(cleaned, use_change=False),
        "effect_sizes_change_ci": contrast_d_by_outcome(cleaned, use_change=True),
        "correlation_matrix": full_correlation(cleaned),
        "levene_post": levene_for_post(cleaned),
        "total_n": len(cleaned),
    }

    OUT_PATH.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Saved:", OUT_PATH)
    print("\n== ANCOVA (post ~ group + baseline) ==")
    for label, r in result["ancova"].items():
        print(f"{label}: F({r['df1']},{r['df2']})={r['F']}, P={r['P']}, partial eta2={r['partial_eta_sq']}, "
              f"slope-homog P={r['slope_homogeneity_P']}, adjMeans={r['adjusted_means']}")
    print("\n== Effect sizes (post, d [95% CI]) ==")
    for label, es in result["effect_sizes_post_ci"].items():
        s = es["active_vs_sham"]
        c = es["active_vs_control"]
        print(f"{label}: vs sham d={s['d']} [{s['ci_lo']},{s['ci_hi']}], vs control d={c['d']} [{c['ci_lo']},{c['ci_hi']}]")
    print("\n== Correlations ==")
    cm = result["correlation_matrix"]
    print(f"{cm['n_tests']} tests; nominal P<0.05: {cm['nominal_sig_count']}; BH-q<0.05: {cm['fdr_sig_count']}; min q={cm['min_q']}")
    for cell in cm["cells"]:
        flag = "*" if cell["P"] < 0.05 else ""
        print(f"  {cell['drinking']} x {cell['cognition']}: r={cell['r']}, P={cell['P']}, q={cell['q_bh']}{flag}")
    print("\n== Levene P (post) ==", result["levene_post"])


if __name__ == "__main__":
    main()
