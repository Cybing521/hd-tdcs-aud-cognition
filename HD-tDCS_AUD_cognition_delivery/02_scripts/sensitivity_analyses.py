#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sensitivity analyses for the post-randomization data-verification step.

Sets (unique participants only):
  S1_primary        cleaned analysis set as published (21/27/29); P-E in sham
  S2_lei_as_active  same 77 participants; P-E's retained record analyzed in active (22/26/29)
  S3_lei_excluded   P-E excluded entirely (21/26/29)
  S4_as_recorded    all five ambiguous participants analyzed in active (26/22/29)

For each set and primary outcome (post-intervention TMT-B, VFT, HVLT-R):
  - Levene P; omnibus ANOVA (classic or Welch per Levene) F, P
  - ANCOVA post ~ group + baseline: group F, P (type I/II identical for this design)
  - active-vs-sham pairwise: mean diff, P (pooled t), Cohen's d with large-sample 95% CI, Hedges g
  - Kruskal-Wallis H, P (rank-based robustness)
Also: minimum detectable effect (Cohen's d) at 80% power, n=21 vs 27, alpha 0.05 two-sided.
"""

from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

import numpy as np
import openpyxl
from scipy import stats

ROOT = Path(__file__).resolve().parents[1]
TABLE_DIR = ROOT / "03_分析表格"
OUT_PATH = ROOT / "04_data" / "processed" / "sensitivity_analyses.json"

SHAM_ONLY = {"P-A", "P-B", "P-C", "P-D", "P-E"}

PRIMARY = {
    "TMT-B": ("TMT-B1", "TMT-B2"),   # lower post score = better
    "VFT": ("VFT1", "VFT2"),
    "HVLT-R": ("HVLT-R1", "HVLT-R2"),
}


def load_rows() -> list[dict[str, Any]]:
    """Load the cleaned per-arm tables (真/伪/对照 比较 files).

    The pre-cleaning workbook was overwritten by the cleaning pipeline, so sets are
    rebuilt from the cleaned unique records: S1 as published; S2/S3 move/drop P-E;
    S4 places all five ambiguous participants in the active arm using their retained
    records (worst-case bound for the active-vs-sham contrast).
    """
    rows: list[dict[str, Any]] = []
    for fname, grp in (("真刺激组比较.xlsx", 1), ("伪刺激组比较.xlsx", 2), ("酒依赖比较.xlsx", 3)):
        wb = openpyxl.load_workbook(TABLE_DIR / fname, data_only=True)
        ws = wb.active
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            rec = dict(zip(headers, row))
            rec["分组"] = grp
            rows.append(rec)
        wb.close()
    return rows


def build_sets(rows: list[dict[str, Any]]) -> dict[str, dict[int, list[dict]]]:
    lei_rows = [r for r in rows if r.get("姓名") == "P-E"]
    lei_rows.sort(key=lambda x: (x.get("年龄") != 38, x.get("年龄") or 999))
    lei = dict(lei_rows[0]) if lei_rows else None

    def without_ambiguous(recs: list[dict]) -> list[dict]:
        return [r for r in recs if r.get("姓名") not in SHAM_ONLY]

    sets: dict[str, dict[int, list[dict]]] = {}
    sets["S1_primary"] = group_by([dict(r) for r in rows])

    s2 = [dict(r) for r in rows if r.get("姓名") != "P-E"]
    if lei is not None:
        r = dict(lei)
        r["分组"] = 1
        s2.append(r)
    sets["S2_lei_as_active"] = group_by(s2)

    s3 = [dict(r) for r in rows if r.get("姓名") != "P-E"]
    sets["S3_lei_excluded"] = group_by(s3)

    s4 = without_ambiguous(rows)
    for r in rows:
        if r.get("姓名") in SHAM_ONLY and r.get("姓名") != "P-E":
            rec = dict(r)
            rec["分组"] = 1
            s4.append(rec)
    if lei is not None:
        rec = dict(lei)
        rec["分组"] = 1
        s4.append(rec)
    sets["S4_as_recorded"] = group_by(s4)

    return sets


def group_by(rows: list[dict]) -> dict[int, list[dict]]:
    g: dict[int, list[dict]] = {1: [], 2: [], 3: []}
    for r in rows:
        g[r.get("分组")].append(r)
    return g


def col(rows: list[dict], name: str) -> np.ndarray:
    return np.array([float(r[name]) for r in rows if r.get(name) is not None], dtype=float)


def cohens_d(a: np.ndarray, b: np.ndarray) -> tuple[float, float, float]:
    """d = (m1-m2)/sd_pooled with large-sample 95% CI; also Hedges g."""
    n1, n2 = len(a), len(b)
    sp = math.sqrt(((n1 - 1) * a.var(ddof=1) + (n2 - 1) * b.var(ddof=1)) / (n1 + n2 - 2))
    d = (a.mean() - b.mean()) / sp
    se = math.sqrt((n1 + n2) / (n1 * n2) + d * d / (2 * (n1 + n2 - 2)))
    lo, hi = d - 1.96 * se, d + 1.96 * se
    j = 1 - 3 / (4 * (n1 + n2) - 9)
    return d, lo, hi, d * j


def ancova_group_test(post: list[np.ndarray], base_all: np.ndarray, grp: np.ndarray) -> tuple[float, float, int, int]:
    """ANCOVA post ~ group + baseline via OLS; return group F, P, df1, df2."""
    y = np.concatenate(post)
    x_grp = np.zeros(len(y))
    base = np.zeros(len(y))
    # design: intercept + two group dummies + baseline (centered not required)
    n = len(y)
    g = grp
    b = base_all
    X = np.column_stack([np.ones(n), (g == 1).astype(float), (g == 2).astype(float), b])
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    resid = y - X @ beta
    df2 = n - X.shape[1]
    mse = (resid @ resid) / df2
    # F test for the two group dummies jointly
    X0 = np.column_stack([np.ones(n), b])
    beta0, *_ = np.linalg.lstsq(X0, y, rcond=None)
    resid0 = y - X0 @ beta0
    f = ((resid0 @ resid0 - resid @ resid) / 2) / mse
    p = float(stats.f.sf(f, 2, df2))
    return float(f), p, 2, int(df2)


def analyze_set(name: str, sets: dict) -> dict:
    g = sets[name]
    ns = {k: len(v) for k, v in g.items()}
    res: dict[str, Any] = {"n_active": ns[1], "n_sham": ns[2], "n_control": ns[3]}
    for outcome, (bcol, pcol) in PRIMARY.items():
        a = col(g[1], pcol)   # active post
        s = col(g[2], pcol)   # sham post
        c = col(g[3], pcol)   # control post
        ab = col(g[1], bcol)
        sb = col(g[2], bcol)
        cb = col(g[3], bcol)

        lev_stat, lev_p = stats.levene(a, s, c, center="median")
        if lev_p < 0.05:
            w = stats.anova_oneway([a, s, c], use_var="welch")
            f, p, method = float(w.statistic), float(w.pvalue), "Welch"
        else:
            f, p = float(stats.f_oneway(a, s, c).statistic), float(stats.f_oneway(a, s, c).pvalue)
            method = "classic"

        kw_h, kw_p = stats.kruskal(a, s, c)

        post = [a, s, c]
        base_all = np.concatenate([ab, sb, cb])
        grp = np.concatenate([np.full(len(x), k) for k, x in zip((1, 2, 3), post)])
        af, ap, df1, df2 = ancova_group_test(post, base_all, grp)

        t, tp = stats.ttest_ind(a, s, equal_var=True)
        d, dlo, dhi, gval = cohens_d(a, s)

        res[outcome] = {
            "levene_P": round(float(lev_p), 4),
            "omnibus_method": method,
            "omnibus_F": round(f, 3),
            "omnibus_P": round(p, 4) if p >= 0.0001 else "<0.001",
            "ancova_F": round(af, 3),
            "ancova_P": round(ap, 4) if ap >= 0.0001 else "<0.001",
            "ancova_df": f"{df1},{df2}",
            "kw_H": round(float(kw_h), 3),
            "kw_P": round(float(kw_p), 4) if kw_p >= 0.0001 else "<0.001",
            "active_vs_sham_mean_diff": round(float(a.mean() - s.mean()), 2),
            "active_vs_sham_P": round(float(tp), 4) if tp >= 0.0001 else "<0.001",
            "cohen_d": round(d, 2),
            "d_CI": [round(dlo, 2), round(dhi, 2)],
            "hedges_g": round(gval, 2),
        }
    return res


def mde() -> dict:
    """Minimum detectable Cohen's d at 80% power, two-sided alpha .05, 21 vs 27."""
    from statsmodels.stats.power import TTestIndPower

    pw = TTestIndPower()
    d = pw.solve_power(effect_size=None, nobs1=21, ratio=27 / 21, alpha=0.05, power=0.8)
    return {"min_detectable_d_80power_21v27": round(float(d), 2)}


def main() -> None:
    rows = load_rows()
    sets = build_sets(rows)
    out = {name: analyze_set(name, sets) for name in sets}
    out["minimum_detectable_effect"] = mde()
    out["note"] = (
        "S1 = published cleaned set; S2 = P-E analyzed in active; S3 = P-E excluded; "
        "S4 = all five ambiguous participants analyzed in active (as-recorded arms, "
        "randomization log disregarded). TMT-B lower is better; negative d favors active on TMT-B."
    )
    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=1))
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
