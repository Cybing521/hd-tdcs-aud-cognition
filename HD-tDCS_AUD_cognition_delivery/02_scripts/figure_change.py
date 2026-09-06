#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Figure 4 — baseline→post change trajectories for primary outcomes.

Mean ± 95% CI with group connecting lines; thin individual trajectories;
annotates change-score ANOVA F/P and ANCOVA partial η² from JSON.
"""

from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
from matplotlib import font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.lines import Line2D

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "03_analysis"))
from sensitivity_analyses import load_rows  # noqa: E402

AR_PATH = ROOT / "04_data" / "processed" / "analysis_results.json"
EN_PATH = ROOT / "04_data" / "processed" / "enhanced_analyses.json"
OUT_DIR = ROOT / "01_论文正文" / "figures"

INK = "#000000"
GRAY = "#666666"
TRAJ = "#9A9A9A"

PRIMARY = (
    ("TMT-B", "TMT-B1", "TMT-B2", "TMT-B (s)", True),
    ("VFT", "VFT1", "VFT2", "VFT (words)", False),
    ("HVLT-R", "HVLT-R1", "HVLT-R2", "HVLT-R (words)", False),
)
# group_id, label, color, marker, filled, linestyle
GROUPS = (
    (1, "Active", INK, "s", True, "-"),
    (2, "Sham", INK, "o", False, "--"),
    (3, "Control", GRAY, "o", True, ":"),
)

_FONT_DIR = Path("/System/Library/Fonts/Supplemental")
for _name in (
    "Times New Roman.ttf",
    "Times New Roman Bold.ttf",
    "Times New Roman Italic.ttf",
    "Times New Roman Bold Italic.ttf",
):
    _p = _FONT_DIR / _name
    if _p.exists():
        fm.fontManager.addfont(str(_p))
matplotlib.rcParams.update(
    {
        "font.family": "Times New Roman",
        "font.size": 9,
        "pdf.fonttype": 42,
        "ps.fonttype": 42,
        "axes.unicode_minus": False,
    }
)
_resolved = fm.findfont("Times New Roman", fallback_to_default=False)
assert "times" in Path(_resolved).name.lower(), f"TNR not loaded, got {_resolved}"
assert "dejavu" not in Path(_resolved).name.lower(), f"DejaVu fallback: {_resolved}"


def fmt_p(v) -> str:
    if v is None:
        return "—"
    v = float(v)
    if v < 0.001:
        return "<0.001"
    return f"{v:.3f}"


def fmt_f(v) -> str:
    return f"{float(v):.2f}"


def p_eq(v) -> str:
    s = fmt_p(v)
    return f"P{s}" if s.startswith("<") else f"P={s}"


def mean_ci(vals: np.ndarray) -> tuple[float, float, float]:
    n = len(vals)
    m = float(vals.mean())
    se = float(vals.std(ddof=1) / math.sqrt(n))
    # large-sample 95% CI (matches manuscript d-CI convention)
    return m, m - 1.96 * se, m + 1.96 * se


def load_stats() -> tuple[dict, dict]:
    ar = json.loads(AR_PATH.read_text(encoding="utf-8"))
    en = json.loads(EN_PATH.read_text(encoding="utf-8"))
    change = {row["outcome"]: row for row in ar["change_score_anova"]}
    return change, en["ancova"]


def panel(ax, *, letter: str, outcome: str, c1: str, c2: str, ylabel: str,
          lower_better: bool, rows: list[dict], change: dict, ancova: dict) -> None:
    xs = np.array([0.0, 1.0])

    # Individual trajectories (behind group means)
    for g, _lab, color, _mk, _filled, _ls in GROUPS:
        for r in rows:
            if r.get("分组") != g:
                continue
            if r.get(c1) is None or r.get(c2) is None:
                continue
            ax.plot(xs, [float(r[c1]), float(r[c2])], color=TRAJ, lw=0.7,
                    alpha=0.15, zorder=1)

    # Group mean ± 95% CI
    for g, lab, color, mk, filled, ls in GROUPS:
        base = np.asarray(
            [float(r[c1]) for r in rows if r.get("分组") == g and r.get(c1) is not None],
            dtype=float,
        )
        post = np.asarray(
            [float(r[c2]) for r in rows if r.get("分组") == g and r.get(c2) is not None],
            dtype=float,
        )
        m0, lo0, hi0 = mean_ci(base)
        m1, lo1, hi1 = mean_ci(post)
        ax.plot(xs, [m0, m1], color=color, lw=1.8, ls=ls, zorder=3, label=lab)
        yerr0 = np.array([[m0 - lo0], [hi0 - m0]])
        yerr1 = np.array([[m1 - lo1], [hi1 - m1]])
        face = color if filled else "none"
        ax.errorbar([0], [m0], yerr=yerr0, fmt="none", ecolor=color, elinewidth=1.1,
                    capsize=3.5, capthick=1.1, zorder=4)
        ax.errorbar([1], [m1], yerr=yerr1, fmt="none", ecolor=color, elinewidth=1.1,
                    capsize=3.5, capthick=1.1, zorder=4)
        ax.scatter([0, 1], [m0, m1], s=42, marker=mk, facecolors=face,
                   edgecolors=color, linewidths=1.2, zorder=5)

    ax.set_xticks([0, 1])
    ax.set_xticklabels(["Baseline", "Post"], fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9.5)
    ax.set_xlim(-0.25, 1.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(axis="both", labelsize=8.5)

    ch = change[outcome]
    an = ancova[outcome]
    stats_txt = (
        f"Δ ANOVA: F={fmt_f(ch['F'])}, {p_eq(ch['P'])}\n"
        f"ANCOVA: partial η²={an['partial_eta_sq']:.3f}"
    )
    # Keep annotations above the plot area to avoid trajectory overlap.
    ax.text(0.0, 1.18, stats_txt, transform=ax.transAxes, ha="left", va="bottom",
            fontsize=7.4, color=INK, linespacing=1.25, clip_on=False)
    if lower_better:
        ax.text(1.0, 1.18, "lower is better", transform=ax.transAxes,
                ha="right", va="bottom", fontsize=7.2, style="italic", color=GRAY,
                clip_on=False)

    ax.text(-0.14, 1.32, letter, transform=ax.transAxes, ha="left", va="bottom",
            fontsize=12, fontweight="bold", color=INK, clip_on=False)


def build_figure(rows: list[dict], change: dict, ancova: dict) -> plt.Figure:
    fig, axes = plt.subplots(1, 3, figsize=(10.8, 4.8), dpi=600)
    letters = "ABC"
    for ax, letter, (outcome, c1, c2, ylabel, lower) in zip(axes, letters, PRIMARY):
        panel(ax, letter=letter, outcome=outcome, c1=c1, c2=c2, ylabel=ylabel,
              lower_better=lower, rows=rows, change=change, ancova=ancova)

    legend = [
        Line2D([0], [0], color=INK, lw=1.8, ls="-", marker="s", markersize=6,
               markerfacecolor=INK, markeredgecolor=INK, label="Active"),
        Line2D([0], [0], color=INK, lw=1.8, ls="--", marker="o", markersize=6,
               markerfacecolor="none", markeredgecolor=INK, label="Sham"),
        Line2D([0], [0], color=GRAY, lw=1.8, ls=":", marker="o", markersize=6,
               markerfacecolor=GRAY, markeredgecolor=GRAY, label="Control"),
    ]
    fig.legend(handles=legend, loc="upper center", ncol=3, frameon=False,
               fontsize=8.5, bbox_to_anchor=(0.5, 1.01))
    fig.subplots_adjust(left=0.06, right=0.99, top=0.78, bottom=0.11, wspace=0.32)
    return fig


def main() -> None:
    rows = load_rows()
    assert len(rows) == 77
    change, ancova = load_stats()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fig = build_figure(rows, change, ancova)
    png = OUT_DIR / "Figure4_change_trajectories.png"
    pdf = OUT_DIR / "Figure4_change_trajectories.pdf"
    fig.savefig(png, dpi=600, facecolor="white", bbox_inches="tight", pad_inches=0.08)
    fig.savefig(pdf, facecolor="white", bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)

    from PIL import Image
    with Image.open(png) as im:
        w, h = im.size
    assert w >= 2400, f"PNG too narrow: {w}px"
    print(f"Wrote {png} ({w}x{h}px)")
    print(f"Wrote {pdf}")


if __name__ == "__main__":
    main()
