#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Figure 3 — primary post-intervention outcomes (box + individual points).

Annotates omnibus F/P/q and active-vs-sham Cohen's d (95% CI) from JSON
ground truth; individual scores from Excel via sensitivity_analyses.load_rows().
"""

from __future__ import annotations

import json
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
CTRL_FILL = "#C8C8C8"
ACTIVE_FILL = "#EAEAEA"

PRIMARY = (
    ("TMT-B", "TMT-B2", "TMT-B (s)", True),
    ("VFT", "VFT2", "VFT (words)", False),
    ("HVLT-R", "HVLT-R2", "HVLT-R (words)", False),
)
GROUPS = (
    (1, "Active", INK, ACTIVE_FILL, "s", True),
    (2, "Sham", INK, "white", "o", False),
    (3, "Control", GRAY, CTRL_FILL, "o", True),
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


def q_eq(v) -> str:
    s = fmt_p(v)
    return f"q{s}" if s.startswith("<") else f"q={s}"


def load_stats() -> tuple[dict, dict]:
    ar = json.loads(AR_PATH.read_text(encoding="utf-8"))
    en = json.loads(EN_PATH.read_text(encoding="utf-8"))
    post = {row["outcome"]: row for row in ar["tables"]["table_post"]}
    return {
        "post": post,
        "pairwise": ar["pairwise_post"],
        "fdr": ar["multiplicity"]["fdr_adjusted_p"],
    }, en


def group_values(rows: list[dict], col: str) -> dict[int, np.ndarray]:
    out: dict[int, np.ndarray] = {}
    for g, *_ in GROUPS:
        vals = [float(r[col]) for r in rows if r.get("分组") == g and r.get(col) is not None]
        out[g] = np.asarray(vals, dtype=float)
    return out


def style_box(bp, edge: str, face: str) -> None:
    for box in bp["boxes"]:
        box.set(color=edge, linewidth=1.15)
        box.set_facecolor(face)
    for med in bp["medians"]:
        med.set(color=edge, linewidth=1.4)
    for whisk in bp["whiskers"]:
        whisk.set(color=edge, linewidth=1.0)
    for cap in bp["caps"]:
        cap.set(color=edge, linewidth=1.0)
    for flier in bp["fliers"]:
        flier.set(marker="")


def add_bracket(ax, x1: float, x2: float, y: float, text: str, *, dh: float) -> None:
    ax.plot([x1, x1, x2, x2], [y, y + dh, y + dh, y], color=INK, lw=1.0, clip_on=False)
    ax.text((x1 + x2) / 2, y + dh, text, ha="center", va="bottom", fontsize=8.0, color=INK)


def panel(ax, *, letter: str, outcome: str, col: str, ylabel: str, lower_better: bool,
          rows: list[dict], stats: dict, en: dict) -> None:
    data = group_values(rows, col)
    positions = [1, 2, 3]

    for (g, _name, edge, face, _mk, _filled), xpos in zip(GROUPS, positions):
        bp = ax.boxplot(
            [data[g]],
            positions=[xpos],
            widths=0.55,
            patch_artist=True,
            showfliers=False,
            medianprops=dict(linewidth=1.4),
            whiskerprops=dict(linewidth=1.0),
            capprops=dict(linewidth=1.0),
            boxprops=dict(linewidth=1.15),
            zorder=2,
        )
        style_box(bp, edge, face)

        rng = np.random.default_rng(hash(outcome + str(g)) % (2**32))
        jitter = rng.uniform(-0.12, 0.12, size=len(data[g]))
        mk = "s" if g == 1 else "o"
        if g == 1:
            ax.scatter(xpos + jitter, data[g], s=22, marker=mk, c=INK, edgecolors=INK,
                       linewidths=0.6, alpha=0.85, zorder=3)
        elif g == 2:
            ax.scatter(xpos + jitter, data[g], s=22, marker=mk, facecolors="none",
                       edgecolors=INK, linewidths=0.8, alpha=0.85, zorder=3)
        else:
            ax.scatter(xpos + jitter, data[g], s=22, marker=mk, c=GRAY, edgecolors=GRAY,
                       linewidths=0.6, alpha=0.80, zorder=3)

    ax.set_xticks(positions)
    ax.set_xticklabels([g[1] for g in GROUPS], fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9.5)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(axis="both", labelsize=8.5)
    ax.set_xlim(0.45, 3.55)

    post = stats["post"][outcome]
    fdr_q = stats["fdr"][outcome]
    pw = stats["pairwise"][outcome]["active_vs_sham"]
    dinfo = en["effect_sizes_post_ci"][outcome]["active_vs_sham"]
    d_txt = f"d={dinfo['d']:.2f} [{dinfo['ci_lo']:.2f}, {dinfo['ci_hi']:.2f}]"
    stats_txt = (
        f"F={fmt_f(post['F'])}, {p_eq(post['P'])}, {q_eq(fdr_q)}\n"
        f"active vs sham: {d_txt}"
    )
    # Stats sit above the axes so they never collide with points/boxes.
    ax.text(0.0, 1.18, stats_txt, transform=ax.transAxes, ha="left", va="bottom",
            fontsize=7.4, color=INK, linespacing=1.25, clip_on=False)

    if lower_better:
        ax.text(1.0, 1.18, "lower is better", transform=ax.transAxes,
                ha="right", va="bottom", fontsize=7.2, style="italic", color=GRAY,
                clip_on=False)

    ymax = max(float(np.max(v)) for v in data.values())
    ymin = min(float(np.min(v)) for v in data.values())
    pad = 0.12 * (ymax - ymin if ymax > ymin else 1.0)
    bracket_y = ymax + pad * 0.55
    add_bracket(ax, 1.0, 2.0, bracket_y, p_eq(pw), dh=pad * 0.35)
    ax.set_ylim(ymin - pad * 0.6, bracket_y + pad * 1.35)

    ax.text(-0.14, 1.32, letter, transform=ax.transAxes, ha="left", va="bottom",
            fontsize=12, fontweight="bold", color=INK, clip_on=False)


def build_figure(rows: list[dict], stats: dict, en: dict) -> plt.Figure:
    fig, axes = plt.subplots(1, 3, figsize=(10.8, 4.8), dpi=600)
    letters = "ABC"
    for ax, letter, (outcome, col, ylabel, lower) in zip(axes, letters, PRIMARY):
        panel(ax, letter=letter, outcome=outcome, col=col, ylabel=ylabel,
              lower_better=lower, rows=rows, stats=stats, en=en)

    legend = [
        Line2D([0], [0], marker="s", color="none", markerfacecolor=INK,
               markeredgecolor=INK, markersize=7, label="Active"),
        Line2D([0], [0], marker="o", color="none", markerfacecolor="none",
               markeredgecolor=INK, markersize=7, label="Sham"),
        Line2D([0], [0], marker="o", color="none", markerfacecolor=GRAY,
               markeredgecolor=GRAY, markersize=7, label="Control"),
    ]
    fig.legend(handles=legend, loc="upper center", ncol=3, frameon=False,
               fontsize=8.5, bbox_to_anchor=(0.5, 1.01))
    fig.subplots_adjust(left=0.06, right=0.99, top=0.78, bottom=0.11, wspace=0.32)
    return fig


def main() -> None:
    rows = load_rows()
    assert len(rows) == 77
    stats, en = load_stats()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fig = build_figure(rows, stats, en)
    png = OUT_DIR / "Figure3_primary_outcomes_post.png"
    pdf = OUT_DIR / "Figure3_primary_outcomes_post.pdf"
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
