#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CONSORT 2010 three-arm participant flow diagram (corrected accounting).

Data source: 04_data/processed/consort_flow.json (verified counts).
Corrected analysis-stage accounting (see 03_analysis/sensitivity_analyses.py,
SHAM_ONLY set, and data_cleaning_log.json):
  - active arm as filed: 28 allocated - 1 discontinued = 27 completing records
  - 5 erroneous active-arm records removed (4 cross-arm entries whose valid
    copies already sit in the sham file + 1 duplicate row)
  - 1 further record (P-E) reassigned from active to sham per the original
    randomization log -> active 21, sham 26 + 1 = 27, control 29 (total 77).

Style: black-on-white CONSORT layout, Times New Roman, vector PDF + 600 dpi PNG.
No footnote text is baked into the image (manuscript legend carries it).
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
from matplotlib import font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle

ROOT = Path(__file__).resolve().parents[1]
FLOW_PATH = ROOT / "04_data" / "processed" / "consort_flow.json"
OUT_DIR = ROOT / "01_论文正文" / "figures"

INK = "#000000"

# ---------------------------------------------------------------- fonts ----
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


def load_flow() -> dict:
    data = json.loads(FLOW_PATH.read_text(encoding="utf-8"))
    f = data["flow"]
    # Guardrails: displayed numbers must match the verified JSON.
    assert f["enrolled_assessed"] == f["randomized"] == 85
    assert sum(f["allocated"].values()) == f["randomized"]
    d = f["discontinued_intervention"]
    assert (d["control"], d["active_hd_tdcs"], d["sham_hd_tdcs"]) == (0, 1, 2)
    a = f["analyzed"]
    assert (a["control"], a["active_hd_tdcs"], a["sham_hd_tdcs"]) == (29, 21, 27)
    assert a["control"] + a["active_hd_tdcs"] + a["sham_hd_tdcs"] == a["total"] == 77
    return data


def box(ax, cx, cy, w, h, text, *, fontsize=8.0, weight="normal"):
    ax.add_patch(
        Rectangle((cx - w / 2, cy - h / 2), w, h, linewidth=1.0,
                  edgecolor=INK, facecolor="white", zorder=2)
    )
    ax.text(cx, cy, text, ha="center", va="center", fontsize=fontsize,
            color=INK, fontweight=weight, linespacing=1.32, zorder=3)
    return {"cx": cx, "cy": cy, "w": w, "h": h,
            "top": cy + h / 2, "bottom": cy - h / 2,
            "left": cx - w / 2, "right": cx + w / 2}


def arrow(ax, x1, y1, x2, y2, *, connstyle=None, lw=1.0):
    ax.add_patch(
        FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                        mutation_scale=9, linewidth=lw, color=INK,
                        shrinkA=0, shrinkB=0, joinstyle="miter",
                        connectionstyle=connstyle or "arc3,rad=0", zorder=1)
    )


def build_figure(data: dict) -> plt.Figure:
    f = data["flow"]
    disc = f["discontinued_intervention"]
    ana = f["analyzed"]

    fig, ax = plt.subplots(figsize=(7.6, 7.0), dpi=600)
    ax.set_xlim(-0.90, 10.6)
    ax.set_ylim(3.78, 13.15)
    ax.set_aspect("equal")
    ax.axis("off")

    xs = (2.05, 4.90, 7.75)          # control / active / sham columns
    aw = 2.60                        # arm-box width
    cw = 3.00                        # single-column box width
    mcx = xs[1]                      # main spine centred on active column

    # --- Enrollment -------------------------------------------------------
    assessed = box(ax, mcx, 12.55, cw, 0.66,
                   f"Assessed for eligibility\n(n={f['enrolled_assessed']})", fontsize=9)
    excluded = box(ax, 8.95, 12.55, 2.70, 0.66,
                   f"Excluded before\nrandomization (n={f['excluded_before_randomization']})",
                   fontsize=8.2)
    arrow(ax, assessed["right"], 12.55, excluded["left"], 12.55)

    randomized = box(ax, mcx, 11.35, 2.60, 0.58, f"Randomized (n={f['randomized']})",
                     fontsize=9, weight="bold")
    arrow(ax, mcx, assessed["bottom"], mcx, randomized["top"])

    # --- Allocation -------------------------------------------------------
    rail_y = 10.80
    arrow(ax, mcx, randomized["bottom"], mcx, rail_y)
    ax.plot([xs[0], xs[2]], [rail_y, rail_y], color=INK, lw=1.0,
            solid_capstyle="butt", zorder=1)

    alloc_texts = [
        "Allocated to routine\ncare (control)\n"
        f"(n={f['allocated']['control']})\n"
        f"Received intervention\nas allocated (n={f['allocated']['control']})",
        "Allocated to active\nHD-tDCS\n"
        f"(n={f['allocated']['active_hd_tdcs']})\n"
        f"Received intervention\nas allocated (n={f['allocated']['active_hd_tdcs']})",
        "Allocated to sham\nHD-tDCS\n"
        f"(n={f['allocated']['sham_hd_tdcs']})\n"
        f"Received intervention\nas allocated (n={f['allocated']['sham_hd_tdcs']})",
    ]
    alloc_top = rail_y - 0.28
    alloc_h = 1.34
    allocs = [box(ax, x, alloc_top - alloc_h / 2, aw, alloc_h, t, fontsize=8.0)
              for x, t in zip(xs, alloc_texts)]

    # --- Follow-up --------------------------------------------------------
    disc_lines = {
        0: "Discontinued intervention\nbefore completing\n14 sessions (n=0)",
        1: "Discontinued intervention\nbefore completing\n14 sessions (n=1;\nearly hospital discharge)",
        2: "Discontinued intervention\nbefore completing\n14 sessions (n=2;\nearly hospital discharge)",
    }
    disc_n = [disc["control"], disc["active_hd_tdcs"], disc["sham_hd_tdcs"]]
    follow_texts = [
        "Follow-up\n" + disc_lines[n] + "\nLost to follow-up (n=0)" for n in disc_n
    ]
    follow_top = alloc_top - alloc_h - 0.26
    follow_h = 1.52
    follows = [box(ax, x, follow_top - follow_h / 2, aw, follow_h, t, fontsize=7.6)
               for x, t in zip(xs, follow_texts)]
    for x, al in zip(xs, allocs):
        arrow(ax, x, al["bottom"], x, follow_top)

    # --- Analysis ---------------------------------------------------------
    asfiled_n = (f["allocated"]["control"] - disc["control"],
                 f["allocated"]["active_hd_tdcs"] - disc["active_hd_tdcs"],
                 f["allocated"]["sham_hd_tdcs"] - disc["sham_hd_tdcs"])
    af_top = follow_top - follow_h - 0.26
    af_h = 0.70
    asfileds = [
        box(ax, x, af_top - af_h / 2, aw, af_h,
            f"Completing records\nas filed (n={n})", fontsize=7.8)
        for x, n in zip(xs, asfiled_n)
    ]
    for x, fo in zip(xs, follows):
        arrow(ax, x, fo["bottom"], x, af_top)

    ver_cx, ver_w, ver_h = xs[1], 3.00, 1.30
    ver_top = af_top - af_h - 0.26
    verify = box(
        ax, ver_cx, ver_top - ver_h / 2, ver_w, ver_h,
        "Data verification:\n5 erroneous active-arm\nrecords removed\n(4 cross-arm entries,\n1 duplicate row)",
        fontsize=7.4,
    )

    ana_top = ver_top - ver_h - 0.26
    ana_h = 0.62
    analyzed_boxes = [
        box(ax, x, ana_top - ana_h / 2, aw, ana_h,
            f"Analyzed (n={ana[k]})", fontsize=8.4, weight="bold")
        for x, k in zip(xs, ("control", "active_hd_tdcs", "sham_hd_tdcs"))
    ]

    # control: straight through; active: through verification box;
    # sham: straight down (its analyzed count includes the reassigned record).
    arrow(ax, xs[0], asfileds[0]["bottom"], xs[0], ana_top)
    arrow(ax, xs[1], asfileds[1]["bottom"], xs[1], ver_top)
    arrow(ax, xs[1], verify["bottom"], xs[1] + 0.55, ana_top)
    arrow(ax, xs[2], asfileds[2]["bottom"], xs[2], ana_top)

    # Reassignment: labelled arrow from the verification step into the sham arm.
    re_entry_x = xs[2] + 0.95           # enters sham 'Analyzed' box beside main arrow
    arrow(ax, verify["right"], ver_top - ver_h / 2, re_entry_x, ana_top,
          connstyle="arc3,rad=-0.22")
    ax.text(xs[2] + 1.02, ver_top - 0.08,
            "1 record reassigned\nto sham per the\nrandomization log",
            ha="left", va="top", fontsize=7.4, style="italic", color=INK)

    # --- Stage labels (left rail) -----------------------------------------
    stages = {
        "Enrollment": (assessed["cy"] + randomized["cy"]) / 2,
        "Allocation": allocs[0]["cy"],
        "Follow-up": follows[0]["cy"],
        "Analysis": (af_top - af_h / 2 + ana_top - ana_h / 2) / 2,
    }
    for name, sy in stages.items():
        ax.text(-0.82, sy, name, ha="left", va="center", fontsize=9.5,
                fontweight="bold", color=INK)

    fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)
    return fig


def main() -> None:
    data = load_flow()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fig = build_figure(data)
    png = OUT_DIR / "Figure1_CONSORT_flowchart.png"
    pdf = OUT_DIR / "Figure1_CONSORT_flowchart.pdf"
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
