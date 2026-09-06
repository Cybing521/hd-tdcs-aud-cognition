#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Figure 2 — bilateral HD-tDCS 4×1 montage schematic (B/W, Times New Roman).

Panel A: superior view with F3 anode / F4 cathode rings.
Panel B: left lateral view of the F3 4×1 ring.
"""

from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
from matplotlib import font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Circle, Ellipse, Polygon

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "01_论文正文" / "figures"

INK = "#000000"
GRAY = "#666666"
LIGHT = "#B0B0B0"

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


# Approximate 10-10 positions on a unit head (nose = +y), superior view.
SUP = {
    "Nz": (0.00, 0.92),
    "Fpz": (0.00, 0.80),
    "FP1": (-0.22, 0.72),
    "FP2": (0.22, 0.72),
    "F7": (-0.65, 0.38),
    "F3": (-0.35, 0.38),
    "Fz": (0.00, 0.42),
    "F4": (0.35, 0.38),
    "F8": (0.65, 0.38),
    "C3": (-0.50, 0.00),
    "C1": (-0.25, 0.00),
    "Cz": (0.00, 0.00),
    "C2": (0.25, 0.00),
    "C4": (0.50, 0.00),
    "T7": (-0.85, 0.00),
    "T8": (0.85, 0.00),
    "P3": (-0.40, -0.40),
    "Pz": (0.00, -0.42),
    "P4": (0.40, -0.40),
    "Oz": (0.00, -0.80),
}

ANODE = ("F3", ["FP1", "F7", "C1", "C3"])
CATHODE = ("F4", ["FP2", "F8", "C2", "C4"])
REF_LABELS = ["Fz", "Cz", "Pz", "Fpz", "Oz", "T7", "T8"]


def _draw_head_superior(ax) -> None:
    ax.add_patch(Circle((0, 0), 1.0, fill=False, lw=1.4, ec=INK, zorder=1))
    # Nose
    ax.add_patch(
        Polygon([(-0.12, 0.96), (0.0, 1.14), (0.12, 0.96)], closed=True,
                fill=False, lw=1.1, ec=INK, zorder=1)
    )
    # Ears
    ax.add_patch(Ellipse((-1.05, 0.0), 0.14, 0.32, fill=False, lw=1.0, ec=INK, zorder=1))
    ax.add_patch(Ellipse((1.05, 0.0), 0.14, 0.32, fill=False, lw=1.0, ec=INK, zorder=1))
    # Midline guide
    ax.plot([0, 0], [-0.95, 0.95], color=LIGHT, lw=0.6, ls=":", zorder=0)
    ax.plot([-0.95, 0.95], [0, 0], color=LIGHT, lw=0.6, ls=":", zorder=0)


def _elec(ax, xy, *, kind: str, label: str, label_offset=(0.0, 0.0), fontsize=7.5):
    x, y = xy
    s = 95
    if kind == "anode":
        ax.scatter([x], [y], s=s, marker="s", c=INK, edgecolors=INK,
                   linewidths=1.1, zorder=5)
    elif kind == "cathode":
        ax.scatter([x], [y], s=s, marker="s", c="white", edgecolors=INK,
                   linewidths=1.3, zorder=5)
    elif kind == "return":
        ax.scatter([x], [y], s=55, marker="o", c="white", edgecolors=INK,
                   linewidths=1.1, zorder=4)
    else:
        ax.scatter([x], [y], s=18, marker="o", c=LIGHT, edgecolors=LIGHT,
                   linewidths=0.6, zorder=2)
    ax.text(x + label_offset[0], y + label_offset[1], label,
            ha="center", va="center", fontsize=fontsize, color=INK if kind != "ref" else GRAY,
            zorder=6)


def _ring_links(ax, center: str, returns: list[str]) -> None:
    cx, cy = SUP[center]
    for r in returns:
        rx, ry = SUP[r]
        ax.plot([cx, rx], [cy, ry], color=INK, lw=0.9, alpha=0.55, zorder=3)


def panel_a(ax) -> None:
    ax.set_xlim(-1.35, 1.35)
    ax.set_ylim(-1.35, 1.30)
    ax.set_aspect("equal")
    ax.axis("off")
    _draw_head_superior(ax)

    for lab in REF_LABELS:
        _elec(ax, SUP[lab], kind="ref", label=lab, fontsize=6.5,
              label_offset=(0.0, -0.10 if lab != "Fpz" else 0.10))

    _ring_links(ax, *ANODE)
    _ring_links(ax, *CATHODE)

    for r in ANODE[1]:
        off = {"FP1": (-0.02, 0.11), "F7": (-0.12, 0.02), "C1": (0.0, -0.12),
               "C3": (-0.12, -0.02)}[r]
        _elec(ax, SUP[r], kind="return", label=r, label_offset=off, fontsize=7.2)
    for r in CATHODE[1]:
        off = {"FP2": (0.02, 0.11), "F8": (0.12, 0.02), "C2": (0.0, -0.12),
               "C4": (0.12, -0.02)}[r]
        _elec(ax, SUP[r], kind="return", label=r, label_offset=off, fontsize=7.2)

    _elec(ax, SUP["F3"], kind="anode", label="F3", label_offset=(0.0, 0.14), fontsize=8.5)
    _elec(ax, SUP["F4"], kind="cathode", label="F4", label_offset=(0.0, 0.14), fontsize=8.5)

    ax.text(-0.70, -1.12, "Anode ring (F3)", ha="center", va="center", fontsize=8.0, color=INK)
    ax.text(0.70, -1.12, "Cathode ring (F4)", ha="center", va="center", fontsize=8.0, color=INK)
    ax.text(-1.28, 1.22, "A", ha="left", va="top", fontsize=12, fontweight="bold", color=INK)

    legend = [
        Line2D([0], [0], marker="s", color="none", markerfacecolor=INK,
               markeredgecolor=INK, markersize=8, label="Anode (F3)"),
        Line2D([0], [0], marker="s", color="none", markerfacecolor="white",
               markeredgecolor=INK, markersize=8, label="Cathode (F4)"),
        Line2D([0], [0], marker="o", color="none", markerfacecolor="white",
               markeredgecolor=INK, markersize=7, label="Return electrodes"),
    ]
    ax.legend(handles=legend, loc="upper right", frameon=False, fontsize=7.5,
              handletextpad=0.4, borderaxespad=0.1)


def _draw_head_lateral(ax) -> None:
    # Left profile: nose to the right (+x), vertex up (+y).
    ax.add_patch(Ellipse((0.0, 0.05), 1.70, 1.85, fill=False, lw=1.4, ec=INK, zorder=1))
    # Nose
    ax.add_patch(
        Polygon([(0.78, 0.18), (1.12, 0.05), (0.78, -0.08)], closed=True,
                fill=False, lw=1.1, ec=INK, zorder=1)
    )
    # Ear (schematic)
    ax.add_patch(Ellipse((-0.55, -0.05), 0.22, 0.38, fill=False, lw=1.0, ec=INK, zorder=1))
    # Neck (stop above panel caption)
    ax.plot([-0.25, -0.18], [-0.85, -1.05], color=INK, lw=1.1, zorder=1)
    ax.plot([0.15, 0.12], [-0.85, -1.05], color=INK, lw=1.1, zorder=1)


# Lateral (left) approximate placements for F3 ring.
LAT = {
    "F3": (0.05, 0.42),
    "FP1": (0.55, 0.48),
    "F7": (0.48, 0.10),
    "C1": (-0.15, 0.28),
    "C3": (-0.05, -0.05),
}


def panel_b(ax) -> None:
    ax.set_xlim(-1.15, 1.30)
    ax.set_ylim(-1.45, 1.20)
    ax.set_aspect("equal")
    ax.axis("off")
    _draw_head_lateral(ax)

    cx, cy = LAT["F3"]
    for r in ("FP1", "F7", "C1", "C3"):
        rx, ry = LAT[r]
        ax.plot([cx, rx], [cy, ry], color=INK, lw=0.9, alpha=0.55, zorder=3)
        ax.scatter([rx], [ry], s=55, marker="o", c="white", edgecolors=INK,
                   linewidths=1.1, zorder=4)
        off = {"FP1": (0.14, 0.06), "F7": (0.14, -0.02), "C1": (-0.14, 0.06),
               "C3": (-0.02, -0.14)}[r]
        ax.text(rx + off[0], ry + off[1], r, ha="center", va="center",
                fontsize=7.5, color=INK, zorder=6)

    ax.scatter([cx], [cy], s=100, marker="s", c=INK, edgecolors=INK,
               linewidths=1.1, zorder=5)
    ax.text(cx - 0.02, cy + 0.16, "F3 anode", ha="center", va="bottom",
            fontsize=8.5, color=INK, zorder=6)

    ax.text(-1.05, 1.12, "B", ha="left", va="top", fontsize=12, fontweight="bold", color=INK)
    # Caption below the neck (neck ends near y=-1.15); keep clear gap from lines.
    ax.text(0.05, -1.32, "Left lateral view — F3 4×1 ring", ha="center", va="center",
            fontsize=8.0, color=INK)
    ax.annotate("", xy=(0.95, -0.55), xytext=(0.55, -0.55),
                arrowprops=dict(arrowstyle="-|>", color=GRAY, lw=0.9))
    ax.text(0.75, -0.68, "anterior", ha="center", va="top", fontsize=7.0, color=GRAY)


def build_figure() -> plt.Figure:
    fig, axes = plt.subplots(1, 2, figsize=(9.6, 5.0), dpi=600)
    panel_a(axes[0])
    panel_b(axes[1])
    fig.subplots_adjust(left=0.02, right=0.98, top=0.96, bottom=0.08, wspace=0.08)
    return fig


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fig = build_figure()
    png = OUT_DIR / "Figure2_HD_tDCS_montage.png"
    pdf = OUT_DIR / "Figure2_HD_tDCS_montage.pdf"
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
