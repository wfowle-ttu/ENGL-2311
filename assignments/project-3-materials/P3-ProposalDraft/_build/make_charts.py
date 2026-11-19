"""Rebuild the two survey charts from the case-study spreadsheets.

Run from anywhere:  python _build/make_charts.py
Writes fig-patient-ratings.png and fig-employee-ratings.png next to the .qmd.
"""
from collections import Counter
from pathlib import Path

import matplotlib.pyplot as plt
import openpyxl
from matplotlib import font_manager

HERE = Path(__file__).resolve().parent
OUT = HERE.parent
DATA = OUT.parent  # project-3-materials

# Palette validated with the dataviz validator (light, white surface): all checks pass.
BLUE, ORANGE = "#2F6DB5", "#C9772A"
INK, MUTED, GRID = "#1F2933", "#5B6670", "#E3E6EA"

for f in ("calibri.ttf", "calibrib.ttf"):
    p = Path("C:/Windows/Fonts") / f
    if p.exists():
        font_manager.fontManager.addfont(str(p))
plt.rcParams.update({
    "font.family": "Calibri",
    "font.size": 10,
    "text.color": INK,
    "axes.labelcolor": INK,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
})


def likert_counts(values):
    c = Counter(v for v in values if v is not None)
    return [c.get(k, 0) for k in range(1, 6)]


def load_patients():
    ws = openpyxl.load_workbook(DATA / "P3 Spoke City Patient Responses-updated.xlsx", data_only=True).active
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
    return likert_counts(r[5] for r in rows), likert_counts(r[6] for r in rows)


def load_employees():
    ws = openpyxl.load_workbook(DATA / "P3 Spoke City Employee Survey Responses.xlsx", data_only=True)["Full Survey Responses"]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
    return likert_counts(r[5] for r in rows), likert_counts(r[25] for r in rows)


def grouped_bars(a, b, labels, ylabel, xlabel, path):
    fig, ax = plt.subplots(figsize=(6.5, 3.3), dpi=300)
    x = range(1, 6)
    w = 0.36
    gap = 0.02
    bars_a = ax.bar([i - w / 2 - gap / 2 for i in x], a, w, color=BLUE, label=labels[0], zorder=3)
    bars_b = ax.bar([i + w / 2 + gap / 2 for i in x], b, w, color=ORANGE, label=labels[1], zorder=3)
    for bars in (bars_a, bars_b):
        ax.bar_label(bars, padding=2, fontsize=9, color=INK)
    ax.set_xticks(list(x))
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    top = max(a + b)
    ax.set_ylim(0, top * 1.18)
    ax.yaxis.grid(True, color=GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color("#9AA3AB")
    ax.tick_params(axis="y", length=0)
    ax.legend(loc="upper left", frameon=False, ncol=2, fontsize=9.5, handlelength=1.2)
    fig.tight_layout()
    fig.savefig(path, facecolor="white")
    plt.close(fig)


purpose, clear = load_patients()
assert purpose == [3, 16, 20, 17, 3] and clear == [5, 16, 20, 16, 2], (purpose, clear)
assert sum(purpose) == 59 and sum(clear) == 59
assert sum(clear[:3]) == 41 and sum(purpose[:3]) == 39  # figures quoted in the draft
grouped_bars(purpose, clear, ("Understood purpose", "Instructions clear and easy"),
             "Number of patients", "Patient rating (1 = lowest, 5 = highest)",
             OUT / "fig-patient-ratings.png")

consistency, interdept = load_employees()
assert consistency == [2, 4, 19, 31, 4] and interdept == [2, 4, 15, 29, 10], (consistency, interdept)
assert sum(consistency[:3]) == 25 and sum(interdept[:3]) == 21  # figures quoted in the draft
grouped_bars(consistency, interdept, ("Consistency across instruction sources", "Interdepartmental communication"),
             "Number of employees", "Employee rating (1 = lowest, 5 = highest)",
             OUT / "fig-employee-ratings.png")
print("charts written")
