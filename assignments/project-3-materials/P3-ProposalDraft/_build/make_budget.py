"""Budget Allocation: one source of truth for the table and the chart.

Run from anywhere:  python _build/make_budget.py
Writes _build/budget-table.md (included by the .qmd) and fig-budget.png.
All rates and hours are planning estimates, not quotes.
"""
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib import font_manager

HERE = Path(__file__).resolve().parent
OUT = HERE.parent

TOTAL = 100_000
CLINICAL_STAFF = 44  # "About the Clinic": 44 clinical staff

# (category, line item, basis shown in the table, amount)
ITEMS = [
    ("Shared application", "EHR communication module (license and setup, first year)", "Estimated annual license", 24_000),
    ("Shared application", "IT specialist configuration and integration", "160 hrs × $115/hr", 160 * 115),
    ("Shared application", "Security and privacy (HIPAA) review", "Flat fee", 5_000),
    ("Patient communication", "Staff training: teach-back and the new application",
     f"{CLINICAL_STAFF} staff × 4 hrs × $45/hr + trainer", CLINICAL_STAFF * 4 * 45 + 2_080),
    ("Patient communication", "Exercise demonstration videos and updated handouts", "Production + therapist time", 6_000),
    ("Patient communication", "Early in-person follow-up visits (added therapist time)", "180 hrs × $50/hr", 180 * 50),
    # Facility: items named in patient feedback cards 2, 8, 11, 12 and employee survey Q18/Q19 comments.
    ("Facility and equipment", "Replace and service worn therapy equipment",
     "2 bikes $3,600; bands, pads, balance boards $1,200; machine repair $1,200; inspection $1,000",
     EQUIP := 3_600 + 1_200 + 1_200 + 1_000),
    ("Facility and equipment", "Minor renovations: therapy gym walls, lighting, privacy partitions",
     "Estimated contractor work", 3_000),
    ("Consulting team", "Our team: analysis, implementation support, evaluation", "140 hrs × $75/hr", 140 * 75),
    ("Evaluation", "Follow-up patient and employee surveys", "Survey tools + staff time", 2_100),
    ("Contingency", "Contingency reserve (5%)", "5% of total", 5_000),
]
CONSULT_HOURS = {"Research and analysis": 40, "Implementation support": 60, "Evaluation and final report": 40}

amounts = [a for *_, a in ITEMS]
assert sum(amounts) == TOTAL, sum(amounts)
assert round(sum(round(a / TOTAL * 100, 1) for a in amounts), 1) == 100.0  # shares add to 100.0%
assert ITEMS[3][3] == 10_000 and sum(CONSULT_HOURS.values()) == 140
assert EQUIP == 7_000 and ITEMS[-1][3] == TOTAL * 5 // 100


def money(v):
    return f"${v:,.0f}"


# ---- table (markdown, included by the qmd) ----
lines = ["| Category | Line item | Basis | Amount | Share |",
         "|:---------|:----------|:------|-------:|------:|"]
for cat, item, basis, amt in ITEMS:
    lines.append(f"| {cat} | {item} | {basis} | {money(amt)} | {amt / TOTAL:.1%} |")
lines.append(f"| **Total** | | | **{money(TOTAL)}** | **100.0%** |")
lines += ["", ': Proposed allocation of the $100,000 improvement budget. '
          '{#tbl-budget tbl-colwidths="[17,33,30,11,9]"}', ""]
(HERE / "budget-table.md").write_text("\n".join(lines), encoding="utf8")

# ---- chart: totals by category, horizontal bars, largest first ----
cats = {}
for cat, *_, amt in ITEMS:
    cats[cat] = cats.get(cat, 0) + amt
order = sorted(cats.items(), key=lambda kv: kv[1])  # barh draws bottom-up
assert sum(cats.values()) == TOTAL
print({k: v for k, v in cats.items()})

BLUE, GRAY, INK, MUTED = "#2F6DB5", "#9AA3AB", "#1F2933", "#5B6670"
for f in ("calibri.ttf", "calibrib.ttf"):
    p = Path("C:/Windows/Fonts") / f
    if p.exists():
        font_manager.fontManager.addfont(str(p))
plt.rcParams.update({"font.family": "Calibri", "font.size": 10, "text.color": INK,
                     "xtick.color": MUTED, "ytick.color": INK})

fig, ax = plt.subplots(figsize=(6.5, 2.9), dpi=300)
labels = [k for k, _ in order]
vals = [v for _, v in order]
colors = [GRAY if k == "Contingency" else BLUE for k in labels]
bars = ax.barh(labels, vals, color=colors, height=0.62, zorder=3)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 900, b.get_y() + b.get_height() / 2,
            f"{money(v)}  ({v / TOTAL:.1%})", va="center", fontsize=9.5, color=INK)
ax.set_xlim(0, max(vals) * 1.28)
ax.xaxis.set_major_formatter(lambda x, _: f"${x / 1000:.0f}k")
ax.xaxis.grid(True, color="#E3E6EA", linewidth=0.8, zorder=0)
ax.set_xlabel("Allocated amount (total $100,000)", color=INK)
for side in ("top", "right", "bottom"):
    ax.spines[side].set_visible(False)
ax.spines["left"].set_color("#9AA3AB")
ax.tick_params(axis="both", length=0)
fig.tight_layout()
fig.savefig(OUT / "fig-budget.png", facecolor="white")
print("budget table + chart written")
