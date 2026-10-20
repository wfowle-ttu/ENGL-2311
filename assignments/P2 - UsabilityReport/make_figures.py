# Generates the report figures (figures/*.png) from the usability test results.
# Values come from "P2-UsabilityTestResults-ParticipantInfo LSports.xlsx" and
# match the Task Completion Success Rate tables in fowle_usabilityreport.qmd.
# Run: python make_figures.py
from pathlib import Path

import matplotlib.pyplot as plt

OUT = Path(__file__).parent / "figures"
OUT.mkdir(exist_ok=True)

BLUE, ORANGE = "#2a78d6", "#eb6834"
INK, MUTED, GRID = "#0b0b0b", "#52514e", "#e4e3df"

plt.rcParams.update({
    "font.family": "Georgia",
    "font.size": 10,
    "text.color": INK,
    "axes.labelcolor": MUTED,
    "xtick.color": MUTED,
    "ytick.color": INK,
    "axes.edgecolor": GRID,
})


def style_axes(ax):
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.xaxis.grid(True, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)


# ---- Figure 1: task completion rate for every task ----
tasks = [
    ("S1 · T1  Find sports facilities section", 100),
    ("S1 · T2  Identify soccer field rentals", 83),
    ("S1 · T3  Assess field capacity", 100),
    ("S1 · T4  Check field availability", 67),
    ("S1 · T5  Find contact / booking form", 83),
    ("S1 · T6  Say \"I'm done\"", 83),
    ("S2 · T1  Start at homepage", 100),
    ("S2 · T2  Find LCU facilities", 83),
    ("S2 · T3  Find number and types of LCU venues", 83),
    ("S2 · T4  Estimate venue capacity", 83),
    ("S2 · T5  Find more facility information", 83),
    ("S2 · T6  Say \"I'm done\"", 100),
    ("S3 · T1  Start at homepage", 100),
    ("S3 · T2  Find driving distance from Austin", 50),
    ("S3 · T3  Count major highways", 83),
    ("S3 · T4  Find local transportation options", 50),
    ("S3 · T5  Find transportation contact", 100),
    ("S3 · T6  Say \"I'm done\"", 100),
]
THRESHOLD = 70
labels = [t[0] for t in tasks][::-1]
rates = [t[1] for t in tasks][::-1]
colors = [ORANGE if r < THRESHOLD else BLUE for r in rates]

fig, ax = plt.subplots(figsize=(7.5, 6.2), dpi=200)
bars = ax.barh(labels, rates, color=colors, height=0.62, edgecolor="white", linewidth=1.5)
for bar, r in zip(bars, rates):
    ax.text(bar.get_width() + 1.2, bar.get_y() + bar.get_height() / 2, f"{r}%",
            va="center", fontsize=9, color=INK)
# light separators between scenarios
for y in (5.5, 11.5):
    ax.axhline(y, color=GRID, linewidth=1)
ax.set_xlim(0, 110)
ax.set_xticks(range(0, 101, 25))
ax.set_xticklabels([f"{x}%" for x in range(0, 101, 25)])
ax.set_xlabel("Participants who completed the task (n = 6)")
style_axes(ax)
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=ORANGE, label="Below 70% completion"),
                   Patch(color=BLUE, label="70% completion or higher")],
          loc="lower center", bbox_to_anchor=(0.35, 1.0), ncol=2, frameon=False, fontsize=9)
fig.tight_layout()
fig.savefig(OUT / "fig-completion.png", bbox_inches="tight")
plt.close(fig)

# ---- Figure 2: product reaction cards chosen in the debrief ----
cards = [  # (word, times chosen, negative?)
    ("Incomplete", 3, True), ("Difficult", 2, True), ("Confusing", 2, True),
    ("Frustrating", 2, True), ("Unclear", 2, True), ("Unorganized", 2, True),
    ("Informative", 2, False), ("Professional", 2, False), ("Engaging", 1, False),
    ("Familiar", 1, False), ("Simple", 1, False), ("Ordinary", 1, False),
]
cards = cards[::-1]
fig, ax = plt.subplots(figsize=(7.5, 4.6), dpi=200)
bars = ax.barh([c[0] for c in cards], [c[1] for c in cards],
               color=[ORANGE if c[2] else BLUE for c in cards],
               height=0.62, edgecolor="white", linewidth=1.5)
for bar, c in zip(bars, cards):
    ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height() / 2, str(c[1]),
            va="center", fontsize=9, color=INK)
ax.set_xlim(0, 3.5)
ax.set_xticks([0, 1, 2, 3])
ax.set_xlabel("Number of participants who chose the word (n = 6)")
style_axes(ax)
ax.legend(handles=[Patch(color=ORANGE, label="Negative (13 of 21 picks)"),
                   Patch(color=BLUE, label="Positive or neutral (8 of 21 picks)")],
          loc="lower center", bbox_to_anchor=(0.4, 1.0), ncol=2, frameon=False, fontsize=9)
fig.tight_layout()
fig.savefig(OUT / "fig-reaction-cards.png", bbox_inches="tight")
plt.close(fig)
print("wrote", *sorted(p.name for p in OUT.glob("*.png")))
