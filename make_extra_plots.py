"""
Extra visualizations for the AI vs Student comparison.
Produces:
  plots/heatmap_disagreements.png        — question × paper heatmap for top-disagreement papers
  plots/overall_confusion_matrix.png     — TP/FP/FN/TN aggregated across all questions
  plots/agreement_by_n_disagree.png      — histogram of papers by number of disagreements
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import warnings, os
warnings.filterwarnings("ignore")

os.makedirs("plots", exist_ok=True)

print("Loading data...")
comp    = pd.read_excel("comparison_ai_vs_student.xlsx", sheet_name="Comparison")
summary = pd.read_excel("summary_stats.xlsx")

agree_cols = [c for c in comp.columns if "__agreement" in c]
comp["n_disagree"] = (comp[agree_cols] == "disagree").sum(axis=1)

SHORT_LABELS = {
    "Birthweight__agreement":       "Birthweight",
    "GA at delivery__agreement":    "GA at delivery",
    "GA at collection__agreement":  "GA at collection",
    "Sex of offspring__agreement":  "Sex of offspring",
    "Parity__agreement":            "Parity",
    "Gravidity__agreement":         "Gravidity",
    "N offspring/preg__agreement":  "N offspring/preg",
    "Race/ethnicity__agreement":    "Race/ethnicity",
    "Genetic ancestry__agreement":  "Genetic ancestry",
    "Maternal height__agreement":   "Maternal height",
    "Maternal pre-preg wt__agreement": "Maternal pre-preg wt",
    "Paternal height__agreement":   "Paternal height",
    "Paternal weight__agreement":   "Paternal weight",
    "Maternal age__agreement":      "Maternal age",
    "Paternal age__agreement":      "Paternal age",
    "Complication samples__agreement": "Complication samples",
    "Mode of delivery__agreement":  "Mode of delivery",
    "Fetal complications__agreement": "Fetal complications",
}

print("Generating heatmap...")
top40 = comp.nlargest(40, "n_disagree").copy()
top40_labels = (top40["GEO_ID"].astype(str) + " (" + top40["n_disagree"].astype(str) + ")")

heat_data = []
for col, short in SHORT_LABELS.items():
    if col in top40.columns:
        vals = top40[col].map({"agree": 0, "disagree": 1, "unclear": 0.5}).fillna(0.5)
        heat_data.append(vals.values)

heat_matrix = np.array(heat_data)   # shape: (n_questions, n_papers)
row_labels  = [SHORT_LABELS[c] for c in SHORT_LABELS if c in top40.columns]

fig, ax = plt.subplots(figsize=(18, 7))
cmap = mcolors.ListedColormap(["#C6EFCE", "#FFEB9C", "#FFC7CE"])
bounds = [-0.1, 0.3, 0.7, 1.1]
norm = mcolors.BoundaryNorm(bounds, cmap.N)

im = ax.imshow(heat_matrix, cmap=cmap, norm=norm, aspect="auto")

ax.set_xticks(range(len(top40_labels)))
ax.set_xticklabels(top40_labels, rotation=75, ha="right", fontsize=7.5)
ax.set_yticks(range(len(row_labels)))
ax.set_yticklabels(row_labels, fontsize=9)

from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor="#C6EFCE", label="Agree"),
    Patch(facecolor="#FFEB9C", label="Unclear"),
    Patch(facecolor="#FFC7CE", label="Disagree"),
]
ax.legend(handles=legend_elements, loc="upper right", fontsize=9)
ax.set_title("AI vs Student Disagreement Heatmap\n(Top 40 papers by disagreement count; columns sorted by #disagreements)",
             fontsize=11)
plt.tight_layout()
plt.savefig("plots/heatmap_disagreements.png", dpi=150)
plt.close()
print("  Saved plots/heatmap_disagreements.png")

print("Generating overall confusion matrix...")
tp = summary["TP (both yes)"].sum()
fp = summary["FP (AI yes, Stu no)"].sum()
fn = summary["FN (AI no,  Stu yes)"].sum()
tn = summary["TN (both no)"].sum()
total = tp + fp + fn + tn

matrix = np.array([[tp, fp], [fn, tn]])
labels_text = [
    [f"Both YES\n(TP)\nn={tp}\n{100*tp/total:.1f}%", f"AI YES\nStudent NO\n(FP)\nn={fp}\n{100*fp/total:.1f}%"],
    [f"AI NO\nStudent YES\n(FN)\nn={fn}\n{100*fn/total:.1f}%", f"Both NO\n(TN)\nn={tn}\n{100*tn/total:.1f}%"],
]

fig, ax = plt.subplots(figsize=(6, 5))
colors = np.array([[0.78, 0.95, 0.80], [1.0, 0.78, 0.79],
                   [1.0, 0.90, 0.70], [0.78, 0.95, 0.80]])
for i in range(2):
    for j in range(2):
        c = colors[i*2 + j]
        ax.add_patch(plt.Rectangle((j, 1-i), 1, 1, color=c))
        ax.text(j + 0.5, 1.5 - i, labels_text[i][j],
                ha="center", va="center", fontsize=10, fontweight="bold" if i==j else "normal")

ax.set_xlim(0, 2); ax.set_ylim(0, 2)
ax.set_xticks([0.5, 1.5]); ax.set_xticklabels(["AI says YES", "AI says NO"], fontsize=11)
ax.set_yticks([0.5, 1.5]); ax.set_yticklabels(["Student says NO", "Student says YES"], fontsize=11)
ax.set_title(f"Aggregated Confusion Matrix (all questions, n={total:,} decisions)", fontsize=11)
ax.tick_params(length=0)
for spine in ax.spines.values():
    spine.set_visible(False)
plt.tight_layout()
plt.savefig("plots/overall_confusion_matrix.png", dpi=150)
plt.close()
print("  Saved plots/overall_confusion_matrix.png")

print("Generating disagreement histogram...")
fig, ax = plt.subplots(figsize=(10, 5))
bins = range(0, comp["n_disagree"].max() + 2)
n, edges, patches = ax.hist(comp["n_disagree"], bins=bins,
                             color="steelblue", edgecolor="white", rwidth=0.85)
ax.set_xlabel("Number of questions with AI-Student disagreement", fontsize=11)
ax.set_ylabel("Number of papers", fontsize=11)
ax.set_title(f"Distribution of Disagreement Count per Paper (n={len(comp)} matched papers)",
             fontsize=11)

# annotate bars
for rect, count in zip(patches, n):
    if count > 0:
        ax.text(rect.get_x() + rect.get_width()/2, rect.get_height() + 0.5,
                str(int(count)), ha="center", va="bottom", fontsize=8)

# cumulative line on secondary axis
ax2 = ax.twinx()
cumulative = np.cumsum(n) / len(comp) * 100
midpoints = [(edges[i]+edges[i+1])/2 for i in range(len(edges)-1)]
ax2.plot(midpoints, cumulative, color="crimson", lw=1.5, linestyle="--", label="Cumulative %")
ax2.set_ylabel("Cumulative % of papers", color="crimson", fontsize=10)
ax2.tick_params(axis="y", labelcolor="crimson")
ax2.set_ylim(0, 105)
ax2.legend(loc="center right")

plt.tight_layout()
plt.savefig("plots/disagreement_histogram.png", dpi=150)
plt.close()
print("  Saved plots/disagreement_histogram.png")

print("Generating sorted agreement rates...")
s = summary.sort_values("Agreement (%)", ascending=True)
fig, ax = plt.subplots(figsize=(10, 7))
colors = ["#C6EFCE" if v >= 90 else "#FFEB9C" if v >= 80 else "#FFC7CE"
          for v in s["Agreement (%)"]]
bars = ax.barh(s["Question (short)"], s["Agreement (%)"], color=colors, edgecolor="white")
ax.axvline(90, color="green",  linestyle="--", lw=1, label="90% threshold")
ax.axvline(80, color="orange", linestyle="--", lw=1, label="80% threshold")
ax.set_xlabel("Agreement (%)", fontsize=11)
ax.set_title("AI vs Student Agreement Rate (sorted)\ngreen ≥ 90%, yellow 80-90%, red < 80%", fontsize=11)
ax.set_xlim(50, 105)
for bar, val in zip(bars, s["Agreement (%)"]):
    ax.text(val + 0.3, bar.get_y() + bar.get_height()/2,
            f"{val:.1f}%", va="center", fontsize=9)
ax.legend()
plt.tight_layout()
plt.savefig("plots/agreement_sorted.png", dpi=150)
plt.close()
print("  Saved plots/agreement_sorted.png")

print("\nAll extra plots done.")
