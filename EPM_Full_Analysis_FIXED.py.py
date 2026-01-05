# ============================================================
# Elevated Plus Maze – FULL FINAL ANALYSIS PIPELINE
# Author: A. Babaei
# ============================================================

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from statsmodels.stats.multitest import multipletests

# ============================================================
# PATHS (RAW STRINGS – WINDOWS SAFE)
# ============================================================

DATA_PATH = r"G:\Master\Experiment\Statistics\EPM\EPM.xlsx"
OUT_DIR   = r"G:\Master\Experiment\Statistics\EPM\EPM_Results"
os.makedirs(OUT_DIR, exist_ok=True)

# ============================================================
# STYLE & COLORS
# ============================================================

sns.set(style="whitegrid", context="talk")
plt.rcParams["pdf.fonttype"] = 42
plt.rcParams["ps.fonttype"] = 42

PD_COLOR    = "#D55E00"   # vermillion
POINT_COLOR = "#55A868"   # green
COND_ORDER  = ["No-Stim", "Stim"]

# ============================================================
# LOAD DATA
# ============================================================

df = pd.read_excel(DATA_PATH)

PARAM_COL = df.columns[0]
DATA_COLS = df.columns[1:]

# ============================================================
# ROBUST COLUMN PARSER
# Handles: PD_1_NoStim, PD1_Stim, CO_2_NoStim, CO2_Stim
# ============================================================

def parse_column(col):
    c = col.replace(" ", "").replace("-", "")
    if c.startswith("PD"):
        group = "PD"
    elif c.startswith("CO"):
        group = "Control"
    else:
        return None

    if "NoStim" in c:
        condition = "No-Stim"
    elif "Stim" in c:
        condition = "Stim"
    else:
        return None

    subject = c.split("_")[0]
    return subject, group, condition

# ============================================================
# BUILD LONG DATAFRAME
# ============================================================

records = []
for _, row in df.iterrows():
    param = row[PARAM_COL]
    for col in DATA_COLS:
        parsed = parse_column(col)
        if parsed is None:
            continue
        subject, group, condition = parsed
        records.append({
            "Parameter": param,
            "Subject": subject,
            "Group": group,
            "Condition": condition,
            "Value": row[col]
        })

long_df = pd.DataFrame(records)

if long_df.empty:
    raise RuntimeError("No data parsed. Check column names in EPM.xlsx")

# ============================================================
# PAIRED STATISTICS FUNCTION
# ============================================================

def paired_stats(data):
    off = data[data["Condition"] == "No-Stim"]["Value"].values
    on  = data[data["Condition"] == "Stim"]["Value"].values

    if len(off) < 3 or len(on) < 3:
        return None

    diff = on - off
    p_norm = stats.shapiro(diff).pvalue

    if p_norm > 0.05:
        stat, p = stats.ttest_rel(on, off)
        test = "Paired t-test"
    else:
        stat, p = stats.wilcoxon(on, off)
        test = "Wilcoxon signed-rank"

    dz = diff.mean() / diff.std(ddof=1)
    return test, stat, p, dz

# ============================================================
# RUN STATISTICS FOR ALL PARAMETERS
# ============================================================

results = []

for param in long_df["Parameter"].unique():
    for group in ["PD", "Control"]:
        subset = long_df[
            (long_df["Parameter"] == param) &
            (long_df["Group"] == group)
        ]
        res = paired_stats(subset)
        if res is None:
            continue
        test, stat, p, dz = res
        results.append({
            "Parameter": param,
            "Group": group,
            "Test": test,
            "Statistic": stat,
            "p_raw": p,
            "Cohens_dz": dz
        })

results_df = pd.DataFrame(
    results,
    columns=["Parameter","Group","Test","Statistic","p_raw","Cohens_dz"]
)

# ============================================================
# HOLM CORRECTION (WITHIN GROUP)
# ============================================================

results_df["p_holm"] = np.nan

for grp in results_df["Group"].unique():
    mask = results_df["Group"] == grp
    pvals = results_df.loc[mask,"p_raw"].values
    _, p_corr, _, _ = multipletests(pvals, method="holm")
    results_df.loc[mask,"p_holm"] = p_corr

# ============================================================
# SAVE SUPPLEMENTARY TABLES S1 & S2
# ============================================================

results_df[results_df["Group"]=="PD"].to_csv(
    os.path.join(OUT_DIR,"Supplementary_Table_S1_EPM_PD.csv"),
    index=False
)

results_df[results_df["Group"]=="Control"].to_csv(
    os.path.join(OUT_DIR,"Supplementary_Table_S2_EPM_Control.csv"),
    index=False
)

# ============================================================
# PRIMARY EPM PARAMETERS (PRE-SPECIFIED)
# ============================================================

PRIMARY_EPM_PARAMS = ["Time_OpenArms", "Percent_OpenArms"]

# ============================================================
# SUBJECT-LEVEL TABLE (S3)
# ============================================================

subject_tables = []

for param in PRIMARY_EPM_PARAMS:
    temp = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]
    if temp.empty:
        continue

    wide = temp.pivot_table(
        index="Subject",
        columns="Condition",
        values="Value"
    ).reset_index()

    wide["Delta_Stim_minus_NoStim"] = wide["Stim"] - wide["No-Stim"]
    wide["Parameter"] = param
    subject_tables.append(wide)

if subject_tables:
    subject_level_df = pd.concat(subject_tables, ignore_index=True)
    subject_level_df.to_csv(
        os.path.join(OUT_DIR,"Supplementary_Table_S3_EPM_SubjectLevel.csv"),
        index=False
    )

# ============================================================
# LOCOMOTION CONTROL (S4)
# ============================================================

LOCOMOTION_PARAMS = ["MeanSpeed_Overall_cm/s", "Entries_ClosedArms"]

locomotion_results = []

for param in LOCOMOTION_PARAMS:
    subset = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]
    if subset.empty:
        continue

    off = subset[subset["Condition"]=="No-Stim"]["Value"].values
    on  = subset[subset["Condition"]=="Stim"]["Value"].values

    stat, p = stats.ttest_rel(on, off)
    dz = (on - off).mean() / (on - off).std(ddof=1)

    locomotion_results.append([param, stat, p, dz])

locomotion_df = pd.DataFrame(
    locomotion_results,
    columns=["Parameter","Statistic","p_value","Cohens_dz"]
)

locomotion_df.to_csv(
    os.path.join(OUT_DIR,"Supplementary_Table_S4_EPM_LocomotionControl.csv"),
    index=False
)

# ============================================================
# RAINCLOUD PLOTS (PRIMARY PARAMETERS)
# ============================================================

for param in PRIMARY_EPM_PARAMS:
    data_param = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]
    if data_param.empty:
        continue

    plt.figure(figsize=(6,5))

    sns.violinplot(
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        inner=None,
        cut=0,
        linewidth=0,
        color=PD_COLOR
    )

    sns.boxplot(
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        width=0.25,
        showcaps=True,
        boxprops={"facecolor":"none","edgecolor":"black","linewidth":1.4},
        whiskerprops={"linewidth":1.4},
        medianprops={"color":"black","linewidth":1.6},
        showfliers=False
    )

    sns.stripplot(
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        jitter=0.12,
        size=8,
        color=POINT_COLOR,
        edgecolor="black",
        linewidth=0.6
    )

    plt.xlabel("")
    plt.ylabel(param)
    plt.title(f"EPM – {param} (PD Rats)")
    plt.tight_layout()
    plt.savefig(
        os.path.join(OUT_DIR,f"Fig_EPM_PD_{param}_Raincloud.png"),
        dpi=600
    )
    plt.show()

print("\nEPM FULL ANALYSIS COMPLETED SUCCESSFULLY.")

# ============================================================
# WITHIN-SUBJECT STATISTICAL TABLE (PD only)
# ============================================================

within_subject_tables = []

for param in PRIMARY_EPM_PARAMS:
    temp = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]

    if temp.empty:
        continue

    wide = temp.pivot_table(
        index="Subject",
        columns="Condition",
        values="Value"
    ).reset_index()

    wide["Delta_Stim_minus_NoStim"] = wide["Stim"] - wide["No-Stim"]
    wide["Percent_Change"] = (
        wide["Delta_Stim_minus_NoStim"] / wide["No-Stim"]
    ) * 100

    wide["Parameter"] = param
    within_subject_tables.append(wide)

within_subject_df = pd.concat(within_subject_tables, ignore_index=True)

within_subject_df.to_csv(
    os.path.join(OUT_DIR, "Supplementary_Table_S5_EPM_WithinSubject.csv"),
    index=False
)

print("\nWithin-subject EPM table created:")
print(within_subject_df.head())

# ============================================================
# FINAL MULTI-PANEL EPM FIGURE (A–F)
# ============================================================

fig, axes = plt.subplots(2, 3, figsize=(15, 10))
axes = axes.flatten()

panel_titles = [
    "A  Time in Open Arms (PD)",
    "B  Percent Open Arms (PD)",
    "C  Paired: Time Open Arms",
    "D  Paired: Percent Open Arms",
    "E  Mean Speed (Locomotion)",
    "F  Closed Arm Entries (Locomotion)"
]

# ---------- Panel A & B: Rainclouds ----------
for i, param in enumerate(PRIMARY_EPM_PARAMS):
    data_param = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]

    sns.violinplot(
        ax=axes[i],
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        inner=None,
        cut=0,
        linewidth=0,
        color=PD_COLOR
    )

    sns.boxplot(
        ax=axes[i],
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        width=0.25,
        showcaps=True,
        boxprops={"facecolor":"none","edgecolor":"black","linewidth":1.2},
        whiskerprops={"linewidth":1.2},
        medianprops={"color":"black","linewidth":1.5},
        showfliers=False
    )

    sns.stripplot(
        ax=axes[i],
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        jitter=0.12,
        size=6,
        color=POINT_COLOR,
        edgecolor="black",
        linewidth=0.6
    )

    axes[i].set_title(panel_titles[i])
    axes[i].set_xlabel("")
    axes[i].set_ylabel(param)

# ---------- Panel C & D: Paired dot plots ----------
for i, param in enumerate(PRIMARY_EPM_PARAMS):
    data_param = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]

    sns.pointplot(
        ax=axes[i+2],
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        join=True,
        markers="o",
        color=PD_COLOR,
        capsize=0.1
    )

    axes[i+2].set_title(panel_titles[i+2])
    axes[i+2].set_xlabel("")
    axes[i+2].set_ylabel(param)

# ---------- Panel E & F: Locomotion controls ----------
LOCOMOTION_PARAMS = [
    "MeanSpeed_Overall_cm/s",
    "Entries_ClosedArms"
]

for i, param in enumerate(LOCOMOTION_PARAMS):
    data_param = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]

    sns.pointplot(
        ax=axes[i+4],
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        join=True,
        markers="o",
        color="gray",
        capsize=0.1
    )

    axes[i+4].set_title(panel_titles[i+4])
    axes[i+4].set_xlabel("")
    axes[i+4].set_ylabel(param)

plt.tight_layout()
plt.savefig(
    os.path.join(OUT_DIR, "Fig_EPM_MultiPanel_Final.png"),
    dpi=600
)
plt.show()


SECONDARY_PARAMS = ["Time_Center", "Entries_OpenArms"]

for param in SECONDARY_PARAMS:
    data_param = long_df[
        (long_df["Parameter"] == param) &
        (long_df["Group"] == "PD")
    ]

    if data_param.empty:
        continue

    plt.figure(figsize=(5,4))

    sns.pointplot(
        data=data_param,
        x="Condition",
        y="Value",
        order=COND_ORDER,
        join=True,
        markers="o",
        color="gray"
    )

    plt.title(f"EPM – {param} (PD)")
    plt.xlabel("")
    plt.ylabel(param)
    plt.tight_layout()
    plt.savefig(
        os.path.join(OUT_DIR, f"Fig_EPM_PD_{param}_Secondary.png"),
        dpi=600
    )
    plt.show()
