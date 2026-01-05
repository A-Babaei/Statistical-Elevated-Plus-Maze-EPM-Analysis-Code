# Elevated Plus Maze (EPM) Analysis Code

## **Project**

**Network-based deep brain stimulation modulates anxiety-like behavior in a Parkinsonian rat model**

**Author:** A. Babaei
**Program:** M.Sc. Bioelectrical Engineering
**Associated Manuscript:** Targeting a thalamic control node for Parkinson’s disease therapy

---

## **Overview**

This repository contains the complete Python code used to analyze **Elevated Plus Maze (EPM)** behavioral data collected as part of a non-motor behavioral assessment in a hemiparkinsonian rat model.

The analysis is designed to:

* Respect the **within-subject experimental design** (DBS ON vs OFF)
* Analyze **all recorded EPM parameters**
* Apply **multiple-comparison correction**
* Identify **primary anxiety-related outcomes**
* Provide **full transparency** via subject-level and supplementary tables
* Generate **publication-quality figures** (raincloud-style and paired plots)

The workflow follows best practices for **behavioral neuroscience and biomedical engineering**, emphasizing effect sizes and transparent reporting.

---

## **Input Data**

### Required file

```
EPM.xlsx
```

### Data structure

* **Rows:** EPM behavioral parameters (e.g., Time_OpenArms, Percent_OpenArms, MeanSpeed_Overall_cm/s)
* **Columns:** Subject × condition combinations, e.g.:

  * `PD_1_NoStim`, `PD_1_Stim`
  * `CO_1_NoStim`, `CO_1_Stim`

The first column must contain the parameter names.

---

## **Statistical Analysis**

### Design

* **Within-subject:** No-Stim vs Stim
* **Between-subject:** PD vs Control
* **Primary outcomes (pre-specified):**

  * Time_OpenArms
  * Percent_OpenArms

### Tests

* Paired *t*-test or Wilcoxon signed-rank test (based on normality)
* **Holm–Bonferroni correction** applied within each group
* **Effect size:** Cohen’s *dz*

### Additional analyses

* Locomotion control (MeanSpeed_Overall_cm/s, Entries_ClosedArms)
* Subject-level DBS response tables

---

## **Outputs**

### Supplementary Tables

* **S1:** Full EPM statistics (PD group)
* **S2:** Full EPM statistics (Control group)
* **S3:** Subject-level EPM responses (primary outcomes)
* **S4:** Locomotion control analysis

### Figures

* Raincloud-style plots for primary EPM outcomes (PD)
* Paired dot plots illustrating within-subject changes
* Locomotion control figures (Supplementary)

All figures are saved at **600 dpi**, suitable for journal submission.

---

## **How to Run**

1. Update the file paths in the script:

```python
DATA_PATH = r"G:\...\EPM.xlsx"
OUT_DIR   = r"G:\...\EPM_Results"
```

2. Run the script:

```bash
python EPM_Full_Analysis_FINAL.py
```

---

## **Interpretation Notes**

* Lack of significance after multiple-comparison correction does **not** imply absence of effect.
* Primary outcomes are reported based on **effect sizes and directional consistency**, with full transparency provided in Supplementary Tables.
* Locomotion controls ensure that anxiety-related effects are **not driven by hyperactivity**.

---

## **Citation**

If you use or adapt this code, please cite or acknowledge:

> Babaei, A. Bioelectrical Engineering. Network-based modulation of non-motor behavior via parafascicular deep brain stimulation.


