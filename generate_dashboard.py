import pandas as pd
import json
import numpy as np

# =====================
# LOAD EXCEL
# =====================

df = pd.read_excel(
    "Freight Status report 10062026.XLSX"
)

df.columns = df.columns.str.strip()

# Standardize columns
column_mapping = {
    "ATP - formula (no negative": "ATP",
    "Vendor Ctry.1": "Vendor Country Name"
}
df.rename(columns=column_mapping, inplace=True)

# Classify Import / Local dynamically if not present
if "Import / Local" not in df.columns:
    country_col = "Vendor Country Name" if "Vendor Country Name" in df.columns else ("Vendor Ctry.1" if "Vendor Ctry.1" in df.columns else ("Vendor Ctry" if "Vendor Ctry" in df.columns else None))
    if country_col:
        df["Import / Local"] = np.where(
            df[country_col]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.strip()
            .isin(["KUWAIT", "SAUDI ARABIA"]),
            "Local",
            "Import"
        )
    else:
        df["Import / Local"] = "Import"
        
df["Import_Type"] = df["Import / Local"]

# Ensure Category exists and combine D&W Production & doors and windows
if "Category" not in df.columns:
    df["Category"] = "N/A"
else:
    df["Category"] = df["Category"].fillna("N/A").astype(str).str.strip()
    category_masks = df["Category"].str.upper().isin(["D&W PRODUCTION", "D&W PRODUCTIONS", "DOORS & WINDOWS", "DOORS AND WINDOWS"])
    df.loc[category_masks, "Category"] = "DOORS & WINDOWS"

# Date conversion
date_columns = ["ATP", "GRP", "AWH", "BAYAN", "ETA", "ETD"]
for col in date_columns:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

# Remove blank PO rows
df = df[df["PO#"].notna()]

DELAY_COLS = ["EXF Delay", "ETD Delay", "ETA Delay", "BAYAN Delay", "AWH Delay", "GR Delay", "Port Delay"]

def compute_po_input_delay(dataframe):
    delay_cols_present = [c for c in DELAY_COLS if c in dataframe.columns]
    if not delay_cols_present or dataframe.empty:
        return {"total_delay_count": 0, "zero_count": 0, "marginal_count": 0, "total_pos": 0, "po_detail": []}

    # Aggregate per PO# — take the max of each delay col (a PO has one true delay value)
    # Clip lower=0 so negative delay days (early arrivals) do not offset positive delays or count as delays
    po_grp = dataframe.groupby("PO#")[delay_cols_present].max().clip(lower=0).reset_index()
    po_grp["total_delay"] = po_grp[delay_cols_present].fillna(0).sum(axis=1)

    total_delay_count = 0
    zero_count = 0
    po_detail = []

    for _, row in po_grp.iterrows():
        td = float(row["total_delay"])
        if td > 0:
            total_delay_count += 1
            status = "Delayed"
        else:
            zero_count += 1
            status = "On Time"
        po_detail.append({
            "PO#": str(row["PO#"]),
            "total_delay": round(td, 2),
            "status": status
        })

    return {
        "total_delay_count": total_delay_count,
        "zero_count": zero_count,
        "marginal_count": 0,
        "total_pos": len(po_grp),
        "po_detail": po_detail
    }

dashboard = {}

# =====================
# KPI CARDS
# =====================

dashboard["total_pos"] = int(
    df["PO#"].nunique()
)

dashboard["total_vendors"] = int(
    df["Vendor Name"].nunique()
)

dashboard["total_containers"] = int(
    df["Container"].nunique()
)

# =====================
# DELAY COUNTS
# PO BASED
# =====================

dashboard["delay_counts"] = {
    "EXF": int(df.loc[df["EXF Delay"] > 0, "PO#"].nunique()),
    "ETD": int(df.loc[df["ETD Delay"] > 0, "PO#"].nunique()),
    "ETA": int(df.loc[df["ETA Delay"] > 0, "PO#"].nunique()),
    "BAYAN": int(df.loc[df["BAYAN Delay"] > 0, "PO#"].nunique()),
    "AWH": int(df.loc[df["AWH Delay"] > 0, "PO#"].nunique()),
    "GR": int(df.loc[df["GR Delay"] > 0, "PO#"].nunique()),
    "PORT": int(df.loc[df["Port Delay"] > 0, "PO#"].nunique())
}

dashboard["total_delay_count"] = sum(
    dashboard["delay_counts"].values()
)

# =====================
# IMPORT KPIs
# =====================

import_df = df[
    df["Import / Local"] == "Import"
]

dashboard["import_pos"] = int(
    import_df["PO#"].nunique()
)

dashboard["import_vendors"] = int(
    import_df["Vendor Name"].nunique()
)

dashboard["import_countries"] = int(
    import_df["Vendor Ctry"].nunique()
)

# =====================
# IMPORT VS LOCAL
# =====================

dashboard["import_local"] = (
    df.groupby("Import / Local")["PO#"]
      .nunique()
      .reset_index()
      .to_dict("records")
)

# =====================
# TOP VENDORS
# =====================

dashboard["top_vendors"] = (
    df.groupby("Vendor Name")["PO#"]
      .nunique()
      .sort_values(ascending=False)
      .head(10)
      .reset_index()
      .to_dict("records")
)

# =====================
# COUNTRY ANALYSIS
# =====================

dashboard["countries"] = (
    df.groupby("Vendor Ctry")["PO#"]
      .nunique()
      .sort_values(ascending=False)
      .reset_index()
      .to_dict("records")
)

# =====================
# CATEGORY ANALYSIS
# =====================

dashboard["categories"] = (
    df.groupby("Category")["PO#"]
      .nunique()
      .sort_values(ascending=False)
      .reset_index()
      .to_dict("records")
)

# =====================
# VENDOR TABLE
# PO BASED
# =====================

vendor_records = []

for (
    vendor,
    country,
    imp_local
), group in df.groupby(
    [
        "Vendor Name",
        "Vendor Ctry",
        "Import / Local"
    ],
    dropna=False
):
    
    exf = int(group.loc[group["EXF Delay"] > 0, "PO#"].nunique())
    etd = int(group.loc[group["ETD Delay"] > 0, "PO#"].nunique())
    eta = int(group.loc[group["ETA Delay"] > 0, "PO#"].nunique())
    bayan = int(group.loc[group["BAYAN Delay"] > 0, "PO#"].nunique())
    awh = int(group.loc[group["AWH Delay"] > 0, "PO#"].nunique())
    gr = int(group.loc[group["GR Delay"] > 0, "PO#"].nunique())
    port = int(group.loc[group["Port Delay"] > 0, "PO#"].nunique())

    vendor_records.append({
        "Vendor Name": str(vendor),
        "Vendor Ctry": str(country),
        "Import / Local": str(imp_local),
        "EXF Delay": exf,
        "ETD Delay": etd,
        "ETA Delay": eta,
        "BAYAN Delay": bayan,
        "AWH Delay": awh,
        "GR Delay": gr,
        "Port Delay": port,
        "Total Delay": (exf + etd + eta + bayan + awh + gr + port)
    })

dashboard["vendors"] = vendor_records

# =====================
# CATEGORY TABLE
# PO BASED
# =====================

def get_category_records(dataframe):
    records = []
    for category, group in dataframe.groupby("Category", dropna=False):
        records.append({
            "Category": str(category),
            "EXF Delay": int(group.loc[group["EXF Delay"] > 0, "PO#"].nunique()),
            "ETD Delay": int(group.loc[group["ETD Delay"] > 0, "PO#"].nunique()),
            "ETA Delay": int(group.loc[group["ETA Delay"] > 0, "PO#"].nunique()),
            "BAYAN Delay": int(group.loc[group["BAYAN Delay"] > 0, "PO#"].nunique()),
            "AWH Delay": int(group.loc[group["AWH Delay"] > 0, "PO#"].nunique()),
            "GR Delay": int(group.loc[group["GR Delay"] > 0, "PO#"].nunique()),
            "Port Delay": int(group.loc[group["Port Delay"] > 0, "PO#"].nunique())
        })
    return records

dashboard["category_table"] = get_category_records(df)
dashboard["import_category_table"] = get_category_records(df[df["Import / Local"] == "Import"])
dashboard["local_category_table"] = get_category_records(df[df["Import / Local"] == "Local"])

# =====================
# FILTERS FOR SLICERS
# =====================

dashboard["filters"] = {

    "countries": sorted(
        df["Vendor Ctry"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    ),

    "categories": sorted(
        df["Category"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    ),

    "vendors": sorted(
        df["Vendor Name"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    ),

    "import_local": sorted(
        df["Import / Local"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    ),

    "import_countries": sorted(
        df.loc[
            df["Import / Local"] == "Import",
            "Vendor Ctry"
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    ),

    "local_countries": sorted(
        df.loc[
            df["Import / Local"] == "Local",
            "Vendor Ctry"
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
}

# =====================
# IMPORT & LOCAL ANALYTICS
# =====================

import_only = df[
    df["Import / Local"] == "Import"
]

local_only = df[
    df["Import / Local"] == "Local"
]

# Local KPIs
dashboard["local_pos"] = int(
    local_only["PO#"].nunique()
)

dashboard["local_vendors"] = int(
    local_only["Vendor Name"].nunique()
)

dashboard["local_countries"] = int(
    local_only["Vendor Ctry"].nunique()
)

# Container counts
dashboard["import_containers"] = int(
    import_only["Container"].nunique()
)

dashboard["local_containers"] = int(
    local_only["Container"].nunique()
)

# Country analysis
dashboard["import_country_analysis"] = (
    import_only
    .groupby("Vendor Ctry")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .reset_index()
    .to_dict("records")
)

dashboard["local_country_analysis"] = (
    local_only
    .groupby("Vendor Ctry")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .reset_index()
    .to_dict("records")
)

# Category analysis
dashboard["import_category_analysis"] = (
    import_only
    .groupby("Category")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .reset_index()
    .to_dict("records")
)

dashboard["local_category_analysis"] = (
    local_only
    .groupby("Category")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .reset_index()
    .to_dict("records")
)

# Top Vendors
dashboard["top_import_vendors"] = (
    import_only
    .groupby("Vendor Name")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .head(15)
    .reset_index()
    .to_dict("records")
)

dashboard["top_local_vendors"] = (
    local_only
    .groupby("Vendor Name")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .head(15)
    .reset_index()
    .to_dict("records")
)

# Delay stage breakdown
dashboard["import_delay_counts"] = {
    "EXF": int(import_only.loc[import_only["EXF Delay"] > 0, "PO#"].nunique()),
    "ETD": int(import_only.loc[import_only["ETD Delay"] > 0, "PO#"].nunique()),
    "ETA": int(import_only.loc[import_only["ETA Delay"] > 0, "PO#"].nunique()),
    "BAYAN": int(import_only.loc[import_only["BAYAN Delay"] > 0, "PO#"].nunique()),
    "AWH": int(import_only.loc[import_only["AWH Delay"] > 0, "PO#"].nunique()),
    "GR": int(import_only.loc[import_only["GR Delay"] > 0, "PO#"].nunique()),
    "PORT": int(import_only.loc[import_only["Port Delay"] > 0, "PO#"].nunique())
}

dashboard["local_delay_counts"] = {
    "EXF": int(local_only.loc[local_only["EXF Delay"] > 0, "PO#"].nunique()),
    "ETD": int(local_only.loc[local_only["ETD Delay"] > 0, "PO#"].nunique()),
    "ETA": int(local_only.loc[local_only["ETA Delay"] > 0, "PO#"].nunique()),
    "BAYAN": int(local_only.loc[local_only["BAYAN Delay"] > 0, "PO#"].nunique()),
    "AWH": int(local_only.loc[local_only["AWH Delay"] > 0, "PO#"].nunique()),
    "GR": int(local_only.loc[local_only["GR Delay"] > 0, "PO#"].nunique()),
    "PORT": int(local_only.loc[local_only["Port Delay"] > 0, "PO#"].nunique())
}

# Cumulative delays
dashboard["import_delay_total"] = sum(
    dashboard["import_delay_counts"].values()
)

dashboard["local_delay_total"] = sum(
    dashboard["local_delay_counts"].values()
)

# =====================
# EXTRA ADVANCED METRICS FOR IMPORT & LOCAL
# =====================

# Import Value and Delays
import_delayed = import_only[import_only["Over All Delay"] > 0]
dashboard["import_total_value"] = float(import_only["InbValKWD"].sum())
dashboard["import_delayed_value"] = float(import_delayed["InbValKWD"].sum())
dashboard["import_avg_delay_days"] = float(import_delayed["Over All Delay"].mean()) if len(import_delayed) > 0 else 0.0

# Import Country financial value and duration analysis
import_val_by_ctry = import_delayed.groupby("Vendor Ctry")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
dashboard["import_value_by_country"] = import_val_by_ctry.to_dict("records")

import_delay_by_ctry = import_delayed.groupby("Vendor Ctry")["Over All Delay"].mean().sort_values(ascending=False).reset_index()
dashboard["import_avg_delay_by_country"] = import_delay_by_ctry.to_dict("records")

# Import Value by Category breakdown
import_val_by_cat = import_delayed.groupby("Category")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
dashboard["import_value_by_category"] = import_val_by_cat.to_dict("records")


# Local Value and Delays
local_delayed = local_only[local_only["Over All Delay"] > 0]
dashboard["local_total_value"] = float(local_only["InbValKWD"].sum())
dashboard["local_delayed_value"] = float(local_delayed["InbValKWD"].sum())
dashboard["local_avg_delay_days"] = float(local_delayed["Over All Delay"].mean()) if len(local_delayed) > 0 else 0.0

# Local Country financial value and duration analysis
local_val_by_ctry = local_delayed.groupby("Vendor Ctry")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
dashboard["local_value_by_country"] = local_val_by_ctry.to_dict("records")

local_delay_by_ctry = local_delayed.groupby("Vendor Ctry")["Over All Delay"].mean().sort_values(ascending=False).reset_index()
dashboard["local_avg_delay_by_country"] = local_delay_by_ctry.to_dict("records")

# Local Value by Category breakdown
local_val_by_cat = local_delayed.groupby("Category")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
dashboard["local_value_by_category"] = local_val_by_cat.to_dict("records")

# Top 10 Delayed Vendors by count of delayed POs
dashboard["import_top_delayed_vendors"] = (
    import_delayed.groupby("Vendor Name")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .head(10)
    .reset_index()
    .rename(columns={"PO#": "Delayed POs"})
    .to_dict("records")
)

dashboard["local_top_delayed_vendors"] = (
    local_delayed.groupby("Vendor Name")["PO#"]
    .nunique()
    .sort_values(ascending=False)
    .head(10)
    .reset_index()
    .rename(columns={"PO#": "Delayed POs"})
    .to_dict("records")
)

# =====================
# INSIGHTS
# =====================

top_country = (
    df.groupby("Vendor Ctry")["PO#"]
      .nunique()
      .sort_values(ascending=False)
      .reset_index()
)

top_vendor = (
    df.groupby("Vendor Name")["PO#"]
      .nunique()
      .sort_values(ascending=False)
      .reset_index()
)

top_category = (
    df.groupby("Category")["PO#"]
      .nunique()
      .sort_values(ascending=False)
      .reset_index()
)

dashboard["insights"] = {

    "highest_country":
        top_country.iloc[0]["Vendor Ctry"],

    "highest_vendor":
        top_vendor.iloc[0]["Vendor Name"],

    "highest_category":
        top_category.iloc[0]["Category"],

    "highest_delay_type":
        max(
            dashboard["delay_counts"],
            key=dashboard["delay_counts"].get
        )
}
dashboard["top_delay_vendors"] = sorted(
    vendor_records,
    key=lambda x: x["Total Delay"],
    reverse=True
)[:15]

# =====================
# PO INPUT DELAY ANALYSIS
# =====================
dashboard["po_input_delay"] = {
    "all": compute_po_input_delay(df),
    "import": compute_po_input_delay(df[df["Import / Local"] == "Import"]),
    "local": compute_po_input_delay(df[df["Import / Local"] == "Local"])
}

# =====================
# ANOMALY DETECTION DATA — Distinct PO# based
# =====================
adf = df.copy()

rules_masks = []
if "ATP" in adf.columns and "GRP" in adf.columns:
    rules_masks.append(("ATP_BEFORE_GRP", (adf["ATP"].notna() & adf["GRP"].notna() & (adf["ATP"] < adf["GRP"]))))
if "ATP" in adf.columns and "AWH" in adf.columns:
    rules_masks.append(("ATP_BEFORE_AWH", (adf["ATP"].notna() & adf["AWH"].notna() & (adf["ATP"] < adf["AWH"]))))
if "ATP" in adf.columns and "BAYAN" in adf.columns:
    rules_masks.append(("ATP_BEFORE_BAYAN", (adf["ATP"].notna() & adf["BAYAN"].notna() & (adf["ATP"] < adf["BAYAN"]))))
if "ATP" in adf.columns and "ETA" in adf.columns:
    rules_masks.append(("ATP_BEFORE_ETA", (adf["ATP"].notna() & adf["ETA"].notna() & (adf["ATP"] < adf["ETA"]))))
if "ATP" in adf.columns and "ETD" in adf.columns:
    rules_masks.append(("ATP_BEFORE_ETD", (adf["ATP"].notna() & adf["ETD"].notna() & (adf["ATP"] < adf["ETD"]))))
if "ATP" in adf.columns and "GRP" in adf.columns:
    grp_gap = (adf["ATP"] - adf["GRP"]).dt.days
    rules_masks.append(("ATP_GRP_GAP_GT_5", (adf["ATP"].notna() & adf["GRP"].notna() & (grp_gap > 5))))
    
RULE_FRIENDLY_NAMES = {
    "ATP_BEFORE_GRP": "ATP Date is before GRP Date",
    "ATP_BEFORE_AWH": "ATP Date is before AWH Date",
    "ATP_BEFORE_BAYAN": "ATP Date is before BAYAN Date",
    "ATP_BEFORE_ETA": "ATP Date is before ETA Date",
    "ATP_BEFORE_ETD": "ATP Date is before ETD Date",
    "ATP_GRP_GAP_GT_5": "ATP to GRP Gap exceeds 5 Days"
}

po_violations = {po: set() for po in adf["PO#"].unique()}
total_pos_audited = int(adf["PO#"].nunique())

anomaly_summary = []
for rule_name, mask in rules_masks:
    violated_pos = adf.loc[mask, "PO#"].unique()
    count = len(violated_pos)
    friendly = RULE_FRIENDLY_NAMES.get(rule_name, rule_name)
    anomaly_summary.append({
        "Rule": friendly,
        "Violations": count,
        "Percent": round((count / total_pos_audited) * 100, 2) if total_pos_audited > 0 else 0
    })
    for po in violated_pos:
        po_violations[po].add(friendly)
        
# Group by PO# and compile unique PO records
po_records = []
for po, gp in adf.groupby("PO#"):
    violations = po_violations[po]
    issue_count = len(violations)
    
    def get_first_valid(series, default="N/A"):
        valid = series.dropna()
        if not valid.empty:
            val = valid.iloc[0]
            if pd.notna(val) and str(val).strip() != "":
                return val
        return default

    category = get_first_valid(gp["Category"], "N/A")
    vendor = get_first_valid(gp["Vendor Name"], "N/A")
    import_local = get_first_valid(gp["Import / Local"], "Import")
    
    severity = "No"
    if issue_count > 0:
        severity = "Severe" if import_local == "Import" else "Medium"
        
    atp_val = gp["ATP"].dropna().iloc[0] if not gp["ATP"].dropna().empty else pd.NaT
    grp_val = gp["GRP"].dropna().iloc[0] if not gp["GRP"].dropna().empty else pd.NaT
    awh_val = gp["AWH"].dropna().iloc[0] if not gp["AWH"].dropna().empty else pd.NaT
    bayan_val = gp["BAYAN"].dropna().iloc[0] if not gp["BAYAN"].dropna().empty else pd.NaT
    eta_val = gp["ETA"].dropna().iloc[0] if not gp["ETA"].dropna().empty else pd.NaT
    etd_val = gp["ETD"].dropna().iloc[0] if not gp["ETD"].dropna().empty else pd.NaT
    
    atp_str = atp_val.strftime("%Y-%m-%d") if pd.notna(atp_val) else "N/A"
    grp_str = grp_val.strftime("%Y-%m-%d") if pd.notna(grp_val) else "N/A"
    awh_str = awh_val.strftime("%Y-%m-%d") if pd.notna(awh_val) else "N/A"
    bayan_str = bayan_val.strftime("%Y-%m-%d") if pd.notna(bayan_val) else "N/A"
    eta_str = eta_val.strftime("%Y-%m-%d") if pd.notna(eta_val) else "N/A"
    etd_str = etd_val.strftime("%Y-%m-%d") if pd.notna(etd_val) else "N/A"
    
    # Count rule violations for this PO group across all its line items
    atp_before_grp_cnt = int((gp["ATP"].notna() & gp["GRP"].notna() & (gp["ATP"] < gp["GRP"])).sum()) if "ATP" in gp.columns and "GRP" in gp.columns else 0
    atp_before_awh_cnt = int((gp["ATP"].notna() & gp["AWH"].notna() & (gp["ATP"] < gp["AWH"])).sum()) if "ATP" in gp.columns and "AWH" in gp.columns else 0
    atp_before_bayan_cnt = int((gp["ATP"].notna() & gp["BAYAN"].notna() & (gp["ATP"] < gp["BAYAN"])).sum()) if "ATP" in gp.columns and "BAYAN" in gp.columns else 0
    atp_before_eta_cnt = int((gp["ATP"].notna() & gp["ETA"].notna() & (gp["ATP"] < gp["ETA"])).sum()) if "ATP" in gp.columns and "ETA" in gp.columns else 0
    atp_before_etd_cnt = int((gp["ATP"].notna() & gp["ETD"].notna() & (gp["ATP"] < gp["ETD"])).sum()) if "ATP" in gp.columns and "ETD" in gp.columns else 0
    
    atp_grp_gap_cnt = 0
    if "ATP" in gp.columns and "GRP" in gp.columns:
        grp_gap = (gp["ATP"] - gp["GRP"]).dt.days
        atp_grp_gap_cnt = int((gp["ATP"].notna() & gp["GRP"].notna() & (grp_gap > 5)).sum())

    po_records.append({
        "PO#": str(po),
        "Category": str(category),
        "Vendor": str(vendor),
        "Import_Local": str(import_local),
        "Issue_Count": issue_count,
        "Issue_Details": "; ".join(sorted(violations)),
        "Severity": severity,
        "ATP": atp_str,
        "GRP": grp_str,
        "AWH": awh_str,
        "BAYAN": bayan_str,
        "ETA": eta_str,
        "ETD": etd_str,
        "ATP_BEFORE_GRP": atp_before_grp_cnt,
        "ATP_BEFORE_AWH": atp_before_awh_cnt,
        "ATP_BEFORE_BAYAN": atp_before_bayan_cnt,
        "ATP_BEFORE_ETA": atp_before_eta_cnt,
        "ATP_BEFORE_ETD": atp_before_etd_cnt,
        "ATP_GRP_GAP_GT_5": atp_grp_gap_cnt
    })
    
violations_pos = [r for r in po_records if r["Issue_Count"] > 0]
total_violations = len(violations_pos)
severe_count = sum(1 for r in violations_pos if r["Severity"] == "Severe")
medium_count = sum(1 for r in violations_pos if r["Severity"] == "Medium")

# Detail list of violations sorted by Issue_Count descending
detail_list = sorted(violations_pos, key=lambda x: x["Issue_Count"], reverse=True)

dashboard["anomaly_data"] = {
    "total_rows_audited": total_pos_audited,
    "total_violations": total_violations,
    "severe_count": severe_count,
    "medium_count": medium_count,
    "summary": anomaly_summary,
    "details": detail_list
}


def generate_ai_insights(df, dashboard):
    import os
    import requests
    import json
    
    # Pre-calculate values for fallback insights
    total_pos = dashboard.get("total_pos", 0)
    total_vendors = dashboard.get("total_vendors", 0)
    delay_info = dashboard.get("po_input_delay", {}).get("all", {})
    delay_count = delay_info.get("total_delay_count", 0)
    total_pid_pos = delay_info.get("total_pos", 1) or 1
    delay_rate_calc = round((delay_count / total_pid_pos) * 100, 1)
    
    import_pos = dashboard.get("import_pos", 0)
    import_val = dashboard.get("import_total_value", 0.0)
    import_at_risk = dashboard.get("import_delayed_value", 0.0)
    import_avg_delay = dashboard.get("import_avg_delay_days", 0.0)
    import_risk_rate = round((import_at_risk / import_val * 100) if import_val > 0 else 0.0, 1)
    
    local_pos = dashboard.get("local_pos", 0)
    local_val = dashboard.get("local_total_value", 0.0)
    local_at_risk = dashboard.get("local_delayed_value", 0.0)
    local_avg_delay = dashboard.get("local_avg_delay_days", 0.0)
    local_risk_rate = round((local_at_risk / local_val * 100) if local_val > 0 else 0.0, 1)
    
    anomaly_data = dashboard.get("anomaly_data", {})
    audited_pos = anomaly_data.get("total_rows_audited", 0)
    violated_pos = anomaly_data.get("total_violations", 0)
    anomaly_rate = round((violated_pos / audited_pos * 100), 1) if audited_pos > 0 else 0.0
    
    top_rule = "None"
    top_rule_count = 0
    for r in anomaly_data.get("summary", []):
        if r["Violations"] > top_rule_count:
            top_rule_count = r["Violations"]
            top_rule = r["Rule"]

    # Extract distinct categories from data
    categories_list = ["All"]
    if "Category" in df.columns:
        categories_list += [str(x) for x in df["Category"].dropna().unique() if str(x).strip()]
            
    # Local fallback list in case Groq is unavailable
    fallback_insights = [
        {"type": "Global", "category": "All", "text": f"Global Sourcing: Audited a total of {total_pos:,} Purchase Orders spanning {total_vendors} unique vendors across all business lines."},
        {"type": "Global", "category": "All", "text": f"Tracking Health: {delay_count:,} active Purchase Orders are currently flagged with logistics delays, representing an overall tracking delay rate of {delay_rate_calc}%."},
        {"type": "Global", "category": "All", "text": f"Global Sourcing Nodes: Supply chain operations span {len(dashboard.get('filters', {}).get('countries', []))} countries, with {dashboard.get('insights', {}).get('highest_country', 'N/A')} as the primary source hub."},
        {"type": "Global", "category": "All", "text": "Global Vendor Concentration: The top supply vendors represent a significant portion of active PO volume, requiring strict SLA management."},
        
        {"type": "Import", "category": "All", "text": f"Import Volume: Import shipments total {import_pos:,} POs, managing a gross financial value of {import_val:,.0f} KWD."},
        {"type": "Import", "category": "All", "text": f"Import Financial Risk: Delay-exposed import value stands at {import_at_risk:,.0f} KWD, representing {import_risk_rate}% of the total import portfolio."},
        {"type": "Import", "category": "All", "text": f"Import Bottlenecks: Average import delay is {import_avg_delay:.1f} days, with major clearance delays occurring at BAYAN customs ({dashboard.get('import_delay_counts', {}).get('BAYAN', 0)} POs)."},
        {"type": "Import", "category": "All", "text": f"Import Delay Stage: AWH entry delays affect {dashboard.get('import_delay_counts', {}).get('AWH', 0)} POs, showing warehousing bottlenecks."},
        {"type": "Import", "category": "All", "text": f"Import Vendor Latency: Supplier {dashboard.get('import_top_delayed_vendors', [{}])[0].get('Vendor Name', 'N/A') if len(dashboard.get('import_top_delayed_vendors', [])) > 0 else 'N/A'} is the top delayed supplier with {dashboard.get('import_top_delayed_vendors', [{}])[0].get('Delayed POs', 0) if len(dashboard.get('import_top_delayed_vendors', [])) > 0 else 0} delayed POs."},
        {"type": "Import", "category": "KITCHEN PRODUCTION", "text": f"Import Category Focus: Kitchen Production imports represent a key value driver, experiencing average delays of {import_avg_delay:.1f} days."},
        
        {"type": "Local", "category": "All", "text": f"Local Delivery Comparison: Local operations cover {local_pos:,} POs worth {local_val:,.0f} KWD with a delay rate of {local_risk_rate}% and average delay of {local_avg_delay:.1f} days (significantly lower than import latency)."},
        {"type": "Local", "category": "All", "text": f"Local Warehousing: Local delay counts show low transit-related friction with average fulfillment cycle of {local_avg_delay:.1f} days."},
        
        {"type": "Anomaly", "category": "All", "text": f"Data Compliance Auditing: Out of {audited_pos:,} audited POs, {violated_pos:,} exhibit process compliance anomalies, yielding a {anomaly_rate}% data anomaly rate."},
        {"type": "Anomaly", "category": "All", "text": f"Compliance Violations: The most frequent validation mismatch is '{top_rule}' affecting {top_rule_count:,} distinct POs."},
        {"type": "Anomaly", "category": "All", "text": "Data Governance: Critical date mismatches (such as ATP after Warehousing or BAYAN clearance) drive the majority of compliance flags, suggesting data logging latency."}
    ]
    
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        try:
            with open(".env", "r") as f:
                for line in f:
                    if line.startswith("GROQ_API_KEY="):
                        api_key = line.split("=", 1)[1].strip()
                        break
        except Exception:
            pass
            
    if not api_key:
        return fallback_insights
        
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    context = {
        "global_metrics": {
            "total_pos": total_pos,
            "total_vendors": total_vendors,
            "total_delays": delay_count,
            "delay_rate_percent": delay_rate_calc,
            "countries": dashboard.get("filters", {}).get("countries", [])
        },
        "import_metrics": {
            "total_pos": import_pos,
            "total_value_kwd": import_val,
            "value_at_risk_kwd": import_at_risk,
            "value_at_risk_percent": import_risk_rate,
            "avg_delay_days": import_avg_delay,
            "delay_stages_counts": dashboard.get("import_delay_counts", {}),
            "top_delayed_vendors": dashboard.get("import_top_delayed_vendors", [])[:3]
        },
        "local_metrics": {
            "total_pos": local_pos,
            "total_value_kwd": local_val,
            "value_at_risk_kwd": local_at_risk,
            "value_at_risk_percent": local_risk_rate,
            "avg_delay_days": local_avg_delay,
            "delay_stages_counts": dashboard.get("local_delay_counts", {}),
            "top_delayed_vendors": dashboard.get("local_top_delayed_vendors", [])[:3]
        },
        "anomaly_metrics": {
            "audited_pos": audited_pos,
            "violating_pos": violated_pos,
            "anomaly_rate_percent": anomaly_rate,
            "rules_summary": anomaly_data.get("summary", [])
        },
        "categories": categories_list
    }
    
    prompt = f"""
    You are an expert logistics and supply chain business analyst.
    Analyze the following freight operations dataset metrics and generate EXACTLY 15 actionable, professional business insights.
    
    The insights must follow this specific category and type distribution:
    - 4 Global/Overview insights (overall volume, top supply countries, overall tracking delays)
    - 6 Import insights (specific import value, value at risk, bottleneck stages like BAYAN/AWH, top delayed vendors, and category-level import insights for categories like {categories_list})
    - 2 Local insights (comparison of local delays/values vs import, local bottlenecks)
    - 3 Anomaly validation insights (anomaly rate, most common data error or date-out-of-sequence rules violating compliance, category-specific anomalies)
    
    Ensure all insights are clear, cite specific numbers/percentages from the data, and offer business value. Do not use generic placeholders.
    
    Return the response as a JSON object containing a single key "insights" which is a list of exactly 15 objects.
    Each object MUST have exactly these keys:
    - "type": the insight type (one of: "Global", "Import", "Local", "Anomaly")
    - "category": the specific category this insight relates to (one of: "All", or a specific category name from the list: {categories_list})
    - "text": the descriptive text of the insight, citing actual numbers from the metrics.
    
    Do not output any markdown code blocks, explanation text, or front/back matter. Output raw JSON only.
    """
    
    try:
        payload = {
            "model": "llama-3.1-8b-instant",
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.2,
            "response_format": { "type": "json_object" }
        }
        
        response = requests.post(url, json=payload, headers=headers, timeout=12)
        if response.status_code == 200:
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            parsed = json.loads(content)
            insights_list = parsed.get("insights", [])
            if len(insights_list) == 15 and all(isinstance(x, dict) and "type" in x and "category" in x and "text" in x for x in insights_list):
                return insights_list
    except Exception as e:
        print(f"Error calling Groq API: {str(e)}")
        
    return fallback_insights


dashboard["ai_insights"] = generate_ai_insights(df, dashboard)

# =====================
# SAVE JSON & JS FILE
# =====================

with open(
    "data.json",
    "w",
    encoding="utf-8"
) as f:

    json.dump(
        dashboard,
        f,
        indent=4,
        ensure_ascii=False
    )

with open(
    "dashboard_data.js", 
    "w", 
    encoding="utf-8"
) as f:
    
    f.write("const dashboardData = ")
    
    json.dump(
        dashboard, 
        f, 
        ensure_ascii=False
    )
    
    f.write(";")

# =====================
# CONSOLE LOGS
# =====================

print("dashboard data generated successfully")
print("Total POs:", dashboard["total_pos"])
print("Import POs:", dashboard["import_pos"])
print("EXF:", dashboard["delay_counts"]["EXF"])
print("ETD:", dashboard["delay_counts"]["ETD"])
print("ETA:", dashboard["delay_counts"]["ETA"])
print("PORT:", dashboard["delay_counts"]["PORT"])