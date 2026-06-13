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
# ANOMALY DETECTION DATA — Row based
# =====================
adf = df.copy()
adf["Issue_Count"] = 0
adf["Issue_Details"] = ""

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

anomaly_summary = []
for rule_name, mask in rules_masks:
    count = int(mask.sum())
    friendly = RULE_FRIENDLY_NAMES.get(rule_name, rule_name)
    anomaly_summary.append({
        "Rule": friendly,
        "Violations": count,
        "Percent": round((count / len(adf)) * 100, 2) if len(adf) > 0 else 0
    })
    adf.loc[mask, "Issue_Count"] += 1
    adf.loc[mask, "Issue_Details"] = adf.loc[mask, "Issue_Details"].apply(lambda x: x + friendly + "; ")
    
# Classify Severity
adf["Severity"] = np.select(
    [
        (adf["Issue_Count"] > 0) & (adf["Import / Local"] == "Import"),
        (adf["Issue_Count"] > 0) & (adf["Import / Local"] == "Local")
    ],
    ["Severe", "Medium"],
    default="No"
)

total_rows_audited = len(adf)
total_violations = int((adf["Issue_Count"] > 0).sum())
severe_count = int((adf["Severity"] == "Severe").sum())
medium_count = int((adf["Severity"] == "Medium").sum())

# Detail list of violations
anomalies_only = adf[adf["Issue_Count"] > 0].copy()
anomalies_only.sort_values(by="Issue_Count", ascending=False, inplace=True)

detail_list = []
for idx, row in anomalies_only.iterrows():
    atp_str = row["ATP"].strftime("%Y-%m-%d") if pd.notna(row["ATP"]) else "N/A"
    grp_str = row["GRP"].strftime("%Y-%m-%d") if pd.notna(row["GRP"]) else "N/A"
    awh_str = row["AWH"].strftime("%Y-%m-%d") if pd.notna(row["AWH"]) else "N/A"
    bayan_str = row["BAYAN"].strftime("%Y-%m-%d") if pd.notna(row["BAYAN"]) else "N/A"
    eta_str = row["ETA"].strftime("%Y-%m-%d") if pd.notna(row["ETA"]) else "N/A"
    etd_str = row["ETD"].strftime("%Y-%m-%d") if pd.notna(row["ETD"]) else "N/A"
    
    detail_list.append({
        "PO#": str(row["PO#"]),
        "Category": str(row["Category"]) if pd.notna(row["Category"]) else "N/A",
        "Vendor": str(row["Vendor Name"]) if pd.notna(row["Vendor Name"]) else "N/A",
        "Import_Local": str(row["Import / Local"]),
        "Issue_Count": int(row["Issue_Count"]),
        "Issue_Details": str(row["Issue_Details"]).strip("; "),
        "Severity": str(row["Severity"]),
        "ATP": atp_str,
        "GRP": grp_str,
        "AWH": awh_str,
        "BAYAN": bayan_str,
        "ETA": eta_str,
        "ETD": etd_str
    })
    
dashboard["anomaly_data"] = {
    "total_rows_audited": total_rows_audited,
    "total_violations": total_violations,
    "severe_count": severe_count,
    "medium_count": medium_count,
    "summary": anomaly_summary,
    "details": detail_list
}

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