import pandas as pd
import json

# =====================
# LOAD EXCEL
# =====================

df = pd.read_excel(
    "Freight Status report.XLSX",
    sheet_name="Sheet1"
)

df.columns = df.columns.str.strip()

# Remove blank PO rows
df = df[df["PO#"].notna()]

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