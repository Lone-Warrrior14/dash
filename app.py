from flask import Flask, request, render_template_string, send_file, jsonify
import pandas as pd
import json
import io
import zipfile
import os
import shutil
import uuid
import numpy as np
from pathlib import Path

app = Flask(__name__)

# Config
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit
TEMP_DIR = Path("temp_processing")


DELAY_COLS = ["EXF Delay", "ETD Delay", "ETA Delay", "BAYAN Delay", "AWH Delay", "GR Delay", "Port Delay"]


def compute_po_input_delay(dataframe):
    """
    For each DISTINCT PO#, sum all delay column values (max per PO row group).
    Returns:
        total_delay_count  : # of POs where sum of delays > 0
        zero_count         : # of POs where sum of delays == 0
        marginal_count     : 0 (deprecated)
        total_pos          : total distinct POs considered
        po_detail          : list of dicts per PO with PO#, total_delay, status
    """
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


def process_excel(file_stream):
    """Processes the uploaded excel stream and returns the dashboard JSON dict"""
    df = pd.read_excel(file_stream, sheet_name="Sheet1")
    df.columns = df.columns.str.strip()
    df = df[df["PO#"].notna()]
    
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
            
    dashboard = {}
    
    # KPI CARDS
    dashboard["total_pos"] = int(df["PO#"].nunique())
    dashboard["total_vendors"] = int(df["Vendor Name"].nunique())
    dashboard["total_containers"] = int(df["Container"].nunique())
    
    # DELAY COUNTS PO BASED
    dashboard["delay_counts"] = {
        "EXF": int(df.loc[df["EXF Delay"] > 0, "PO#"].nunique()),
        "ETD": int(df.loc[df["ETD Delay"] > 0, "PO#"].nunique()),
        "ETA": int(df.loc[df["ETA Delay"] > 0, "PO#"].nunique()),
        "BAYAN": int(df.loc[df["BAYAN Delay"] > 0, "PO#"].nunique()),
        "AWH": int(df.loc[df["AWH Delay"] > 0, "PO#"].nunique()),
        "GR": int(df.loc[df["GR Delay"] > 0, "PO#"].nunique()),
        "PORT": int(df.loc[df["Port Delay"] > 0, "PO#"].nunique())
    }
    dashboard["total_delay_count"] = sum(dashboard["delay_counts"].values())
    
    # IMPORT KPIs
    import_df = df[df["Import / Local"] == "Import"]
    dashboard["import_pos"] = int(import_df["PO#"].nunique())
    dashboard["import_vendors"] = int(import_df["Vendor Name"].nunique())
    dashboard["import_countries"] = int(import_df["Vendor Ctry"].nunique()) if "Vendor Ctry" in import_df.columns else 0
    
    # IMPORT VS LOCAL
    dashboard["import_local"] = (
        df.groupby("Import / Local")["PO#"]
        .nunique()
        .reset_index()
        .to_dict("records")
    )
    
    # TOP VENDORS
    dashboard["top_vendors"] = (
        df.groupby("Vendor Name")["PO#"]
        .nunique()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
        .to_dict("records")
    )
    
    # COUNTRY ANALYSIS
    dashboard["countries"] = (
        df.groupby("Vendor Ctry")["PO#"]
        .nunique()
        .sort_values(ascending=False)
        .reset_index()
        .to_dict("records")
    )
    
    # CATEGORY ANALYSIS
    dashboard["categories"] = (
        df.groupby("Category")["PO#"]
        .nunique()
        .sort_values(ascending=False)
        .reset_index()
        .to_dict("records")
    )
    
    # VENDOR TABLE PO BASED
    vendor_records = []
    for (vendor, country, imp_local), group in df.groupby(["Vendor Name", "Vendor Ctry", "Import / Local"], dropna=False):
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
    
    # CATEGORY TABLE
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
    
    # FILTERS FOR SLICERS
    dashboard["filters"] = {
        "countries": sorted(df["Vendor Ctry"].dropna().astype(str).unique().tolist()),
        "categories": sorted(df["Category"].dropna().astype(str).unique().tolist()),
        "vendors": sorted(df["Vendor Name"].dropna().astype(str).unique().tolist()),
        "import_local": sorted(df["Import / Local"].dropna().astype(str).unique().tolist()),
        "import_countries": sorted(df.loc[df["Import / Local"] == "Import", "Vendor Ctry"].dropna().astype(str).unique().tolist()),
        "local_countries": sorted(df.loc[df["Import / Local"] == "Local", "Vendor Ctry"].dropna().astype(str).unique().tolist())
    }
    
    # IMPORT & LOCAL ANALYTICS
    import_only = df[df["Import / Local"] == "Import"]
    local_only = df[df["Import / Local"] == "Local"]
    
    dashboard["local_pos"] = int(local_only["PO#"].nunique())
    dashboard["local_vendors"] = int(local_only["Vendor Name"].nunique())
    dashboard["local_countries"] = int(local_only["Vendor Ctry"].nunique())
    
    dashboard["import_containers"] = int(import_only["Container"].nunique())
    dashboard["local_containers"] = int(local_only["Container"].nunique())
    
    dashboard["import_country_analysis"] = (
        import_only.groupby("Vendor Ctry")["PO#"].nunique()
        .sort_values(ascending=False).reset_index().to_dict("records")
    )
    dashboard["local_country_analysis"] = (
        local_only.groupby("Vendor Ctry")["PO#"].nunique()
        .sort_values(ascending=False).reset_index().to_dict("records")
    )
    
    dashboard["import_category_analysis"] = (
        import_only.groupby("Category")["PO#"].nunique()
        .sort_values(ascending=False).reset_index().to_dict("records")
    )
    dashboard["local_category_analysis"] = (
        local_only.groupby("Category")["PO#"].nunique()
        .sort_values(ascending=False).reset_index().to_dict("records")
    )
    
    dashboard["top_import_vendors"] = (
        import_only.groupby("Vendor Name")["PO#"].nunique()
        .sort_values(ascending=False).head(15).reset_index().to_dict("records")
    )
    dashboard["top_local_vendors"] = (
        local_only.groupby("Vendor Name")["PO#"].nunique()
        .sort_values(ascending=False).head(15).reset_index().to_dict("records")
    )
    
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
    
    dashboard["import_delay_total"] = sum(dashboard["import_delay_counts"].values())
    dashboard["local_delay_total"] = sum(dashboard["local_delay_counts"].values())
    
    # EXTRA ADVANCED METRICS FOR IMPORT & LOCAL
    import_delayed = import_only[import_only["Over All Delay"] > 0]
    dashboard["import_total_value"] = float(import_only["InbValKWD"].sum())
    dashboard["import_delayed_value"] = float(import_delayed["InbValKWD"].sum())
    dashboard["import_avg_delay_days"] = float(import_delayed["Over All Delay"].mean()) if len(import_delayed) > 0 else 0.0
    
    import_val_by_ctry = import_delayed.groupby("Vendor Ctry")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
    dashboard["import_value_by_country"] = import_val_by_ctry.to_dict("records")
    
    import_delay_by_ctry = import_delayed.groupby("Vendor Ctry")["Over All Delay"].mean().sort_values(ascending=False).reset_index()
    dashboard["import_avg_delay_by_country"] = import_delay_by_ctry.to_dict("records")
    
    import_val_by_cat = import_delayed.groupby("Category")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
    dashboard["import_value_by_category"] = import_val_by_cat.to_dict("records")
    
    local_delayed = local_only[local_only["Over All Delay"] > 0]
    dashboard["local_total_value"] = float(local_only["InbValKWD"].sum())
    dashboard["local_delayed_value"] = float(local_delayed["InbValKWD"].sum())
    dashboard["local_avg_delay_days"] = float(local_delayed["Over All Delay"].mean()) if len(local_delayed) > 0 else 0.0
    
    local_val_by_ctry = local_delayed.groupby("Vendor Ctry")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
    dashboard["local_value_by_country"] = local_val_by_ctry.to_dict("records")
    
    local_delay_by_ctry = local_delayed.groupby("Vendor Ctry")["Over All Delay"].mean().sort_values(ascending=False).reset_index()
    dashboard["local_avg_delay_by_country"] = local_delay_by_ctry.to_dict("records")
    
    local_val_by_cat = local_delayed.groupby("Category")["InbValKWD"].sum().sort_values(ascending=False).reset_index()
    dashboard["local_value_by_category"] = local_val_by_cat.to_dict("records")
    
    dashboard["import_top_delayed_vendors"] = (
        import_delayed.groupby("Vendor Name")["PO#"].nunique()
        .sort_values(ascending=False).head(10).reset_index().rename(columns={"PO#": "Delayed POs"}).to_dict("records")
    )
    dashboard["local_top_delayed_vendors"] = (
        local_delayed.groupby("Vendor Name")["PO#"].nunique()
        .sort_values(ascending=False).head(10).reset_index().rename(columns={"PO#": "Delayed POs"}).to_dict("records")
    )
    
    # INSIGHTS
    top_country = df.groupby("Vendor Ctry")["PO#"].nunique().sort_values(ascending=False).reset_index()
    top_vendor = df.groupby("Vendor Name")["PO#"].nunique().sort_values(ascending=False).reset_index()
    top_category = df.groupby("Category")["PO#"].nunique().sort_values(ascending=False).reset_index()
    
    dashboard["insights"] = {
        "highest_country": top_country.iloc[0]["Vendor Ctry"] if len(top_country) > 0 else "N/A",
        "highest_vendor": top_vendor.iloc[0]["Vendor Name"] if len(top_vendor) > 0 else "N/A",
        "highest_category": top_category.iloc[0]["Category"] if len(top_category) > 0 else "N/A",
        "highest_delay_type": max(dashboard["delay_counts"], key=dashboard["delay_counts"].get) if len(dashboard["delay_counts"]) > 0 else "N/A"
    }
    dashboard["top_delay_vendors"] = sorted(vendor_records, key=lambda x: x["Total Delay"], reverse=True)[:15]

    # PO INPUT DELAY ANALYSIS — distinct per PO#
    dashboard["po_input_delay"] = {
        "all": compute_po_input_delay(df),
        "import": compute_po_input_delay(import_only),
        "local": compute_po_input_delay(local_only)
    }

    # ANOMALY DETECTION DATA — Row based
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

    return dashboard

def process_excel_validator(file_path: Path, output_dir: Path, original_filename: str):
    # Load file
    df = pd.read_excel(file_path)
    rows_loaded = len(df)
    columns_loaded = len(df.columns)
    
    # 1. Standardize columns
    column_mapping = {
        "ATP - formula (no negative": "ATP",
        "Vendor Ctry.1": "Vendor Country Name"
    }
    df.rename(columns=column_mapping, inplace=True)
    
    # 2. Update category for project articles and combine D&W Production & doors and windows
    if "Category" in df.columns:
        df["Category"] = df["Category"].fillna("N/A").astype(str).str.strip()
        category_masks = df["Category"].str.upper().isin(["D&W PRODUCTION", "D&W PRODUCTIONS", "DOORS & WINDOWS", "DOORS AND WINDOWS"])
        df.loc[category_masks, "Category"] = "DOORS & WINDOWS"
        
    if "Article" in df.columns and "Category" in df.columns:
        project_mask = df["Article"].astype(str).str.startswith("P-", na=False)
        df.loc[project_mask, "Category"] = "KITCHEN PROJECT"
        
    # 3. Date conversion
    date_columns = ["ATP", "GRP", "AWH", "BAYAN", "ETA", "ETD"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
            
    # 4. Add Row ID
    df.insert(0, "Row_ID", range(1, len(df) + 1))
    
    # 5. Issue tracking columns
    df["Issue_Count"] = 0
    df["Issue_Details"] = ""
    
    # 6. Local / Import
    country_col = "Vendor Country Name" if "Vendor Country Name" in df.columns else ("Vendor Ctry.1" if "Vendor Ctry.1" in df.columns else None)
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
        df["Import_Type"] = df["Import / Local"]
    else:
        df["Import / Local"] = "Import"
        df["Import_Type"] = "Import"
        
    # Create output folders
    summary_folder = output_dir / "Summary"
    atp_folder = output_dir / "ATP_Before_Event"
    gap_folder = output_dir / "Gap_Validation"
    master_folder = output_dir / "Master"
    new_file_folder = output_dir / "new_file"
    
    for folder in [summary_folder, atp_folder, gap_folder, master_folder, new_file_folder]:
        folder.mkdir(parents=True, exist_ok=True)
        
    summary = []
    
    RULE_FRIENDLY_NAMES = {
        "ATP_BEFORE_GRP": "ATP Date is before GRP Date",
        "ATP_BEFORE_AWH": "ATP Date is before AWH Date",
        "ATP_BEFORE_BAYAN": "ATP Date is before BAYAN Date",
        "ATP_BEFORE_ETA": "ATP Date is before ETA Date",
        "ATP_BEFORE_ETD": "ATP Date is before ETD Date",
        "ATP_GRP_GAP_GT_5": "ATP to GRP Gap exceeds 5 Days"
    }

    def register_issue(mask, issue_name, folder):
        count = int(mask.sum())
        friendly = RULE_FRIENDLY_NAMES.get(issue_name, issue_name)
        summary.append({
            "Rule": friendly,
            "Violations": count,
            "Percent": round((count / len(df)) * 100, 2) if len(df) > 0 else 0
        })
        if count == 0:
            return
        df.loc[mask, "Issue_Count"] += 1
        df.loc[mask, "Issue_Details"] += friendly + "; "
        issue_df = df.loc[mask].copy()
        issue_df.to_excel(folder / f"{issue_name}.xlsx", index=False)
        
    # Run Validations
    if "ATP" in df.columns and "GRP" in df.columns:
        register_issue((df["ATP"].notna() & df["GRP"].notna() & (df["ATP"] < df["GRP"])), "ATP_BEFORE_GRP", atp_folder)
    if "ATP" in df.columns and "AWH" in df.columns:
        register_issue((df["ATP"].notna() & df["AWH"].notna() & (df["ATP"] < df["AWH"])), "ATP_BEFORE_AWH", atp_folder)
    if "ATP" in df.columns and "BAYAN" in df.columns:
        register_issue((df["ATP"].notna() & df["BAYAN"].notna() & (df["ATP"] < df["BAYAN"])), "ATP_BEFORE_BAYAN", atp_folder)
    if "ATP" in df.columns and "ETA" in df.columns:
        register_issue((df["ATP"].notna() & df["ETA"].notna() & (df["ATP"] < df["ETA"])), "ATP_BEFORE_ETA", atp_folder)
    if "ATP" in df.columns and "ETD" in df.columns:
        register_issue((df["ATP"].notna() & df["ETD"].notna() & (df["ATP"] < df["ETD"])), "ATP_BEFORE_ETD", atp_folder)
        
    # Gap validation
    if "ATP" in df.columns and "GRP" in df.columns:
        grp_gap = (df["ATP"] - df["GRP"]).dt.days
        df["ATP_GRP_Gap_Days"] = grp_gap
        register_issue((df["ATP"].notna() & df["GRP"].notna() & (grp_gap > 5)), "ATP_GRP_GAP_GT_5", gap_folder)
        
    # Severity classification
    df["Severity"] = np.select(
        [
            (df["Issue_Count"] > 0) & (df["Import_Type"] == "Import"),
            (df["Issue_Count"] > 0) & (df["Import_Type"] == "Local")
        ],
        ["Severe", "Medium"],
        default="No"
    )
    
    # Save Master files
    all_anomalies = df[df["Issue_Count"] > 0].copy()
    all_anomalies.sort_values(by="Issue_Count", ascending=False, inplace=True)
    all_anomalies.to_excel(master_folder / "All_Anomalies.xlsx", index=False)
    
    high_severity = df[df["Severity"] == "Severe"].copy()
    high_severity.to_excel(master_folder / "High_Severity.xlsx", index=False)
    
    df.to_excel(master_folder / "Freight_With_RowID.xlsx", index=False)
    
    # Save clean updated datasets in new_file folder (entire dataset, no Row_ID)
    category_updates = df.copy()
    if "Row_ID" in category_updates.columns:
        category_updates.drop(columns=["Row_ID"], inplace=True)
        
    category_updates.to_excel(new_file_folder / "Category_Updates.xlsx", index=False)
    category_updates.to_excel(new_file_folder / original_filename, index=False)
    
    # Save Summary
    summary_df = pd.DataFrame(summary)
    summary_df.to_excel(summary_folder / "Validation_Summary.xlsx", index=False)
    
    total_violations = int(all_anomalies.shape[0])
    
    return {
        "rows_loaded": rows_loaded,
        "columns_loaded": columns_loaded,
        "violations": total_violations,
        "summary": summary
    }

@app.route("/", methods=["GET"])
def index():
    try:
        with open("templates/index.html", "r", encoding="utf-8") as f:
            content = f.read()
        return render_template_string(content)
    except Exception as e:
        return f"Error loading index template: {str(e)}", 500

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "No file uploaded"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "error": "No selected file"}), 400
        
    action = request.form.get("action", "dashboard")
    
    # Clean up previous temp files to free space
    if TEMP_DIR.exists():
        try:
            shutil.rmtree(TEMP_DIR)
        except Exception:
            pass
            
    # Generate unique ID for this processing session
    session_id = str(uuid.uuid4())
    session_dir = TEMP_DIR / session_id
    upload_path = session_dir / file.filename
    
    session_dir.mkdir(parents=True, exist_ok=True)
    file.save(upload_path)
    
    try:
        if action == "dashboard":
            # 1. Process Excel for Dashboard
            with open(upload_path, "rb") as f:
                dashboard_json = process_excel(f)
            
            # 2. Build dashboard_data.js content
            js_data = f"const dashboardData = {json.dumps(dashboard_json, ensure_ascii=False)};"
            
            # 3. Read the master dashboard.html
            if not os.path.exists("dashboard.html"):
                return jsonify({"success": False, "error": "Master dashboard.html template not found on server."}), 500
                
            with open("dashboard.html", "r", encoding="utf-8") as f:
                html_content = f.read()
                
            # 4. Create ZIP file
            zip_path = session_dir / "freight_analytics_dashboard.zip"
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_archive:
                zip_archive.writestr("dashboard.html", html_content)
                zip_archive.writestr("dashboard_data.js", js_data)
                
            return jsonify({
                "success": True,
                "action": "dashboard",
                "download_url": f"/download/dashboard/{session_id}"
            })
            
        elif action == "validation":
            # 1. Process Excel for Validator
            output_dir = session_dir / "Validation_Outputs"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            stats = process_excel_validator(upload_path, output_dir, file.filename)
            
            # 2. Package output as zip
            shutil.make_archive(str(session_dir / "Validation_Outputs"), "zip", output_dir)
            
            return jsonify({
                "success": True,
                "action": "validation",
                "rows_loaded": stats["rows_loaded"],
                "columns_loaded": stats["columns_loaded"],
                "violations": stats["violations"],
                "summary": stats["summary"],
                "download_url": f"/download/validation/{session_id}"
            })
            
        else:
            return jsonify({"success": False, "error": "Invalid action specified."}), 400
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/download/<action>/<session_id>", methods=["GET"])
def download(action, session_id):
    if action == "dashboard":
        zip_path = TEMP_DIR / session_id / "freight_analytics_dashboard.zip"
        download_name = "freight_analytics_dashboard.zip"
    elif action == "validation":
        zip_path = TEMP_DIR / session_id / "Validation_Outputs.zip"
        download_name = "Validation_Outputs.zip"
    else:
        return "Invalid download action", 400
        
    if not zip_path.exists():
        return "File not found or session expired", 404
        
    return send_file(zip_path, as_attachment=True, download_name=download_name)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5500))
    app.run(host="0.0.0.0", port=port)

