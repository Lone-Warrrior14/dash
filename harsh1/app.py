import os
import shutil
import uuid
from pathlib import Path
import pandas as pd
import numpy as np
from flask import Flask, request, jsonify, send_file, render_template

app = Flask(__name__)

# Base temporary directory for processing
TEMP_DIR = Path("temp_processing")

def process_excel(file_path: Path, output_dir: Path, original_filename: str):
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
    
    # 2. Update category for project articles (Article starts with "P-")
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
    
    def register_issue(mask, issue_name, folder):
        count = int(mask.sum())
        summary.append({
            "Rule": issue_name,
            "Violations": count,
            "Percent": round((count / len(df)) * 100, 2) if len(df) > 0 else 0
        })
        if count == 0:
            return
        df.loc[mask, "Issue_Count"] += 1
        df.loc[mask, "Issue_Details"] += issue_name + "; "
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

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded"}), 400
        
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "error": "No selected file"}), 400
        
    # Clean up previous temp files to free space
    if TEMP_DIR.exists():
        try:
            shutil.rmtree(TEMP_DIR)
        except Exception:
            pass
            
    # Generate unique ID for this processing session
    request_id = str(uuid.uuid4())
    session_dir = TEMP_DIR / request_id
    upload_path = session_dir / file.filename
    output_dir = session_dir / "Validation_Outpust"
    
    output_dir.mkdir(parents=True, exist_ok=True)
    file.save(upload_path)
    
    try:
        # Process file
        stats = process_excel(upload_path, output_dir, file.filename)
        
        # Package output as zip
        shutil.make_archive(str(session_dir / "Validation_Outpust"), "zip", output_dir)
        
        return jsonify({
            "success": True,
            "request_id": request_id,
            "filename": file.filename,
            "rows_loaded": stats["rows_loaded"],
            "columns_loaded": stats["columns_loaded"],
            "violations": stats["violations"],
            "summary": stats["summary"]
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/download/<request_id>", methods=["GET"])
def download_zip(request_id):
    zip_path = TEMP_DIR / request_id / "Validation_Outpust.zip"
    if not zip_path.exists():
        return "File not found or session expired", 404
    return send_file(zip_path, as_attachment=True, download_name="Validation_Outpust.zip")

if __name__ == "__main__":
    app.run(debug=True, port=5001)