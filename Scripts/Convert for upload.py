import os
import shutil
import pandas as pd
import re
from datetime import datetime

# === Constants ===
input_dir = "Finalize For upload/Input"
done_dir = os.path.join(input_dir, "Done")
output_dir = "Finalize For upload/Output"
template_file = "Template/DATA_MBR_UPDATE.csv"
log_file = "Log/logs.txt"

os.makedirs(done_dir, exist_ok=True)

# === Mapping Excel -> Template CSV Columns ===
column_mapping = {
    "Contract No": "Cont No",
    "Mbr No": "Mbr No",
    "email_address": "Email",
    "PhoneNumber": "Phone Number",
    "national_id": "Natlidno",
    "Address line 1": "Line 1",
    "Address line 2": "Line 2",
    "Address line 3": "Line 3",
    "Post Code": "Post Code",
    "City": "City",
    "State": "State",
    "Country": "Country",
    "Type": "Type"
}

# === Step 1: Find Excel File ===
excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xls", ".xlsx"))]
if len(excel_files) != 1:
    raise Exception(f"Expected exactly one Excel file in {input_dir}, found {len(excel_files)}")

excel_path = os.path.join(input_dir, excel_files[0])
excel_filename = excel_files[0]

# === Step 2: Load Excel and Try to Get Gnr ===
df_excel = pd.read_excel(excel_path)

# Try from filename
match = re.search(r'(G\d{6})', excel_filename)
gnr = match.group(1) if match else None

# If not in filename, search in content
if not gnr:
    for col in df_excel.columns:
        if df_excel[col].astype(str).str.contains(r'G\d{6}').any():
            gnr = df_excel[col].astype(str).str.extract(r'(G\d{6})').dropna().values[0][0]
            break

if not gnr:
    raise Exception("Gnr (e.g., G002611) not found in filename or Excel content.")

# === Step 3: Load Template CSV ===
df_template = pd.read_csv(template_file)

# === Step 4: Map & Insert Data ===
for src_col, dest_col in column_mapping.items():
    if src_col in df_excel.columns:
        df_template[dest_col] = df_excel[src_col]
    else:
        df_template[dest_col] = ""  # Fill with blank if missing

# === ✅ FIX: Handle empty Country if needed ===
# If template column "Country" is empty and variable 'Country' is available
if "Country" in df_template.columns:
    empty_country_mask = df_template["Country"].astype(str).str.strip() == ""
    if empty_country_mask.any():
        # Try to use the value from df_excel['Country'] if possible (from variable)
        # Determine the country by scanning original Excel column if it exists
        detected_country = None
        if "Country" in df_excel.columns:
            unique_vals = df_excel["Country"].dropna().unique()
            if len(unique_vals) == 1:
                detected_country = unique_vals[0].strip().lower()
        
        # Apply logic
        if detected_country == "curacao":
            df_template.loc[empty_country_mask, "Country"] = "CURACAO"
        elif detected_country == "aruba":
            df_template.loc[empty_country_mask, "Country"] = "ARUBA"

# === Step 5: Save to Output Directory ===
output_filename = f"{gnr} - Compass Upload.csv"
output_path = os.path.join(output_dir, output_filename)
df_template.to_csv(output_path, index=False)

# === Step 6: Log the operation ===
with open(log_file, "a") as log:
    log.write(f"{excel_filename}  --------  {output_filename}\n")

print(f"✅ Processed {output_filename}")

# ✅ MOVE the Excel file to Input\Done ONLY IF successful
shutil.move(excel_path, os.path.join(done_dir, os.path.basename(excel_path)))
