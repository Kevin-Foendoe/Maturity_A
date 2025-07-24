import os
import pandas as pd

# Set the directory path
folder_path = r"C:\Users\foendoe.kevin\Documents\Excel\Cassy\Maturity - Send to Companies"
output_folder = r"C:\Users\foendoe.kevin\Documents\MyFiles\MyPythonProjects\Maturity_A\Excel Output"
os.makedirs(output_folder, exist_ok=True)

# List to store row data
rows = []

# Loop through all .xlsx files
for filename in os.listdir(folder_path):
    if filename.lower().endswith(".xlsx"):
        # Extract G-number (everything up to first space)
        gnr = filename.split(" ")[0]
        note = os.path.splitext(filename)[0]  # remove .xlsx extension
        rows.append({"Gnr": gnr, "Status": "","Note": note, })

# Create DataFrame
df = pd.DataFrame(rows, columns=["Gnr", "Note", "Status"])

# Save to Excel
output_path = os.path.join(output_folder, "Gnr_Overview.xlsx")
df.to_excel(output_path, index=False)

print(f"✅ Excel file created at: {output_path}")
