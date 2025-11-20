import pandas as pd
import os
import sys
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook

# ----- Filter constants -----
CUSTOMER_FILTER = {"customertype": "Consumer", "customersubtype": "Regular"}

EXTRACTS = [
    {"delayreason": "(X) CHC-HOUSE/UNIT CLOSED", "skillset": "Install"},
    {"delayreason": "(X) CHC-HOUSE/UNIT CLOSED", "skillset": "Repair"},
    {"delayreason": "(X) CRES - RESKED  WITH PREFERRED DATE", "skillset": "Install"},
    {"delayreason": "(X) CRES - RESKED  WITH PREFERRED DATE", "skillset": "Repair"},
]

COLUMNS = [
    "workordernumber","customername","customeraddress","customercontact",
    "customertype","customersubtype","skillset","queue","substatus",
    "delaycode","delayreason","delaynotes","lastupdatedate"
]

# ----- Select file -----
Tk().withdraw()  # hide main window
file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

if not file_path:
    messagebox.showinfo("Cancelled", "No file selected. Exiting program.")
    sys.exit()

# ----- Create Extracted folder (auto-rename if exists) -----
base_folder = os.path.join(os.path.dirname(file_path), "Extracted")
output_folder = base_folder
counter = 1
while os.path.exists(output_folder):
    output_folder = f"{base_folder}{counter}"
    counter += 1

os.makedirs(output_folder, exist_ok=True)

# ----- Read all sheets into a single DataFrame -----
df = pd.concat(pd.read_excel(file_path, sheet_name=None), ignore_index=True)

# ----- Clean DataFrame -----
df.columns = df.columns.str.replace('\xa0',' ').str.strip().str.lower().str.replace(' ','_')
df = df.apply(lambda x: x.astype(str).str.strip() if x.dtype == "object" else x)
df['workordernumber'] = df['workordernumber'].replace({'nan': None}).ffill()

# ----- Filter by customer type and subtype -----
for k, v in CUSTOMER_FILTER.items():
    df = df[df.get(k) == v]

# ----- Extract combinations -----
for item in EXTRACTS:
    tmp = df.copy()
    for k, v in item.items():
        tmp = tmp[tmp.get(k) == v]

    tmp = tmp.drop_duplicates()

    # Keep only desired columns that exist
    cols_exist = [c for c in COLUMNS if c.lower().replace(' ','_') in tmp.columns]
    tmp = tmp[cols_exist]

    # Reorder: first 5 columns fixed, then rest
    first_cols = ['workordernumber','customername','customercontact','customeraddress','lastupdatedate']
    rest_cols = [c for c in tmp.columns if c not in first_cols]
    tmp = tmp[first_cols + rest_cols]

    # ----- Save file with auto-rename if exists -----
    skill = item['skillset'].replace(' ', '_')
    reason = item['delayreason'].replace(' ', '_').replace('/', '_')
    base_file = f"{skill}_{reason}.xlsx"
    output_path = os.path.join(output_folder, base_file)
    file_counter = 1
    while os.path.exists(output_path):
        output_path = os.path.join(output_folder, f"{skill}_{reason}_{file_counter}.xlsx")
        file_counter += 1

    tmp.to_excel(output_path, index=False)

    # ----- Auto-fit columns -----
    wb = load_workbook(output_path)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try: 
                max_length = max(max_length, len(str(cell.value)))
            except: 
                pass
        ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(output_path)

    print(f"Saved {len(tmp)} rows to {output_path} (columns auto-fit)")

messagebox.showinfo("Done", f"Extraction complete! Files saved to:\n{output_folder}")
