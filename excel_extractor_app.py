import pandas as pd
import os
from io import BytesIO
from zipfile import ZipFile
import streamlit as st
from openpyxl import load_workbook

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

st.title("Excel Extractor Tool")
st.write("Upload an Excel file, and download the filtered outputs automatically.")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = pd.concat(pd.read_excel(uploaded_file, sheet_name=None), ignore_index=True)
    df.columns = df.columns.str.replace('\xa0',' ').str.strip().str.lower().str.replace(' ','_')
    df = df.apply(lambda x: x.astype(str).str.strip() if x.dtype == "object" else x)
    df['workordernumber'] = df['workordernumber'].replace({'nan': None}).ffill()

    for k, v in CUSTOMER_FILTER.items():
        df = df[df.get(k) == v]

    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "a") as zip_file:
        for item in EXTRACTS:
            tmp = df.copy()
            for k, v in item.items():
                tmp = tmp[tmp.get(k) == v]

            tmp = tmp.drop_duplicates()
            cols_exist = [c for c in COLUMNS if c.lower().replace(' ','_') in tmp.columns]
            tmp = tmp[cols_exist]

            first_cols = ['workordernumber','customername','customercontact','customeraddress','lastupdatedate']
            rest_cols = [c for c in tmp.columns if c not in first_cols]
            tmp = tmp[first_cols + rest_cols]

            skill = item['skillset'].replace(' ', '_')
            reason = item['delayreason'].replace(' ', '_').replace('/', '_')
            output_name = f"{skill}_{reason}.xlsx"

            excel_buffer = BytesIO()
            tmp.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)

            # Auto-fit columns
            wb = load_workbook(excel_buffer)
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
            wb.save(excel_buffer)
            excel_buffer.seek(0)

            zip_file.writestr(output_name, excel_buffer.read())

    zip_buffer.seek(0)
    st.download_button(
        label="Download All Extracted Files",
        data=zip_buffer,
        file_name="Extracted_Files.zip",
        mime="application/zip"
    )
    st.success("Extraction complete! Click the button above to download all files.")
