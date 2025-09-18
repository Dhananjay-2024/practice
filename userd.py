import pandas as pd

# Load Sheet1 from Excel (Case ID + Account Number)
sheet1 = pd.read_excel("Risk.xlsx", sheet_name="Case Details")

# Load Sheet2 from CSV (Notes + account + userid)
sheet2 = pd.read_csv("Query.csv")

# Debug: print actual column names in CSV
print("Sheet2 columns:", sheet2.columns.tolist())


# Rename columns in sheet2 to match sheet1
sheet2 = sheet2.rename(columns={
    "npa_faa_notes_w.account_number": "Account Number",
    "npa_faa_notes_w.notes_datetime": "Note Date",
    "Note": "Note",
    "userID": "userID"
})

# Skip rows where userID starts with 'FAAP' or 'FRDASST1'
sheet2 = sheet2[~sheet2['userID'].astype(str).str.startswith(('FAAP', 'FRDASST1'))]


# Merge (keep ALL rows from Sheet1, even if no matching row in Sheet2)
merged = sheet1[["Case ID", "Account Number"]].merge(
    sheet2[["Account Number", "Note Date", "Note", "userID"]],
    on="Account Number",
    how="left"
)


# Select final columns
final = merged[[
    "Case ID",
    "Account Number",
    "Note Date",
    "Note",
    "userID"
]]

# Replace NaN with blanks
final = final.fillna("")

# Save to Excel
final.to_excel("final_output.xlsx", index=False)
