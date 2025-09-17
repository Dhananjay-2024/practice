import pandas as pd

# Load both sheets
sheet1 = pd.read_excel("Risk.xlsx", sheet_name="Case Details")   # has "Case ID" + "Account Number"
sheet2 = pd.read_excel("Query.xlsx", sheet_name="Details")       # has "npa_faa_notes_w.account_number", "npa_faa_notes_w.notes_datetime", "Note", "userID"

# Rename columns in sheet2 to simpler names for merging
sheet2 = sheet2.rename(columns={
    "npa_faa_notes_w.account_number": "Account Number",
    "npa_faa_notes_w.notes_datetime": "Note Date",
    "Note": "Note",
    "userID": "userID"
})

# Merge (keep ALL cases from sheet1, even if no notes/userID)
merged = sheet1.merge(
    sheet2,
    on="Account Number",
    how="left"
)

# Select only required columns
final = merged[[
    "Case ID",
    "Account Number",
    "Note Date",
    "Note",
    "userID"
]]

# Fill NaN with blank
final = final.fillna("")

# Save to Excel
final.to_excel("final_output.xlsx", index=False)
