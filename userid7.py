import pandas as pd

# Load both sheets
sheet1 = pd.read_excel("Risk.xlsx", sheet_name="Case Details")   # Has Case ID + Account Number
sheet2 = pd.read_excel("Query.xlsx", sheet_name="Details")       # Has notes + account + userid

# Merge: keep ALL cases from sheet1
merged = sheet1.merge(
    sheet2[[
        "npa_faa_notes_w.account_number",
        "npa_faa_notes_w.notes_datetime",
        "npa_faa_notes_w.notes",
        "userID"
    ]],
    left_on="Account Number",
    right_on="npa_faa_notes_w.account_number",
    how="left"
)

# Select required columns
final = merged[[
    "Case ID",
    "Account Number",
    "npa_faa_notes_w.notes_datetime",
    "npa_faa_notes_w.notes",
    "userID"
]]

# Fill NaN with blanks for missing values
final = final.fillna("")

# Save result
final.to_excel("final_output.xlsx", index=False)
