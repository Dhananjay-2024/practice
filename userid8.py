import pandas as pd

# Load both sheets
sheet1 = pd.read_excel("Risk.xlsx", sheet_name="Case Details")   # Case ID + Account Number
sheet2 = pd.read_excel("Query.xlsx", sheet_name="Details")       # Notes + account + userid

# Print column names to debug
print("Sheet2 columns:", sheet2.columns.tolist())

# Rename Sheet2 columns for consistency
sheet2 = sheet2.rename(columns={
    sheet2.columns[0]: "account_number", 
    sheet2.columns[1]: "notes_datetime",
    sheet2.columns[2]: "notes",
    sheet2.columns[3]: "userID"
})

# Merge (keep all cases from sheet1)
merged = sheet1.merge(
    sheet2,
    left_on="Account Number",
    right_on="account_number",
    how="left"
)

# Select final columns
final = merged[[
    "Case ID", 
    "Account Number", 
    "notes_datetime", 
    "notes", 
    "userID"
]]

# Fill NaN with blanks
final = final.fillna("")

# Save output
final.to_excel("final_output.xlsx", index=False)
