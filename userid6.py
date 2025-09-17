import pandas as pd

# Load both sheets
sheet1 = pd.read_excel("file1.xlsx", sheet_name="ONSE")        # Has Case + Account Number
sheet2 = pd.read_excel("file2.xlsx", sheet_name="Details")     # Has Acc_No + Note Date + Note + UserID

# Merge: keep all rows from sheet2
merged = sheet2.merge(
    sheet1[["Case", "Account Number"]],
    left_on="Acc_No",
    right_on="Account Number",
    how="left"
)

# Select required columns (Case may be NaN if no match found)
final = merged[["Case", "Account Number", "Note Date", "Note", "UserID"]]

# Replace NaN with blank for Case
final["Case"] = final["Case"].fillna("")

# Save
final.to_excel("final_output.xlsx", index=False)
