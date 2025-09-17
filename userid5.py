import pandas as pd

# Load sheets
sheet1 = pd.read_excel("file1.xlsx", sheet_name="ONSE")        # Has Case + Account Number
sheet2 = pd.read_excel("file2.xlsx", sheet_name="Details")     # Has Acc_No + Note Date + Note + UserID

# Merge on account number
merged = sheet1.merge(
    sheet2[["Acc_No", "Note Date", "Note", "UserID"]],
    left_on="Account Number",
    right_on="Acc_No",
    how="inner"   # keep only matching rows
)

# Select required columns
final = merged[["Case", "Account Number", "Note Date", "Note", "UserID"]]

# Save to Excel
final.to_excel("final_output.xlsx", index=False)
