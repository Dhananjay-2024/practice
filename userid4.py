import pandas as pd

# Load sheets
sheet1 = pd.read_excel("file1.xlsx", sheet_name="ONSE")
sheet2 = pd.read_excel("file2.xlsx", sheet_name="Details")

# Collapse unique UserIDs per account into a single string
agg_users = sheet2.groupby("Acc_No", as_index=False)["UserID"] \
                  .apply(lambda x: ", ".join(sorted(set(map(str, x)))))

# Left join (keep all rows from sheet1)
merged = sheet1.merge(agg_users, left_on="Account Number", right_on="Acc_No", how="left")

# Drop duplicate key column
merged = merged.drop(columns=["Acc_No"])

# Replace NaN with blank if no UserID exists
merged["UserID"] = merged["UserID"].fillna("")

# Save output
merged.to_excel("output.xlsx", index=False)
