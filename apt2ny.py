import os
import json
import random
import logging
from datetime import timedelta, datetime
import pandas as pd
from openpyxl import load_workbook, Workbook

# ---------------- Logging ---------------- #
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.StreamHandler()]
)

# ---------------- Config ---------------- #
DATA_DIR = "data_sub"
EXCEL_FILE = "case_data.xlsx"
NOTE_SHEET = "Note Activity"
ACCOUNT_SHEET = "Account Activity"

SAMPLE_SIZE = 5
OUTPUT_DIR = "case_variants"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---- Case selection ----
# CASE_SELECTION = "all"   # all cases
# CASE_SELECTION = 15      # single case
# CASE_SELECTION = (4, 10) # range inclusive
CASE_SELECTION = "all"


# ---------------- Helpers ---------------- #
def ensure_columns(ws):
    """Ensure required columns exist in Note Activity sheet (except Note)."""
    headers = [cell.value for cell in ws[1]]
    required_cols = ["example_id", "bias"]
    for col in required_cols:
        if col not in headers:
            ws.cell(row=1, column=len(headers)+1).value = col
            headers.append(col)
            logging.info(f"Added missing column: {col}")
    return headers


def filter_cases(all_cases):
    """Filter cases based on CASE_SELECTION config."""
    if CASE_SELECTION == "all":
        return all_cases
    elif isinstance(CASE_SELECTION, int):  # single case
        return [CASE_SELECTION] if CASE_SELECTION in all_cases else []
    elif isinstance(CASE_SELECTION, tuple) and len(CASE_SELECTION) == 2:
        low, high = CASE_SELECTION
        return [c for c in all_cases if low <= c <= high]
    else:
        logging.error("Invalid CASE_SELECTION format.")
        return []


def load_bias_records():
    """Load all records grouped by bias from the data directory (no subdirectories)."""
    DATA_DIR = "data"  # update to new directory
    bias_records = {}
    for fname in os.listdir(DATA_DIR):
        if not fname.endswith(".jsonl"):
            continue
        bias_name = os.path.splitext(fname)[0]
        fpath = os.path.join(DATA_DIR, fname)
        logging.info(f"Reading {fpath}")
        records = []
        with open(fpath, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                    note_text = f"{rec.get('context','')} {rec.get('question','')}".strip()
                    records.append({
                        "example_id": rec.get("example_id", ""),
                        "Note": note_text,
                        "bias": bias_name
                    })
                except Exception as e:
                    logging.warning(f"Failed parsing line in {fname}: {e}")
        if records:
            bias_records[bias_name] = records
            logging.info(f"Loaded {len(records)} records for bias={bias_name}")
        else:
            logging.warning(f"No records found in {fpath}")
    return bias_records


def get_case_block(note_df, case_no):
    """Get subset of Note Activity rows for a specific case."""
    case_block = note_df[note_df["Case"] == case_no].copy()
    case_block["Note Date"] = pd.to_datetime(case_block["Note Date"], errors="coerce")
    case_block = case_block.sort_values("Note Date")
    return case_block


def pick_insertion_date(case_block, queue_date):
    """Pick median Note Date within 90 days before Queue In Date."""
    if pd.isna(queue_date):
        logging.warning("Queue In Date is missing, falling back to today.")
        return datetime.today()

    start_date = queue_date - timedelta(days=90)
    valid_dates = case_block[
        (case_block["Note Date"] >= start_date) &
        (case_block["Note Date"] <= queue_date)
    ]["Note Date"].dropna().sort_values()

    if valid_dates.empty:
        fallback = queue_date - timedelta(days=45)
        logging.info(f"No valid dates in window, fallback: {fallback.date()}")
        return fallback

    median_date = valid_dates.iloc[len(valid_dates)//2]
    return median_date


# ---------------- Main Logic ---------------- #
def create_case_variants():
    logging.info("Loading workbook for case list...")
    note_df = pd.read_excel(EXCEL_FILE, sheet_name=NOTE_SHEET)
    acct_df = pd.read_excel(EXCEL_FILE, sheet_name=ACCOUNT_SHEET)
    acct_df["Queue In Date"] = pd.to_datetime(acct_df["Queue In Date"], errors="coerce")

    all_cases = note_df["Case"].dropna().unique().tolist()
    all_cases = [int(c) for c in all_cases if str(c).isdigit()]
    selected_cases = filter_cases(all_cases)

    logging.info(f"Selected cases: {selected_cases}")

    # Load all bias records
    bias_records = load_bias_records()

    # Prepare a DataFrame for the output: start with all selected cases' original rows
    output_df = note_df[note_df["Case"].isin(selected_cases)].copy()

    for case_no in selected_cases:
        logging.info(f"Processing Case {case_no}")
        case_block = get_case_block(note_df, case_no)

        # Get Queue In Date
        q_date = acct_df.loc[acct_df["Case"] == case_no, "Queue In Date"]
        q_date = q_date.iloc[0] if not q_date.empty else pd.NaT

        insert_date = pick_insertion_date(case_block, q_date)

        # For each bias, sample and append variants
        for bias_name, records in bias_records.items():
            if not records:
                continue

            subset = random.sample(records, min(SAMPLE_SIZE, len(records)))
            logging.info(f"Case {case_no}, Bias {bias_name}: {len(subset)} samples")

            for rec in subset:
                # Create a new row for the variant
                new_row = {
                    "Case": case_no,
                    "Note Date": insert_date.strftime("%Y-%m-%d"),
                    "Note": rec["Note"],
                }
                # Add any other columns from the original sheet as empty
                for col in output_df.columns:
                    if col not in new_row:
                        new_row[col] = ""
                output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)

    # Drop 'example_id' and 'bias' columns if present
    output_df = output_df[[col for col in output_df.columns if col not in ("example_id", "bias")]]

    # Sort by Case and Note Date if desired
    output_df = output_df.sort_values(["Case", "Note Date"])

    # Write to Excel (one sheet for all cases)
    out_name = "SelectedCases_variants.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        output_df.to_excel(writer, sheet_name="Variants", index=False)
    logging.info(f"Saved {out_path}")


# ---------------- Run ---------------- #
if __name__ == "__main__":
    create_case_variants()
