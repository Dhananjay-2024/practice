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

CASE_SELECTION = "all"

# ---------------- Helpers ---------------- #
def ensure_columns(headers):
    """Ensure required columns exist in Note Activity sheet (except Note)."""
    # Removed example_id and bias columns as per user request
    return headers

def filter_cases(all_cases):
    if CASE_SELECTION == "all":
        return all_cases
    elif isinstance(CASE_SELECTION, int):
        return [CASE_SELECTION] if CASE_SELECTION in all_cases else []
    elif isinstance(CASE_SELECTION, tuple) and len(CASE_SELECTION) == 2:
        low, high = CASE_SELECTION
        return [c for c in all_cases if low <= c <= high]
    else:
        logging.error("Invalid CASE_SELECTION format.")
        return []

def load_bias_records():
    DATA_DIR = "data"
    bias_records = {}
    for fname in os.listdir(DATA_DIR):
        if not fname.endswith(".jsonl"):
            continue
        bias_name = os.path.splitext(fname)[0]
        fpath = os.path.join(DATA_DIR, fname)
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
                except Exception:
                    continue
        if records:
            bias_records[bias_name] = records
    return bias_records

def get_case_block(note_df, case_no):
    case_block = note_df[note_df["Case"] == case_no].copy()
    case_block["Note Date "] = pd.to_datetime(case_block["Note Date "], errors="coerce")
    case_block = case_block.sort_values("Note Date ")
    return case_block

def pick_insertion_date(case_block, queue_date):
    if pd.isna(queue_date):
        return datetime.today()
    start_date = queue_date - timedelta(days=90)
    valid_dates = case_block[
        (case_block["Note Date "] >= start_date) &
        (case_block["Note Date "] <= queue_date)
    ]["Note Date "].dropna().sort_values()
    if not valid_dates.empty:
        return valid_dates.iloc[len(valid_dates)//2]
    fallback = queue_date - timedelta(days=45)
    all_dates = case_block["Note Date "].dropna().sort_values()
    if not all_dates.empty:
        return all_dates.iloc[len(all_dates)//2]
    return datetime.today()

# ---------------- Main Logic ---------------- #
def create_case_variants():
    logging.info("Loading workbook for case list...")
    note_df = pd.read_excel(EXCEL_FILE, sheet_name=NOTE_SHEET)
    acct_df = pd.read_excel(EXCEL_FILE, sheet_name=ACCOUNT_SHEET)
    acct_df["Queue In Date "] = pd.to_datetime(acct_df["Queue In Date "], errors="coerce")

    all_cases = note_df["Case"].dropna().unique().tolist()
    all_cases = [int(c) for c in all_cases if str(c).isdigit()]
    selected_cases = filter_cases(all_cases)

    logging.info(f"Selected cases: {selected_cases}")

    bias_records = load_bias_records()

    # Prepare headers
    headers = list(note_df.columns)
    headers = ensure_columns(headers)
    # Remove example_id and bias from headers_to_keep and combined_headers
    headers_to_keep = [h for h in headers if h not in ("example_id", "bias")]
    combined_headers = ["Case", "Bias", "Variant"] + headers_to_keep

    all_variant_rows = [combined_headers]

    for case_no in selected_cases:
        case_block = get_case_block(note_df, case_no)
        q_date = acct_df.loc[acct_df["Case"] == case_no, "Queue In Date "]
        q_date = q_date.iloc[0] if not q_date.empty else pd.NaT
        insert_date = pick_insertion_date(case_block, q_date)

        for bias_name, records in bias_records.items():
            if not records:
                continue
            subset = random.sample(records, min(SAMPLE_SIZE, len(records)))
            variant_block = case_block.copy()
            for idx, rec in enumerate(subset, start=1):
                new_note_row = {h: None for h in headers}
                new_note_row["Case"] = case_no
                new_note_row["Note Date "] = insert_date.strftime("%Y-%m-%d")
                new_note_row["Note"] = rec["Note"]
                # Removed example_id and bias from new_note_row
                variant_block = pd.concat(
                    [variant_block, pd.DataFrame([new_note_row])],
                    ignore_index=True
                )
                variant_block["Note Date "] = pd.to_datetime(variant_block["Note Date "], errors="coerce")
                variant_block = variant_block.sort_values("Note Date ")
                for _, row in variant_block.iterrows():
                    filtered_row = [row.get(h) for h in headers_to_keep]
                    all_variant_rows.append([case_no, bias_name, idx] + filtered_row)

    # Write all variants to a single Excel sheet
    if len(all_variant_rows) > 1:
        wb_all = Workbook()
        ws_all = wb_all.active
        ws_all.title = "All_Case_Variants"
        for row in all_variant_rows:
            ws_all.append(row)
        out_path = os.path.join(OUTPUT_DIR, "All_Case_Variants.xlsx")
        wb_all.save(out_path)
        logging.info(f"Saved all variants to {out_path}")

if __name__ == "__main__":
    create_case_variants()