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
    # Keep example_id in output, remove only bias
    headers_to_keep = [h for h in headers if h != "bias"]
    # Explicitly include example_id near the front for clarity
    combined_headers = ["Case", "Bias", "Variant", "example_id"] + [h for h in headers_to_keep if h != "example_id"]

    all_variant_rows = [combined_headers]

    for case_no in selected_cases:
        case_block = get_case_block(note_df, case_no)
        q_date = acct_df.loc[acct_df["Case"] == case_no, "Queue In Date "]
        q_date = q_date.iloc[0] if not q_date.empty else pd.NaT

        variant_counter = 1  # <-- Start variant numbering per case

        for bias_name, records in bias_records.items():
            if not records:
                continue
            subset = random.sample(records, min(SAMPLE_SIZE, len(records)))
            for rec in subset:
                # Start from the original notes for each variant
                variant_block = case_block.copy()
                # Insert the new note at the correct date
                insert_date = pick_insertion_date(variant_block, q_date)
                new_note_row = {h: None for h in headers}
                case_id = f"{case_no}_{rec['example_id']}_{bias_name}"  # <-- Unique case variant ID
                new_note_row["Case"] = case_id
                new_note_row["Note Date "] = insert_date.strftime("%Y-%m-%d")
                new_note_row["Note"] = rec["Note"]
                new_note_row["example_id"] = rec.get("example_id", "")  # <-- Include example_id

                # Insert the new note
                variant_block = pd.concat(
                    [variant_block, pd.DataFrame([new_note_row])],
                    ignore_index=True
                )
                variant_block["Note Date "] = pd.to_datetime(variant_block["Note Date "], errors="coerce")
                variant_block = variant_block.sort_values("Note Date ")

                # Output all notes for this variant
                for _, row in variant_block.iterrows():
                    filtered_row = [row.get(h) for h in headers_to_keep]
                    all_variant_rows.append([case_id, bias_name, variant_counter, rec.get("example_id", "")] + filtered_row)

                variant_counter += 1  # <-- Increment for next variant

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
