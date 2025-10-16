import camelot
import pandas as pd
import os
import glob
from collections import defaultdict

# --- CONFIGURATION ---
# 1. Directory containing your source PDF files
INPUT_DIR = './pdfs_to_process' 
# 2. Directory where the output Excel files will be saved
OUTPUT_DIR = './extracted_data' 

def batch_extract_to_pagewise_sheets(input_dir, output_dir):
    """
    Processes all PDF files in a directory. For each PDF, it creates a 
    separate Excel file where tables from different pages are saved into 
    separate sheets (e.g., Sheet_Page_1, Sheet_Page_2, etc.).
    """
    
    # 1. Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    print(f"Output directory created/verified: {output_dir}")

    # 2. Find all PDF files in the input directory
    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf'))
    
    if not pdf_files:
        print(f"❌ No PDF files found in {input_dir}. Please check the path and contents.")
        return

    print(f"Found {len(pdf_files)} PDF files to process.\n")
    
    # 3. Process each PDF file individually
    for pdf_path in pdf_files:
        pdf_filename = os.path.basename(pdf_path)
        base_name = os.path.splitext(pdf_filename)[0]
        
        # Define the output path for the Excel file
        output_excel_path = os.path.join(output_dir, f"{base_name}.xlsx")
        
        print(f"--- Processing: {pdf_filename} ---")
        
        try:
            # Attempt to extract all tables from the PDF
            # Use 'stream' or 'lattice' based on table type
            tables = camelot.read_pdf(
                pdf_path, 
                pages='all', 
                flavor='lattice',
                strip_text='\n'
            )
        except Exception as e:
            print(f"   ⚠️ Could not process {pdf_filename}. Error: {e}")
            continue

        num_tables = tables.n
        if num_tables == 0:
            print("   (No tables detected on any page. Skipping export.)")
            continue
            
        print(f"   Detected {num_tables} table(s) across different pages.")

        # 4. Group tables by their page number
        # We use defaultdict to make it easy to append DataFrames to a list per page
        tables_by_page = defaultdict(list)
        for table in tables:
            tables_by_page[table.page].append(table.df)

        # 5. Create and save the multi-sheet Excel file
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                for page_num, list_of_dfs in tables_by_page.items():
                    # Stack all tables from the current page vertically
                    combined_page_df = pd.concat(list_of_dfs, ignore_index=True)
                    
                    # Define the sheet name using the page number
                    sheet_name = f"Page_{page_num}"
                    
                    # Write the combined page data to a sheet
                    combined_page_df.to_excel(
                        writer, 
                        sheet_name=sheet_name, 
                        index=False, 
                        header=False
                    )
                    
                    print(f"   > Saved tables from Page {page_num} to sheet: {sheet_name}")

            print(f"   ✅ Successfully saved all page-wise data to: {output_excel_path}")
            
        except Exception as e:
            print(f"   ❌ Error saving to Excel for {pdf_filename}: {e}")

# --- Execution ---
# Ensure the input directory exists before running
os.makedirs(INPUT_DIR, exist_ok=True)
batch_extract_to_pagewise_sheets(INPUT_DIR, OUTPUT_DIR)
