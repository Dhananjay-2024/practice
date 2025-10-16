import camelot
import pandas as pd
import os
import glob

# --- CONFIGURATION ---
# 1. Directory containing your source PDF files
INPUT_DIR = './pdfs_to_process' 
# 2. Directory where the output Excel files will be saved
OUTPUT_DIR = './extracted_data' 
# 3. The name of the single sheet in each output Excel file
SHEET_NAME = 'Extracted Tables' 

def batch_extract_to_separate_files(input_dir, output_dir, sheet_name):
    """
    Processes all PDF files in a directory, extracts tables, and saves 
    the results for each PDF into a separate Excel file in the output directory.
    All tables from one PDF are stacked vertically onto a single sheet.
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
            tables = camelot.read_pdf(
                pdf_path, 
                pages='all', 
                flavor='lattice', # Use 'stream' if tables lack lines
                strip_text='\n' # Remove newline characters within cells
            )
        except Exception as e:
            print(f"   ⚠️ Could not process {pdf_filename}. Error: {e}")
            continue

        num_tables = tables.n
        if num_tables == 0:
            print("   (No tables detected on any page. Skipping export.)")
            continue
            
        print(f"   Detected {num_tables} table(s). Combining and saving...")

        # 4. Combine all extracted tables into one master DataFrame
        #    We need a list of DataFrames for concatenation
        all_dfs = [table.df for table in tables]
        
        # Use pandas.concat to stack all DataFrames vertically
        # ignore_index=True resets the row numbering
        combined_df = pd.concat(all_dfs, ignore_index=True)
        
        # 5. Save the combined DataFrame to the single sheet Excel file
        try:
            combined_df.to_excel(output_excel_path, sheet_name=sheet_name, index=False, header=False)
            print(f"   ✅ Successfully saved {num_tables} table(s) to: {output_excel_path}")
        except Exception as e:
            print(f"   ❌ Error saving combined DataFrame to Excel: {e}")


# --- Execution ---
# Ensure the input directory exists before running
os.makedirs(INPUT_DIR, exist_ok=True)
batch_extract_to_separate_files(INPUT_DIR, OUTPUT_DIR, SHEET_NAME)
