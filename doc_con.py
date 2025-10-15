import camelot
import pandas as pd
import os

def extract_pdf_tables(pdf_path, output_excel_path):
    """
    Extracts all tables from a PDF file using Camelot and saves them 
    into a multi-sheet Excel file.
    """
    
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found at {pdf_path}")
        return

    print(f"Attempting to extract tables from: {pdf_path}")
    
    # -----------------------------------------------------------
    # Core Extraction Step
    # -----------------------------------------------------------
    
    # Use 'lattice' flavor for tables with clear lines (borders)
    # Use 'stream' flavor for tables defined by whitespace (no lines)
    # 'pages="all"' ensures all pages are checked
    try:
        tables = camelot.read_pdf(
            pdf_path, 
            pages='all', 
            flavor='lattice', 
            table_areas=None
        )
    except Exception as e:
        print(f"An error occurred during extraction: {e}")
        return

    # Check if any tables were found
    num_tables = tables.n
    if num_tables == 0:
        print("No tables were detected in the PDF.")
        return

    print(f"\nSuccessfully detected {num_tables} table(s).")
    
    # -----------------------------------------------------------
    # Exporting to Excel
    # -----------------------------------------------------------
    
    # Initialize the Excel writer to handle multiple sheets
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            # Access the table data as a Pandas DataFrame
            df = table.df
            
            # Create a sheet name based on the table's location
            sheet_name = f"Page_{table.page}_Table_{i+1}"
            
            # Write the DataFrame to a new sheet in the Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            # Optional: Print extraction report for the table
            print(f"Exported {sheet_name}. Parsing Report: {table.parsing_report}")

    print(f"\n✅ All tables extracted and saved to: {output_excel_path}")


# --- Configuration ---
# NOTE: Replace 'my_document.pdf' with the actual path to your PDF file.
PDF_FILE = 'my_document.pdf' 

# Set the desired output file name
OUTPUT_FILE = 'extracted_data.xlsx'

# --- Run the function ---
extract_pdf_tables(PDF_FILE, OUTPUT_FILE)
