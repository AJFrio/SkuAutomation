import pandas as pd
import os
from main import process_row_data

def validate_files(excel_path, base_path="C:\\Users\\admin\\Desktop\\V2"):
    """
    Validates that all expected files were created and processes any missing ones.
    
    Args:
        excel_path (str): Path to the Excel file
        base_path (str): Base directory where files should be created
    """
    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(excel_path)
        
        # Sheets to skip
        skip_sheets = ['Dashboard', 'Artworks for Haley to Remake']
        
        # Track missing files
        missing_files = []
        
        # Iterate through each sheet
        for sheet_name in excel_file.sheet_names:
            if sheet_name in skip_sheets:
                print(f"Skipping validation for sheet: {sheet_name}")
                continue
                
            print(f"Validating files for sheet: {sheet_name}")
            
            # Read the current sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Create vendor directory if it doesn't exist
            vendor_path = os.path.join(base_path, sheet_name)
            if not os.path.exists(vendor_path):
                os.makedirs(vendor_path)
            
            # Check each row
            for index, row in df.iloc[1:].iterrows():
                sku = row.iloc[0]  # Get value from column A
                if pd.notna(sku):  # Only process if SKU is not empty
                    expected_file = os.path.join(vendor_path, f"{sku}.svg")
                    if not os.path.exists(expected_file):
                        missing_files.append((sku, sheet_name))
                        print(f"Missing file: {expected_file}")
        
        # Process missing files
        if missing_files:
            print(f"\nFound {len(missing_files)} missing files. Processing them now...")
            for sku, vendor in missing_files:
                print(f"Processing missing file for SKU: {sku}, Vendor: {vendor}")
                process_row_data(sku, vendor)
        else:
            print("\nAll files are present and accounted for!")
            
    except Exception as e:
        print(f"Error during validation: {e}")

if __name__ == "__main__":
    excel_path = 'product.xlsx'
    validate_files(excel_path)
