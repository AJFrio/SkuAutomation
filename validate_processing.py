import pandas as pd
import os
import time
from main import process_row_data, base_path

def validate_processing(excel_file_path):
    """
    Validates that all SKUs from the Excel file have corresponding files in their vendor folders.
    Automatically reruns processing for any missing files.
    
    Args:
        excel_file_path (str): Path to the original Excel file
    """
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path)
        
        # Check if required columns exist
        if 'SKU' not in df.columns or 'Vendor' not in df.columns:
            print("Error: Excel file must contain 'SKU' and 'Vendor' columns")
            return
        
        failed_items = []
        
        # Check each row
        for index, row in df.iterrows():
            sku = row['SKU']
            vendor = row['Vendor']
            
            # Skip rows with missing data
            if pd.isna(sku) or pd.isna(vendor):
                print(f"Skipping validation for row {index}: Missing SKU or Vendor data")
                continue
            
            # Construct the expected file path
            vendor_path = os.path.join(base_path, str(vendor))
            expected_file = os.path.join(vendor_path, f"{str(sku)}.svg")  # Assuming PDF output
            
            # Check if file exists
            if not os.path.exists(expected_file):
                print(f"Missing file for SKU: {sku} in Vendor folder: {vendor}")
                failed_items.append((sku, vendor))
        
        # If there are failed items, reprocess them automatically
        if failed_items:
            print(f"\nFound {len(failed_items)} items that need to be reprocessed.")
            print("\nStarting reprocessing in 3 seconds...")
            print("Please ensure the target application is ready.")
            time.sleep(3)
            
            # Reprocess failed items
            for sku, vendor in failed_items:
                print(f"\nReprocessing SKU: {sku} for Vendor: {vendor}")
                process_row_data(sku, vendor)
                time.sleep(1)  # Add small delay between items
            
            print("\nReprocessing complete!")
            
            # Validate the reprocessed items
            still_failed = []
            for sku, vendor in failed_items:
                vendor_path = os.path.join(base_path, str(vendor))
                expected_file = os.path.join(vendor_path, f"{str(sku)}.svg")
                if not os.path.exists(expected_file):
                    still_failed.append((sku, vendor))
            
            if still_failed:
                print("\nWARNING: The following items still failed after reprocessing:")
                for sku, vendor in still_failed:
                    print(f"SKU: {sku}, Vendor: {vendor}")
            else:
                print("\nAll items were successfully reprocessed!")
        else:
            print("\nValidation complete! All files were processed successfully.")
            
    except Exception as e:
        print(f"Error during validation: {e}")

if __name__ == "__main__":
    # Get the Excel file path from user input
    excel_file_path = input("Enter the path to your Excel file for validation: ")
    
    # Validate file exists and has correct extension
    if not os.path.exists(excel_file_path):
        print(f"Error: File '{excel_file_path}' not found")
    elif not excel_file_path.lower().endswith(('.xlsx', '.xls')):
        print("Error: File must be an Excel file (.xlsx or .xls)")
    else:
        validate_processing(excel_file_path) 