import pandas as pd
import pyautogui as pg
import time
import os

base_path = "C:\\Users\\josep\\OneDrive\\Desktop\\"

def process_excel_data(file_path):
    """
    Process an Excel file to extract and work with SKU and Vendor data from each row.
    
    Args:
        file_path (str): Path to the Excel file
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Check if required columns exist
        if 'SKU' not in df.columns or 'Vendor' not in df.columns:
            print("Error: Excel file must contain 'SKU' and 'Vendor' columns")
            return
        
        # Iterate through each row
        for index, row in df.iterrows():
            sku = row['SKU']
            vendor = row['Vendor']
            
            # Skip rows with missing data
            if pd.isna(sku) or pd.isna(vendor):
                print(f"Skipping row {index}: Missing SKU or Vendor data")
                continue
            
            print(f"Processing row {index}: SKU={sku}, Vendor={vendor}")
            
            
            process_row_data(sku, vendor)
            
        print("Processing complete!")
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")

def process_row_data(sku, vendor):
    try:
        with open("data.txt", "r") as file:
            steps_data = file.read().strip()
        
        # Split by colon to get individual coordinate pairs
        steps = steps_data.split(':')
        
        # Validate we have enough steps
        if len(steps) < 5:
            print(f"Error: Not enough coordinate steps in data.txt (found {len(steps)}, need at least 5)")
            return
            
        """
        Process data for a single row.
        
        Args:
            sku (str): The SKU value from the row
            vendor (str): The Vendor value from the row
        """
        # Sku
        try:
            coords = steps[0].strip('[]').split(',')
            pg.click(x=int(coords[0]), y=int(coords[1]), clicks=3)
            time.sleep(0.2)
            pg.typewrite(str(sku) + "\\n")  # Using str() to handle non-string values
            time.sleep(0.2)
            
            # Export
            pg.hotkey("ctrl", "e")  # Use hotkey instead of press for key combinations
            time.sleep(2)
            
            # Vendor Path
            coords = steps[1].strip('[]').split(',')
            pg.click(x=int(coords[0]), y=int(coords[1]))
            time.sleep(0.2)
            
            vendor_path = f"{base_path}\\{vendor}"
            if not os.path.exists(vendor_path):
                os.makedirs(vendor_path)
            pg.typewrite(vendor_path)
            time.sleep(0.2)
            pg.press("enter")
            time.sleep(0.2)
            
            # Sku
            coords = steps[2].strip('[]').split(',')
            pg.click(x=int(coords[0]), y=int(coords[1]))
            time.sleep(0.2)
            pg.typewrite(str(sku))
            
            coords = steps[3].strip('[]').split(',')
            pg.click(x=int(coords[0]), y=int(coords[1]))
            time.sleep(0.2)
            
            # Reset 
            pg.hotkey("ctrl", "o")
            time.sleep(0.2)
            coords = steps[4].strip('[]').split(',')
            pg.click(x=int(coords[0]), y=int(coords[1]))
            time.sleep(0.2)
            pg.press("enter")
            time.sleep(3)
            
            print(f"Successfully processed SKU: {sku} for Vendor: {vendor}")
            
        except IndexError:
            print(f"Error: Invalid coordinate format in data.txt")
        except ValueError:
            print(f"Error: Coordinates must be integers in data.txt")
            
    except FileNotFoundError:
        print("Error: data.txt file not found")
    except Exception as e:
        print(f"Error processing row data: {e}")

if __name__ == "__main__":
    # Get the Excel file path from user input
    excel_file_path = input("Enter the path to your Excel file: ")
    
    # Validate file exists and has correct extension
    if not os.path.exists(excel_file_path):
        print(f"Error: File '{excel_file_path}' not found")
    elif not excel_file_path.lower().endswith(('.xlsx', '.xls')):
        print("Error: File must be an Excel file (.xlsx or .xls)")
    else:
        # Add a confirmation before starting automation
        print("\nWARNING: This script will control your mouse and keyboard.")
        print("Make sure the target application is open and ready.")
        confirmation = input("Type 'yes' to continue: ")
        
        if confirmation.lower() == 'yes':
            print("Starting in 3 seconds...")
            time.sleep(3)
            process_excel_data(excel_file_path)
        else:
            print("Operation cancelled.")
