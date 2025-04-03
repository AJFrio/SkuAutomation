import pandas as pd
import pyautogui as pg
import time
import os

base_path = "C:\\Users\\admin\\Desktop\\V2"

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

def process_excel_file(excel_path):
    """
    Process an Excel file by iterating through each sheet and processing rows.
    
    Args:
        excel_path (str): Path to the Excel file to process
    """
    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(excel_path)
        
        # Sheets to skip
        skip_sheets = ['Dashboard', 'Artworks for Haley to Remake']
        
        # Iterate through each sheet
        for sheet_name in excel_file.sheet_names:
            if sheet_name in skip_sheets:
                print(f"Skipping sheet: {sheet_name}")
                continue
                
            print(f"Processing sheet: {sheet_name}")
            
            # Read the current sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Skip the first row and process each row
            for index, row in df.iloc[1:].iterrows():
                sku = row.iloc[0]  # Get value from column A
                if pd.notna(sku):  # Only process if SKU is not empty
                    process_row_data(sku, sheet_name)
                    time.sleep(1)  # Add a small delay between rows
                    
    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    excel_path = 'product.xlsx'
    process_excel_file(excel_path)
