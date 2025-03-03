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
    with open("data.txt", "r") as file:
        steps_data = file.read().strip()
    
    # Split by colon to get individual coordinate pairs
    steps = steps_data.split(':')
    
    """
    Process data for a single row.
    
    Args:
        sku (str): The SKU value from the row
        vendor (str): The Vendor value from the row
    """
    #Sku
    coords = steps[0].strip('[]').split(',')
    pg.click(x=int(coords[0]), y=int(coords[1]), clicks=3)
    time.sleep(.2)
    pg.typewrite(f"{sku}\\n")
    time.sleep(.2)
    #Export
    pg.press("ctrl", "e")
    time.sleep(2)
    #Vendor Path
    coords = steps[1].strip('[]').split(',')
    pg.click(x=int(coords[0]), y=int(coords[1]))
    time.sleep(.2)
    if os.path.exists(f"{base_path}/{vendor}"):
        pg.typewrite(f"{base_path}/{vendor}")
    else:
        os.makedirs(f"{base_path}/{vendor}")
        pg.typewrite(f"{base_path}/{vendor}")
    time.sleep(.2)
    pg.press("enter")
    time.sleep(.2)
    #Sku
    coords = steps[2].strip('[]').split(',')
    pg.click(x=int(coords[0]), y=int(coords[1]))
    time.sleep(.2)
    pg.typewrite(f"{sku}")
    coords = steps[3].strip('[]').split(',')
    pg.click(x=int(coords[0]), y=int(coords[1]))
    time.sleep(.2)
    #Reset 
    pg.press("ctrl", "o")
    time.sleep(.2)
    coords = steps[4].strip('[]').split(',')
    pg.click(x=int(coords[0]), y=int(coords[1]))
    time.sleep(.2)
    pg.press("enter")
    time.sleep(3)

    
    

if __name__ == "__main__":
    # Replace with your actual file path
    excel_file_path = "your_excel_file.xlsx"
    process_excel_data(excel_file_path)
