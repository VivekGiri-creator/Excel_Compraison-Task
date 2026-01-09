import openpyxl
import os
import time

EXCEL_FILE = "Batch 5 - 08 Jan-Vivek.xlsx"

def process_excel():
    """Process Excel file and update Status column"""
    try:
        print(f"Opening file: {EXCEL_FILE}")
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        
        print(f"Total rows: {ws.max_row}")
        
        # Process each row (skip header row 1)
        for row in range(2, ws.max_row + 1):
            try:
                # Column references:  
                # D = Dell Product (4)
                # E = SNOW Product (5)
                # F = Dell Publisher (6)
                # G = SNOW Publisher (7)
                # K = Status (11)
                # L = Prompt (12)
                
                dell_product = ws.cell(row=row, column=4).value
                snow_product = ws.cell(row=row, column=5).value
                dell_publisher = ws.cell(row=row, column=6).value
                snow_publisher = ws.cell(row=row, column=7).value
                status_cell = ws.cell(row=row, column=11)
                prompt = ws.cell(row=row, column=12).value
                
                # Skip if already processed
                if status_cell.value in ["Match", "Not Match"]:
                    print(f"Row {row}: Already processed ({status_cell.value})")
                    continue
                
                # Skip if no prompt
                if not prompt:
                    print(f"Row {row}: No prompt, skipping")
                    continue
                
                print(f"\nRow {row}: Processing...")
                print(f"  Dell Product: {dell_product}")
                print(f"  SNOW Product: {snow_product}")
                print(f"  Dell Publisher: {dell_publisher}")
                print(f"  SNOW Publisher: {snow_publisher}")
                
                # Simple matching logic based on your requirements: 
                # Match = Both product AND publisher are same
                # Not Match = If either product OR publisher differ
                
                product_same = (str(dell_product).lower().strip() == str(snow_product).lower().strip())
                publisher_same = (str(dell_publisher).lower().strip() == str(snow_publisher).lower().strip())
                
                print(f"  Product same: {product_same}")
                print(f"  Publisher same: {publisher_same}")
                
                # Apply matching rule
                if product_same and publisher_same: 
                    status_cell.value = "Match"
                    print(f"  Status: Match ✓")
                else:
                    status_cell.value = "Not Match"
                    print(f"  Status: Not Match ✓")
                
                time.sleep(0.5)
                
            except Exception as e: 
                print(f"Row {row}: Error - {str(e)}")
                ws.cell(row=row, column=11).value = "Error"
                continue
        
        # Save the updated file
        wb.save(EXCEL_FILE)
        print(f"\n✅ SUCCESS! File updated: {EXCEL_FILE}")
        print(f"Total rows processed: {ws.max_row - 1}")
        
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        raise

if __name__ == "__main__": 
    process_excel()
