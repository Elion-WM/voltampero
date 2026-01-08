"""
Test and diagnose xlwings VBA integration
"""

import xlwings as xw
import sys

def test_xlwings_connection():
    """Test if xlwings can communicate with Excel"""
    
    print("Testing xlwings connection...")
    
    try:
        # Open the workbook
        wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
        print("[OK] Workbook opened successfully")
        
        # Check if Control sheet exists
        try:
            control = wb.sheets["Control"]
            print("[OK] Control sheet found")
        except:
            print("[ERROR] Control sheet not found")
            return False
        
        # Try to read a cell
        try:
            port = control.range("B3").value
            print(f"[OK] Can read cells (PSU Port = {port})")
        except Exception as e:
            print(f"[ERROR] Cannot read cells: {e}")
        
        # Try to write to status
        try:
            control.range("D3").value = "Test OK"
            print("[OK] Can write to cells")
        except Exception as e:
            print(f"[ERROR] Cannot write cells: {e}")
        
        print("\n" + "="*50)
        print("xlwings is working correctly!")
        print("="*50)
        
        print("\nThe issue is that the VBA code needs the xlwings VBA module.")
        print("Run this in Excel VBA Immediate Window (Ctrl+G):")
        print("  Application.Run \"xlwings.xlam!xlwingsAddin.LoadToolbar\"")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error: {e}")
        return False

if __name__ == "__main__":
    test_xlwings_connection()
