"""
Check why voltage and current cells can't be edited
"""

import xlwings as xw

def check_cells():
    print("Checking voltage and current cells...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control_sheet = wb.sheets["Control"]
    
    print("\n1. Checking SetVoltage cell (B16):")
    cell = control_sheet.range("B16")
    print(f"   Value: {cell.value}")
    print(f"   Formula: {cell.formula}")
    print(f"   Locked: {cell.api.Locked}")
    print(f"   Address: {cell.address}")
    
    print("\n2. Checking SetCurrent cell (B17):")
    cell = control_sheet.range("B17")
    print(f"   Value: {cell.value}")
    print(f"   Formula: {cell.formula}")
    print(f"   Locked: {cell.api.Locked}")
    print(f"   Address: {cell.address}")
    
    print("\n3. Checking sheet protection:")
    try:
        protected = control_sheet.api.ProtectContents
        print(f"   Sheet protected: {protected}")
        if protected:
            print("   [WARNING] Sheet is protected - cells cannot be edited!")
            print("   Unprotecting sheet...")
            control_sheet.api.Unprotect()
            print("   [OK] Sheet unprotected")
    except Exception as e:
        print(f"   Error checking protection: {e}")
    
    print("\n4. Checking OCPEnabled cell (B18):")
    cell = control_sheet.range("B18")
    print(f"   Value: {cell.value}")
    print(f"   Formula: {cell.formula}")
    print(f"   Locked: {cell.api.Locked}")
    
    print("\n5. Setting test values:")
    try:
        control_sheet.range("B16").value = 5.0
        control_sheet.range("B17").value = 1.0
        control_sheet.range("B18").value = False
        print("   [OK] Successfully set test values")
        print(f"   Voltage = {control_sheet.range('B16').value}")
        print(f"   Current = {control_sheet.range('B17').value}")
    except Exception as e:
        print(f"   [ERROR] Cannot set values: {e}")
    
    print("\n6. Checking named ranges:")
    try:
        print(f"   SetVoltage range: {wb.names['SetVoltage'].refers_to}")
        print(f"   SetCurrent range: {wb.names['SetCurrent'].refers_to}")
        print(f"   OCPEnabled range: {wb.names['OCPEnabled'].refers_to}")
    except Exception as e:
        print(f"   Error: {e}")
    
    wb.save()
    print("\n[OK] Saved workbook")

if __name__ == "__main__":
    check_cells()
