
import xlwings as xw
import os

def verify_excel_structure():
    print("Verifying VoltAmpero.xlsm structure...")
    
    try:
        # Connect to active workbook or open it
        try:
            wb = xw.Book('VoltAmpero.xlsm')
        except Exception:
            # try finding it in current dir
            cwd = r"C:\Users\User\GitHub\voltampero"
            path = os.path.join(cwd, 'VoltAmpero.xlsm')
            if os.path.exists(path):
                wb = xw.Book(path)
            else:
                print(f"ERROR: VoltAmpero.xlsm not found at {path}")
                return

        print(f"Connected to: {wb.name}")
        
        # Check Sheets
        sheet_names = [s.name for s in wb.sheets]
        print(f"Sheets found: {sheet_names}")
        
        required_sheets = ["Control", "Data"]
        for req in required_sheets:
            if req in sheet_names:
                print(f"  [OK] Sheet '{req}' exists")
            else:
                print(f"  [FAIL] Sheet '{req}' MISSING")
                
        # Check Named Ranges
        print("\nChecking Named Ranges...")
        names = [n.name for n in wb.names]
        
        required_names = [
            "PSUPort", "PSUStatus", "DMMStatus", "LoggingStatus", "RampStatus",
            "LogInterval", "LiveVoltage", "LiveCurrent", "LiveDMM",
            "SetVoltage", "SetCurrent", "OCPEnabled",
            "RampStartV", "RampEndV", "RampDuration", "RampCycles", 
            "RampDelay", "RampPingPong", "RampCycle", "RampVoltage", "RampProgress",
            "ExportStatus"
        ]
        
        for name in required_names:
            found = False
            # Names can be workbook scope or sheet scope (Sheet1!Name)
            for n in names:
                if name == n or n.endswith("!" + name):
                    found = True
                    # Check if it refers to a valid range
                    try:
                        ref = wb.names[n].refers_to_range
                        print(f"  [OK] Name '{name}' -> {ref.address}")
                    except:
                        print(f"  [WARN] Name '{name}' exists but has broken reference")
                    break
            
            if not found:
                 print(f"  [FAIL] Name '{name}' MISSING")
                 
        # Check Buttons (Shapes)
        print("\nChecking Buttons (Shapes) on Control sheet...")
        ws = wb.sheets['Control']
        shapes = [s.name for s in ws.shapes]
        # Excel buttons usually have auto-generated names like "Button 1", so we check text if possible
        # but xlwings shape text access can be tricky. We'll list count.
        print(f"  Found {len(shapes)} shapes/buttons.")
        
        # Try to guess if enough buttons exist based on standard layout
        if len(shapes) >= 14:
             print("  [OK] Button count seems reasonable (expecting ~15)")
        else:
             print("  [WARN] Low button count. Might need 'SetupWorkbook' macro run.")

    except Exception as e:
        print(f"Verification failed: {e}")

if __name__ == "__main__":
    verify_excel_structure()
