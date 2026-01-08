"""
Fix Excel issues:
1. Move Data headers to row 1
2. Diagnose the run-time error
"""

import xlwings as xw

def fix_excel_issues():
    print("Opening VoltAmpero.xlsm...")
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    
    # Fix Data sheet headers
    print("\n1. Fixing Data sheet headers...")
    try:
        data_sheet = wb.sheets["Data"]
        
        # Check current header location
        for row in range(1, 10):
            val = data_sheet.range(f"A{row}").value
            if val and "Timestamp" in str(val):
                print(f"   Found headers in row {row}")
                if row != 1:
                    # Copy headers from current row to row 1
                    header_range = data_sheet.range(f"A{row}:I{row}")
                    headers = header_range.value
                    print(f"   Moving headers to row 1: {headers}")
                    
                    # Clear everything first
                    last_row = data_sheet.range("A1").end('down').row
                    if last_row > 1:
                        data_sheet.range(f"A1:I{last_row}").clear_contents()
                    
                    # Write correct headers to row 1
                    data_sheet.range("A1").value = [
                        "Timestamp", "Elapsed_s", "PSU_Voltage_V", "PSU_Current_A",
                        "PSU_Setpoint_V", "PSU_Setpoint_A", "DMM_Value", "DMM_Unit", "DMM_Mode"
                    ]
                    print("   [OK] Headers moved to row 1")
                else:
                    print("   [OK] Headers already in row 1")
                break
        else:
            # No headers found, create them
            print("   No headers found, creating them in row 1...")
            data_sheet.range("A1").value = [
                "Timestamp", "Elapsed_s", "PSU_Voltage_V", "PSU_Current_A",
                "PSU_Setpoint_V", "PSU_Setpoint_A", "DMM_Value", "DMM_Unit", "DMM_Mode"
            ]
            print("   [OK] Headers created")
            
        # Make headers bold
        data_sheet.range("A1:I1").api.Font.Bold = True
        
    except Exception as e:
        print(f"   [ERROR] {e}")
    
    # Check VBA modules
    print("\n2. Checking VBA modules...")
    try:
        vb_project = wb.api.VBProject
        modules = []
        for component in vb_project.VBComponents:
            modules.append(component.Name)
            print(f"   Found module: {component.Name} (Type: {component.Type})")
        
        if "xlwings" not in [m.lower() for m in modules]:
            print("   [WARNING] xlwings module not found!")
        else:
            print("   [OK] xlwings module exists")
            
        if "VoltAmpero" not in modules:
            print("   [WARNING] VoltAmpero module not found!")
        else:
            print("   [OK] VoltAmpero module exists")
            
    except Exception as e:
        print(f"   [ERROR] Cannot access VBA project: {e}")
        print("   Make sure 'Trust access to VBA project' is enabled in Excel")
    
    # Save
    print("\n3. Saving workbook...")
    wb.save()
    print("   [OK] Saved")
    
    print("\n" + "="*60)
    print("To diagnose the error further, please:")
    print("1. Close and reopen Excel")
    print("2. Click a button and note the EXACT error message")
    print("3. If it says 'Run-time error 62', click Debug")
    print("4. Tell me which line of VBA code is highlighted")
    print("="*60)

if __name__ == "__main__":
    fix_excel_issues()
