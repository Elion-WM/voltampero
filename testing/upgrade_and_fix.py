"""
Upgrade xlwings and refresh the VBA module
RUN THIS AFTER CLOSING EXCEL!
"""

import subprocess
import sys
import os

def upgrade_and_fix():
    print("="*60)
    print("Upgrading xlwings and refreshing VBA module")
    print("="*60)
    
    python_exe = r"C:\Users\User\GitHub\voltampero\python\python.exe"
    
    # Step 1: Upgrade xlwings
    print("\n1. Upgrading xlwings...")
    try:
        result = subprocess.run(
            [python_exe, "-m", "pip", "install", "--upgrade", "xlwings"],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            print("   [OK] xlwings upgraded successfully")
            print(result.stdout)
        else:
            print(f"   [ERROR] Upgrade failed:")
            print(result.stderr)
            print("\n   Make sure Excel is COMPLETELY closed!")
            return False
    except Exception as e:
        print(f"   [ERROR] {e}")
        return False
    
    # Step 2: Check new version
    print("\n2. Checking new version...")
    try:
        import xlwings as xw
        # Force reload
        import importlib
        importlib.reload(xw)
        print(f"   [OK] xlwings version: {xw.__version__}")
    except:
        print("   [WARNING] Could not verify version")
    
    # Step 3: Re-import xlwings into Excel
    print("\n3. Opening Excel and refreshing xlwings VBA module...")
    try:
        import xlwings as xw
        wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
        vb_project = wb.api.VBProject
        
        # Remove old xlwings module
        for component in vb_project.VBComponents:
            if component.Name.lower() == "xlwings":
                vb_project.VBComponents.Remove(component)
                print("   [OK] Removed old xlwings module")
                break
        
        # Import new xlwings.bas
        xlwings_dir = os.path.dirname(xw.__file__)
        vba_file = os.path.join(xlwings_dir, "xlwings.bas")
        
        if os.path.exists(vba_file):
            vb_project.VBComponents.Import(vba_file)
            print(f"   [OK] Imported new xlwings module")
        else:
            print(f"   [ERROR] xlwings.bas not found")
            return False
        
        # Save and close
        wb.save()
        wb.close()
        print("   [OK] Saved and closed workbook")
        
    except Exception as e:
        print(f"   [ERROR] {e}")
        return False
    
    print("\n" + "="*60)
    print("SUCCESS! xlwings has been upgraded.")
    print("="*60)
    print("\nNow open VoltAmpero.xlsm and try the buttons!")
    
    return True

if __name__ == "__main__":
    upgrade_and_fix()
