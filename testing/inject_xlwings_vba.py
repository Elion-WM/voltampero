"""
Inject xlwings VBA code into VoltAmpero workbook
This adds the RunPython function and xlwings infrastructure
"""

import xlwings as xw
import os

def inject_xlwings_code():
    """Add xlwings VBA module to the workbook"""
    
    print("Injecting xlwings VBA code into VoltAmpero.xlsm...")
    
    # Get xlwings VBA code directory
    xlwings_dir = os.path.dirname(xw.__file__)
    vba_file = os.path.join(xlwings_dir, "xlwings.bas")
    
    if not os.path.exists(vba_file):
        print(f"[ERROR] xlwings.bas not found at: {vba_file}")
        print("\nAlternative: Use xlwings quickstart command")
        return False
    
    # Open workbook
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    print("[OK] Workbook opened")
    
    # Check if xlwings module already exists
    try:
        # Try to access the VBA project
        vb_project = wb.api.VBProject
        
        # Check if xlwings module exists
        has_xlwings = False
        for component in vb_project.VBComponents:
            if component.Name.lower() == "xlwings":
                print(f"[INFO] Found existing xlwings module: {component.Name}")
                print("      Removing old module...")
                vb_project.VBComponents.Remove(component)
                break
        
        # Import the xlwings.bas file
        print(f"[INFO] Importing xlwings VBA code from: {vba_file}")
        vb_project.VBComponents.Import(vba_file)
        print("[OK] xlwings module injected successfully!")
        
        # Save workbook
        wb.save()
        print("[OK] Workbook saved")
        
        print("\n" + "="*60)
        print("SUCCESS! xlwings VBA module has been added.")
        print("="*60)
        print("\nNow try clicking your buttons in Excel again.")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Failed to inject VBA code: {e}")
        print("\nThis might be due to macro security settings.")
        print("\nTo fix:")
        print("1. Open Excel")
        print("2. File > Options > Trust Center > Trust Center Settings")
        print("3. Macro Settings > Trust access to the VBA project object model")
        print("4. Run this script again")
        return False

if __name__ == "__main__":
    inject_xlwings_code()
