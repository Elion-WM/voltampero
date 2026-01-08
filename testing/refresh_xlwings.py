"""
Refresh xlwings module with latest version
"""

import xlwings as xw
import os

def refresh_xlwings():
    print("Refreshing xlwings VBA module...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    vb_project = wb.api.VBProject
    
    # Remove existing xlwings module
    print("\n1. Removing old xlwings module...")
    for component in vb_project.VBComponents:
        if component.Name.lower() == "xlwings":
            vb_project.VBComponents.Remove(component)
            print("   [OK] Removed old xlwings module")
            break
    
    # Import fresh xlwings.bas
    print("\n2. Importing fresh xlwings module...")
    xlwings_dir = os.path.dirname(xw.__file__)
    vba_file = os.path.join(xlwings_dir, "xlwings.bas")
    
    if os.path.exists(vba_file):
        vb_project.VBComponents.Import(vba_file)
        print(f"   [OK] Imported from: {vba_file}")
    else:
        print(f"   [ERROR] xlwings.bas not found at: {vba_file}")
        return False
    
    # Verify it was imported
    print("\n3. Verifying modules...")
    for component in vb_project.VBComponents:
        if component.Name.lower() == "xlwings":
            print("   [OK] xlwings module present")
            # Check for RunPython
            code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
            if "RunPython" in code:
                print("   [OK] RunPython found in module")
            break
    
    # Save
    print("\n4. Saving...")
    wb.save()
    print("   [OK] Saved")
    
    print("\n" + "="*60)
    print("xlwings module refreshed!")
    print("Close Excel and reopen, then try buttons again.")
    print("="*60)
    
    return True

if __name__ == "__main__":
    refresh_xlwings()
