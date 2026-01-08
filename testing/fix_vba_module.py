"""
Fix VBA module - import VoltAmpero.bas into the workbook
"""

import xlwings as xw
import os

def fix_vba_module():
    print("Fixing VBA module...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    vb_project = wb.api.VBProject
    
    # Remove Module1 if it exists
    print("\n1. Removing old modules...")
    for component in vb_project.VBComponents:
        if component.Name in ["Module1", "VoltAmpero"]:
            print(f"   Removing: {component.Name}")
            vb_project.VBComponents.Remove(component)
    
    # Import VoltAmpero.bas
    print("\n2. Importing VoltAmpero.bas...")
    bas_file = r"C:\Users\User\GitHub\voltampero\VoltAmpero.bas"
    
    if not os.path.exists(bas_file):
        print(f"   [ERROR] File not found: {bas_file}")
        return False
    
    try:
        vb_project.VBComponents.Import(bas_file)
        print(f"   [OK] Imported VoltAmpero.bas")
    except Exception as e:
        print(f"   [ERROR] Failed to import: {e}")
        return False
    
    # List all modules now
    print("\n3. Current VBA modules:")
    for component in vb_project.VBComponents:
        print(f"   - {component.Name} (Type: {component.Type})")
    
    # Save
    print("\n4. Saving...")
    wb.save()
    print("   [OK] Saved")
    
    print("\n" + "="*60)
    print("SUCCESS! VBA module has been fixed.")
    print("Close and reopen Excel, then try the buttons again.")
    print("="*60)
    
    return True

if __name__ == "__main__":
    fix_vba_module()
