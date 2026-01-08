"""
Deep diagnosis of RunPython issue
"""

import xlwings as xw
import os

def diagnose_runpython():
    print("Deep diagnosis of xlwings RunPython...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    
    print("\n1. Checking xlwings configuration...")
    config_file = r"C:\Users\User\GitHub\voltampero\xlwings.conf"
    if os.path.exists(config_file):
        print("   [OK] xlwings.conf exists")
        with open(config_file, 'r') as f:
            print("   Content:")
            for line in f:
                print(f"      {line.rstrip()}")
    else:
        print("   [ERROR] xlwings.conf not found!")
    
    print("\n2. Checking VBA modules...")
    try:
        vb_project = wb.api.VBProject
        
        xlwings_found = False
        voltampero_found = False
        
        for component in vb_project.VBComponents:
            print(f"   Module: {component.Name}")
            
            if component.Name.lower() == "xlwings":
                xlwings_found = True
                # Check if RunPython exists in the code
                code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                if "Sub RunPython" in code or "Function RunPython" in code:
                    print("      [OK] RunPython function found")
                else:
                    print("      [ERROR] RunPython function NOT found in xlwings module!")
                    
            if component.Name == "VoltAmpero":
                voltampero_found = True
        
        if not xlwings_found:
            print("   [ERROR] xlwings module not found!")
        if not voltampero_found:
            print("   [ERROR] VoltAmpero module not found!")
            
    except Exception as e:
        print(f"   [ERROR] Cannot access VBA: {e}")
    
    print("\n3. Testing simple Python execution from Excel...")
    try:
        # Create a simple test macro
        test_code = '''
Sub TestRunPython()
    On Error Resume Next
    RunPython "print('Hello from Python')"
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    Else
        MsgBox "RunPython works!"
    End If
End Sub
'''
        # Try to add a test module
        try:
            test_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            test_module.Name = "TestModule"
            test_module.CodeModule.AddFromString(test_code)
            print("   [OK] Added TestModule with TestRunPython macro")
            print("   Try running 'TestRunPython' macro in Excel (Alt+F8)")
        except Exception as e:
            print(f"   [ERROR] Could not add test module: {e}")
            
    except Exception as e:
        print(f"   [ERROR] {e}")
    
    print("\n4. Checking Python interpreter...")
    interpreter = r"C:\Users\User\GitHub\voltampero\python\python.exe"
    if os.path.exists(interpreter):
        print(f"   [OK] Python found at: {interpreter}")
    else:
        print(f"   [ERROR] Python NOT found at: {interpreter}")
    
    print("\n5. Saving workbook...")
    wb.save()
    print("   [OK] Saved")
    
    print("\n" + "="*60)
    print("NEXT STEPS:")
    print("1. Close and reopen Excel")
    print("2. Press Alt+F8")
    print("3. Run 'TestRunPython' macro")
    print("4. Tell me what error message you get (if any)")
    print("="*60)

if __name__ == "__main__":
    diagnose_runpython()
