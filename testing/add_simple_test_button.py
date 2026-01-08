"""
Add a very simple test button that just writes to a cell
"""

import xlwings as xw

def add_simple_test():
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    vb_project = wb.api.VBProject
    
    # Add simple test macro to VoltAmpero module
    print("Adding SimpleTest macro...")
    
    for component in vb_project.VBComponents:
        if component.Name == "VoltAmpero":
            code_module = component.CodeModule
            
            # Add a very simple test at the end
            test_code = '''

Sub SimpleTest()
    On Error GoTo ErrorHandler
    Range("D3").Value = "Testing..."
    RunPython "from simple_test import simple_test; simple_test()"
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical
End Sub
'''
            code_module.AddFromString(test_code)
            print("[OK] Added SimpleTest macro")
            break
    
    wb.save()
    print("[OK] Saved")
    print("\nNow in Excel:")
    print("1. Press Alt+F8")
    print("2. Run 'SimpleTest' macro")
    print("3. Tell me the exact error message if it fails")

if __name__ == "__main__":
    add_simple_test()
