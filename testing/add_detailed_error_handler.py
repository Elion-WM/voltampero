"""
Add detailed error handling to VBA to see exactly what's failing
"""

import xlwings as xw

def add_detailed_test():
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    vb_project = wb.api.VBProject
    
    # Find VoltAmpero module
    for component in vb_project.VBComponents:
        if component.Name == "VoltAmpero":
            code_module = component.CodeModule
            
            # Add a detailed test macro
            test_code = '''

Sub DetailedTest()
    On Error GoTo ErrorHandler
    
    Dim result As String
    
    ' Test 1: Check if RunPython exists
    Range("D3").Value = "Test 1: Checking RunPython..."
    DoEvents
    
    ' Test 2: Try simplest possible Python command
    Range("D3").Value = "Test 2: Running Python..."
    DoEvents
    
    RunPython "import sys; print('Python OK')"
    
    Range("D3").Value = "Test 3: SUCCESS!"
    MsgBox "All tests passed!", vbInformation
    Exit Sub
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf
    errMsg = errMsg & "Error Source: " & Err.Source & vbCrLf
    errMsg = errMsg & "Help Context: " & Err.HelpContext & vbCrLf
    errMsg = errMsg & vbCrLf & "This error occurred in DetailedTest macro."
    
    Range("D3").Value = "ERROR: " & Err.Number
    MsgBox errMsg, vbCritical, "Detailed Error Information"
End Sub
'''
            code_module.AddFromString(test_code)
            print("[OK] Added DetailedTest macro")
            break
    
    wb.save()
    print("[OK] Saved")
    print("\nIn Excel:")
    print("1. Press Alt+F8")
    print("2. Run 'DetailedTest'")
    print("3. Copy the COMPLETE error message and send it to me")

if __name__ == "__main__":
    add_detailed_test()
