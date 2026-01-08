Option Explicit

' Test if VBA can call Python at all
Dim objExcel, objWorkbook

' Open Excel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

' Open workbook
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")

' Wait a moment
WScript.Sleep 2000

' Try calling the Python function
On Error Resume Next
objExcel.Run "OutputOn"

If Err.Number <> 0 Then
    WScript.Echo "ERROR calling OutputOn: " & Err.Description
    WScript.Echo "Error number: " & Err.Number
Else
    WScript.Echo "OutputOn macro executed (no VBA error)"
End If

' Check if debug log was created
WScript.Sleep 1000

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("C:\Users\User\GitHub\voltampero\debug_log.txt") Then
    WScript.Echo "SUCCESS: debug_log.txt was created!"
    
    Dim file, contents
    Set file = fso.OpenTextFile("C:\Users\User\GitHub\voltampero\debug_log.txt", 1)
    contents = file.ReadAll
    file.Close
    WScript.Echo "Log contents:"
    WScript.Echo contents
Else
    WScript.Echo "FAIL: debug_log.txt was NOT created"
    WScript.Echo "This means Python function was not called"
End If

WScript.Echo ""
WScript.Echo "Press OK to close Excel"

' Clean up
objWorkbook.Close False
objExcel.Quit
