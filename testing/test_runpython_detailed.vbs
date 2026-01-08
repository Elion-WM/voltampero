' VBScript to test RunPython outside of Excel
' This helps isolate the issue

Set xl = CreateObject("Excel.Application")
xl.Visible = True

Set wb = xl.Workbooks.Open("C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")

WScript.Echo "Workbook opened"

On Error Resume Next
xl.Run "VoltAmpero.SimpleTest"

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "Macro ran successfully"
End If

WScript.Echo "Press OK to close"
wb.Close False
xl.Quit
