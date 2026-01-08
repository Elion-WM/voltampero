"""
Improve Apply Settings to show feedback
"""

import xlwings as xw

def improve_vba():
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    vb_project = wb.api.VBProject
    
    # Find VoltAmpero module
    for component in vb_project.VBComponents:
        if component.Name == "VoltAmpero":
            code_module = component.CodeModule
            
            # Find and replace ApplySettings
            print("Improving ApplySettings macro...")
            
            # Search for the ApplySettings sub
            for line_num in range(1, code_module.CountOfLines + 1):
                line = code_module.Lines(line_num, 1)
                if "Sub ApplySettings()" in line:
                    # Found it, now replace the whole sub
                    print(f"   Found ApplySettings at line {line_num}")
                    
                    # Find the End Sub
                    end_line = line_num
                    for i in range(line_num + 1, code_module.CountOfLines + 1):
                        if "End Sub" in code_module.Lines(i, 1):
                            end_line = i
                            break
                    
                    # Delete old code
                    num_lines = end_line - line_num + 1
                    code_module.DeleteLines(line_num, num_lines)
                    
                    # Insert new improved version
                    new_code = '''Sub ApplySettings()
    Dim voltage As Variant, current As Variant, ocp As Variant
    
    voltage = Range("SetVoltage").Value
    current = Range("SetCurrent").Value
    ocp = Range("OCPEnabled").Value
    
    ' Validate inputs
    If Not IsNumeric(voltage) Then
        MsgBox "Voltage must be a number", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(current) Then
        MsgBox "Current must be a number", vbExclamation
        Exit Sub
    End If
    
    ' Show what we're doing
    Range("PSUStatus").Value = "Applying settings..."
    DoEvents
    
    ' Apply settings
    RunPython "from voltampero import get_controller; c=get_controller(); c.set_voltage(" & Replace(voltage, ",", ".") & "); c.set_current(" & Replace(current, ",", ".") & "); c.set_ocp(" & IIf(ocp, "True", "False") & ")"
    
    ' Update status
    Range("PSUStatus").Value = "Settings applied"
    
    ' Show confirmation
    MsgBox "PSU settings applied:" & vbCrLf & _
           "Voltage: " & voltage & " V" & vbCrLf & _
           "Current: " & current & " A" & vbCrLf & vbCrLf & _
           "Note: Turn Output ON to see the voltage/current", vbInformation, "Settings Applied"
End Sub
'''
                    code_module.InsertLines(line_num, new_code)
                    print("   [OK] Improved ApplySettings macro")
                    break
    
    # Also add labels to clarify what the cells mean
    print("\nAdding helpful labels to Control sheet...")
    control = wb.sheets["Control"]
    
    # Add note about live readings
    try:
        control.range("C11").value = "(Turn Output ON)"
        control.range("C11").api.Font.Size = 8
        control.range("C11").api.Font.Italic = True
    except:
        pass
    
    wb.save()
    print("\n[OK] Improvements saved!")
    print("\nChanges made:")
    print("1. ApplySettings now shows confirmation message")
    print("2. Reminds user to turn Output ON to see voltage/current")
    print("3. Added label to clarify live readings")

if __name__ == "__main__":
    improve_vba()
