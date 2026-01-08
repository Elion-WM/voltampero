"""
Add cells to display the PSU setpoints separately from live readings
"""

import xlwings as xw

def add_setpoint_display():
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control = wb.sheets["Control"]
    
    print("Adding PSU setpoint display...")
    
    # Add new section for PSU setpoints
    control.range("D11").value = "PSU Setpoints:"
    control.range("D11").api.Font.Bold = True
    
    control.range("D12").value = "Set Voltage:"
    control.range("E12").value = 0.0
    
    control.range("D13").value = "Set Current:"
    control.range("E13").value = 0.0
    
    # Add named ranges for these new cells
    wb.names.add("PSUSetVoltageDisplay", "=Control!$E$12")
    wb.names.add("PSUSetCurrentDisplay", "=Control!$E$13")
    
    print("   [OK] Added setpoint display cells")
    
    # Now improve ApplySettings to update these cells
    print("\nUpdating ApplySettings to show setpoints...")
    vb_project = wb.api.VBProject
    
    for component in vb_project.VBComponents:
        if component.Name == "VoltAmpero":
            code_module = component.CodeModule
            
            # Find ApplySettings
            for line_num in range(1, code_module.CountOfLines + 1):
                line = code_module.Lines(line_num, 1)
                if "Sub ApplySettings()" in line:
                    # Find the End Sub
                    end_line = line_num
                    for i in range(line_num + 1, code_module.CountOfLines + 1):
                        if "End Sub" in code_module.Lines(i, 1):
                            end_line = i
                            break
                    
                    # Delete and replace
                    num_lines = end_line - line_num + 1
                    code_module.DeleteLines(line_num, num_lines)
                    
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
    Range("PSUStatus").Value = "Applying..."
    DoEvents
    
    ' Apply settings via Python
    RunPython "from voltampero import get_controller; c=get_controller(); c.set_voltage(" & Replace(voltage, ",", ".") & "); c.set_current(" & Replace(current, ",", ".") & "); c.set_ocp(" & IIf(ocp, "True", "False") & ")"
    
    ' Update setpoint display
    Range("PSUSetVoltageDisplay").Value = voltage
    Range("PSUSetCurrentDisplay").Value = current
    
    ' Update status
    Range("PSUStatus").Value = "Connected"
    
    ' Show confirmation
    MsgBox "PSU setpoints updated to:" & vbCrLf & _
           "Voltage: " & voltage & " V" & vbCrLf & _
           "Current: " & current & " A" & vbCrLf & vbCrLf & _
           "Click 'Output ON' to enable the output", vbInformation, "Settings Applied"
End Sub
'''
                    code_module.InsertLines(line_num, new_code)
                    print("   [OK] Updated ApplySettings")
                    break
    
    wb.save()
    print("\n[OK] All improvements saved!")
    print("\nNow you will see:")
    print("- Column B11-B13: LIVE readings (actual output)")
    print("- Column E12-E13: PSU Setpoints (target values)")
    print("\nWorkflow:")
    print("1. Set desired voltage/current in B16, B17")
    print("2. Click 'Apply Settings'")
    print("3. See setpoints update in E12, E13")
    print("4. Click 'Output ON' to enable")
    print("5. See live output in B11, B12")

if __name__ == "__main__":
    add_setpoint_display()
