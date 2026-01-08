"""
Diagnose what happens when Apply Settings is clicked
"""

import xlwings as xw

def diagnose_live():
    print("Diagnosing Apply Settings in live Excel...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control = wb.sheets["Control"]
    
    # Check what's in the cells
    print("\n1. Current cell values:")
    print(f"   B16 (SetVoltage input): {control.range('B16').value}")
    print(f"   B17 (SetCurrent input): {control.range('B17').value}")
    print(f"   E12 (Setpoint display): {control.range('E12').value}")
    print(f"   E13 (Setpoint display): {control.range('E13').value}")
    print(f"   D3 (PSU Status): {control.range('D3').value}")
    
    # Check the controller state
    print("\n2. Checking Python controller state:")
    import sys
    sys.path.insert(0, r"C:\Users\User\GitHub\voltampero")
    
    try:
        from voltampero import get_controller, _controller
        
        print(f"   Global controller exists: {_controller is not None}")
        
        if _controller:
            ctrl = _controller
            print(f"   PSU type: {type(ctrl.psu).__name__}")
            print(f"   PSU connected: {ctrl.psu.is_connected()}")
            
            if ctrl.psu.is_connected():
                print(f"   PSU voltage setpoint: {ctrl.psu.get_voltage_setpoint()}V")
                print(f"   PSU current setpoint: {ctrl.psu.get_current_setpoint()}A")
                psu_v, psu_a = ctrl.psu.get_readings()
                print(f"   PSU output: {psu_v}V, {psu_a}A")
                
                # Check if output is on
                try:
                    status = ctrl.psu.get_status()
                    print(f"   Output enabled: {status.output_on if hasattr(status, 'output_on') else 'Unknown'}")
                except:
                    print(f"   Could not get output status")
            else:
                print("   [WARNING] PSU not connected!")
                print("   Did you click 'Test (Simulated)' first?")
        else:
            print("   [WARNING] No controller initialized!")
            print("   Click 'Test (Simulated)' to initialize")
            
    except Exception as e:
        print(f"   [ERROR] {e}")
        import traceback
        traceback.print_exc()
    
    # Check if the VBA code is correct
    print("\n3. Checking VBA ApplySettings code:")
    try:
        vb_project = wb.api.VBProject
        for component in vb_project.VBComponents:
            if component.Name == "VoltAmpero":
                code_module = component.CodeModule
                
                # Find ApplySettings
                for line_num in range(1, code_module.CountOfLines + 1):
                    line = code_module.Lines(line_num, 1)
                    if "Sub ApplySettings()" in line:
                        print(f"   Found ApplySettings at line {line_num}")
                        
                        # Show the RunPython line
                        for i in range(line_num, min(line_num + 30, code_module.CountOfLines + 1)):
                            line_text = code_module.Lines(i, 1)
                            if "RunPython" in line_text:
                                print(f"   RunPython command: {line_text.strip()}")
                        break
                break
    except Exception as e:
        print(f"   Could not check VBA: {e}")
    
    print("\n" + "="*60)
    print("INSTRUCTIONS TO TEST:")
    print("="*60)
    print("1. Make sure PSU Status (D3) shows 'Connected'")
    print("   - If not, click 'Test (Simulated)' first")
    print("2. Set B16 = 10.0 and B17 = 2.0")
    print("3. Click 'Apply Settings'")
    print("4. Run this script again to see if values changed")
    print("="*60)

if __name__ == "__main__":
    diagnose_live()
