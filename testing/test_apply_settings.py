"""
Test the ApplySettings functionality
"""

import xlwings as xw

def test_apply_settings():
    print("Testing Apply Settings functionality...")
    
    # Open workbook
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control_sheet = wb.sheets["Control"]
    
    # Read current values from Excel
    print("\n1. Reading settings from Excel:")
    voltage = control_sheet.range("B16").value
    current = control_sheet.range("B17").value
    ocp = control_sheet.range("B18").value
    
    print(f"   Voltage: {voltage}")
    print(f"   Current: {current}")
    print(f"   OCP: {ocp}")
    
    # Now test the Python code
    print("\n2. Testing Python voltampero module:")
    import sys
    sys.path.insert(0, r"C:\Users\User\GitHub\voltampero")
    
    try:
        from voltampero import get_controller
        
        # Get controller (should be simulated if InitSimulated was run)
        ctrl = get_controller(simulate=True)
        print(f"   Controller created: {ctrl}")
        print(f"   PSU type: {type(ctrl.psu).__name__}")
        print(f"   PSU connected: {ctrl.psu.is_connected()}")
        
        # Try to connect
        if not ctrl.psu.is_connected():
            print("\n3. Connecting simulated PSU...")
            ctrl.connect_psu("SIM1")
            print(f"   Connected: {ctrl.psu.is_connected()}")
        
        # Get current PSU settings before change
        print("\n4. Current PSU settings:")
        psu_v, psu_a = ctrl.psu.get_readings()
        print(f"   Output: {psu_v}V, {psu_a}A")
        set_v = ctrl.psu.get_voltage_setpoint()
        set_a = ctrl.psu.get_current_setpoint()
        print(f"   Setpoints: {set_v}V, {set_a}A")
        
        # Apply new settings
        print("\n5. Applying new settings from Excel...")
        print(f"   Setting voltage to {voltage}V")
        result_v = ctrl.set_voltage(voltage)
        print(f"   Result: {result_v}")
        
        print(f"   Setting current to {current}A")
        result_i = ctrl.set_current(current)
        print(f"   Result: {result_i}")
        
        # Read back
        print("\n6. Reading back PSU settings:")
        set_v = ctrl.psu.get_voltage_setpoint()
        set_a = ctrl.psu.get_current_setpoint()
        print(f"   Setpoints now: {set_v}V, {set_a}A")
        
        if set_v == voltage and set_a == current:
            print("\n   [SUCCESS] Settings applied correctly!")
        else:
            print("\n   [WARNING] Settings don't match:")
            print(f"   Expected: {voltage}V, {current}A")
            print(f"   Got: {set_v}V, {set_a}A")
        
        # Check if output is on
        print("\n7. Checking output state:")
        status = ctrl.psu.get_status()
        print(f"   Output enabled: {status.get('output_on', 'Unknown')}")
        
        # Turn output on to see the actual voltage/current
        print("\n8. Turning output ON...")
        ctrl.output_on()
        psu_v, psu_a = ctrl.psu.get_readings()
        print(f"   Output now: {psu_v}V, {psu_a}A")
        
    except Exception as e:
        print(f"\n   [ERROR] {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_apply_settings()
