"""
Debug script to test Apply Settings with REAL PSU
Run this with Excel open and PSU connected
"""

import xlwings as xw
import sys
import time

print("=" * 60)
print("Real PSU Apply Settings Debug Test")
print("=" * 60)

# Get Excel workbook
print("\n1. Getting Excel workbook...")
try:
    books = xw.books
    wb = None
    for book in books:
        if "VoltAmpero" in book.name:
            wb = book
            print(f"   [OK] Found: {book.name}")
            break
    if not wb:
        print("   [FAIL] VoltAmpero.xlsm not found. Please open it.")
        sys.exit(1)
    ctrl = wb.sheets["Control"]
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Check current status in Excel
print("\n2. Checking Excel status...")
try:
    psu_port = ctrl.range("PSUPort").value
    psu_status = ctrl.range("PSUStatus").value
    set_voltage = ctrl.range("SetVoltage").value
    set_current = ctrl.range("SetCurrent").value
    
    print(f"   PSU Port: {psu_port}")
    print(f"   PSU Status: {psu_status}")
    print(f"   Set Voltage: {set_voltage}V")
    print(f"   Set Current: {set_current}A")
    
    if not psu_port or "SIM" in str(psu_port).upper():
        print(f"\n   [WARN] Excel is in simulated mode!")
        print(f"   Please enter your real COM port (e.g., COM3) in the PSUPort cell")
        print(f"   Then click 'Connect PSU' button in Excel")
        sys.exit(0)
        
    if psu_status != "Connected":
        print(f"\n   [WARN] PSU is not connected!")
        print(f"   Please click 'Connect PSU' button in Excel first")
        sys.exit(0)
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Import voltampero
print("\n3. Importing voltampero module...")
try:
    from voltampero import get_controller
    print(f"   [OK] Module imported")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Get the controller instance
print("\n4. Getting controller instance...")
try:
    ctrl_obj = get_controller()
    print(f"   [OK] Controller instance obtained")
    print(f"   PSU type: {type(ctrl_obj.psu).__name__}")
    print(f"   PSU connected: {ctrl_obj.psu.is_connected()}")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Attach to Excel
print("\n5. Attaching to Excel...")
try:
    result = ctrl_obj.attach_excel(wb)
    print(f"   attach_excel() returned: {result}")
    print(f"   control_sheet exists: {ctrl_obj.control_sheet is not None}")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Check PSU connection status
print("\n6. Checking PSU connection in detail...")
try:
    is_connected = ctrl_obj.psu.is_connected()
    print(f"   is_connected(): {is_connected}")
    
    if is_connected:
        print(f"   Serial port: {ctrl_obj.psu.port}")
        print(f"   Serial object: {ctrl_obj.psu.serial}")
        if ctrl_obj.psu.serial:
            print(f"   Serial is_open: {ctrl_obj.psu.serial.is_open}")
    else:
        print(f"   [WARN] PSU reports as NOT connected!")
        print(f"   This is the problem - Excel thinks it's connected but Python doesn't")
        print(f"   Attempting to reconnect...")
        
        # Try to reconnect
        reconnect_result = ctrl_obj.connect_psu(psu_port)
        print(f"   Reconnect result: {reconnect_result}")
        if reconnect_result:
            print(f"   [OK] Reconnected successfully!")
        else:
            print(f"   [FAIL] Reconnection failed")
            
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Try to read current setpoints from PSU
print("\n7. Reading current setpoints from PSU...")
try:
    current_v = ctrl_obj.psu.get_voltage_setpoint()
    current_a = ctrl_obj.psu.get_current_setpoint()
    print(f"   Current voltage setpoint: {current_v}V")
    print(f"   Current current setpoint: {current_a}A")
except Exception as e:
    print(f"   [FAIL] Error reading from PSU: {e}")
    print(f"   This suggests communication problem with PSU")
    import traceback
    traceback.print_exc()

# Try to set new values
print("\n8. Attempting to apply settings from Excel...")
try:
    target_v = float(set_voltage)
    target_a = float(set_current)
    
    print(f"   Target: {target_v}V, {target_a}A")
    print(f"   PSU connected before set_voltage: {ctrl_obj.psu.is_connected()}")
    
    # Call set_voltage
    print(f"   Calling set_voltage({target_v})...")
    result_v = ctrl_obj.set_voltage(target_v)
    print(f"   set_voltage returned: {result_v}")
    
    print(f"   PSU connected after set_voltage: {ctrl_obj.psu.is_connected()}")
    
    # Call set_current
    print(f"   Calling set_current({target_a})...")
    result_a = ctrl_obj.set_current(target_a)
    print(f"   set_current returned: {result_a}")
    
    # Wait a bit for PSU to process
    time.sleep(0.2)
    
    # Verify what got set
    print(f"\n   Verifying settings on PSU...")
    actual_v = ctrl_obj.psu.get_voltage_setpoint()
    actual_a = ctrl_obj.psu.get_current_setpoint()
    
    print(f"   PSU reports: {actual_v}V, {actual_a}A")
    print(f"   Expected: {target_v}V, {target_a}A")
    
    if abs(actual_v - target_v) < 0.1 and abs(actual_a - target_a) < 0.1:
        print(f"\n   [OK] SUCCESS! Settings were applied to PSU!")
    else:
        print(f"\n   [FAIL] Settings mismatch!")
        print(f"   Voltage error: {abs(actual_v - target_v):.3f}V")
        print(f"   Current error: {abs(actual_a - target_a):.3f}A")
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Try sending a raw command to PSU
print("\n9. Testing raw PSU communication...")
try:
    # Try to get ID
    idn = ctrl_obj.psu.get_identification()
    print(f"   PSU Identification: {idn}")
    
    # Try to read output voltage
    vout = ctrl_obj.psu.get_output_voltage()
    iout = ctrl_obj.psu.get_output_current()
    print(f"   Output readings: {vout}V, {iout}A")
    
    if idn and idn != "Unknown":
        print(f"   [OK] PSU communication is working!")
    else:
        print(f"   [WARN] PSU communication may be faulty")
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Diagnostic complete!")
print("=" * 60)
print("\nNEXT STEPS:")
print("1. Check if settings were applied above")
print("2. Check the physical PSU display - does it show the new values?")
print("3. If not, there may be a communication protocol issue")
