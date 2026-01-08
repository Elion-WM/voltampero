"""
Debug script to test Apply Settings functionality
Run this while Excel is open to diagnose the issue
"""

import xlwings as xw
import sys

print("=" * 60)
print("Apply Settings Debug Test")
print("=" * 60)

# Test 1: Check if Excel is open with our workbook
print("\n1. Checking for open Excel workbook...")
try:
    books = xw.books
    print(f"   Found {len(books)} open workbook(s)")
    for book in books:
        print(f"   - {book.name}")
    
    # Try to find VoltAmpero.xlsm
    wb = None
    for book in books:
        if "VoltAmpero" in book.name:
            wb = book
            print(f"   [OK] Found VoltAmpero workbook: {book.fullname}")
            break
    
    if not wb:
        print("   [FAIL] VoltAmpero.xlsm not found. Please open it.")
        sys.exit(1)
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Test 2: Check if Control sheet exists
print("\n2. Checking Control sheet...")
try:
    ctrl = wb.sheets["Control"]
    print(f"   [OK] Control sheet found")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Test 3: Check Named Ranges
print("\n3. Checking named ranges...")
try:
    voltage = ctrl.range("SetVoltage").value
    current = ctrl.range("SetCurrent").value
    ocp = ctrl.range("OCPEnabled").value
    psu_port = ctrl.range("PSUPort").value
    psu_status = ctrl.range("PSUStatus").value
    
    print(f"   SetVoltage: {voltage}")
    print(f"   SetCurrent: {current}")
    print(f"   OCPEnabled: {ocp}")
    print(f"   PSUPort: {psu_port}")
    print(f"   PSUStatus: {psu_status}")
except Exception as e:
    print(f"   [FAIL] Error reading named ranges: {e}")
    print("   Named ranges might not be set up correctly")

# Test 4: Import voltampero module
print("\n4. Testing voltampero module import...")
try:
    import voltampero
    print(f"   [OK] voltampero module imported")
    print(f"   Module location: {voltampero.__file__}")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Test 5: Get controller instance
print("\n5. Getting controller instance...")
try:
    from voltampero import get_controller
    c = get_controller(simulate=True)
    print(f"   [OK] Controller created (simulated mode)")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    sys.exit(1)

# Test 6: Attach to Excel
print("\n6. Attaching controller to Excel...")
try:
    result = c.attach_excel(wb)
    print(f"   attach_excel() returned: {result}")
    if c.control_sheet:
        print(f"   [OK] control_sheet attached")
    else:
        print(f"   [FAIL] control_sheet is None")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Test 7: Connect simulated PSU
print("\n7. Connecting to simulated PSU...")
try:
    result = c.connect_psu("SIM1")
    print(f"   connect_psu() returned: {result}")
    print(f"   PSU connected: {c.psu.is_connected()}")
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Test 8: Try to set voltage and current
print("\n8. Testing set_voltage and set_current...")
try:
    test_v = 7.5
    test_a = 2.0
    
    print(f"   Setting voltage to {test_v}V...")
    result_v = c.set_voltage(test_v)
    print(f"   set_voltage() returned: {result_v}")
    
    print(f"   Setting current to {test_a}A...")
    result_a = c.set_current(test_a)
    print(f"   set_current() returned: {result_a}")
    
    # Verify the settings
    actual_v = c.psu.get_voltage_setpoint()
    actual_a = c.psu.get_current_setpoint()
    
    print(f"   Voltage setpoint: {actual_v}V (expected {test_v}V)")
    print(f"   Current setpoint: {actual_a}A (expected {test_a}A)")
    
    if abs(actual_v - test_v) < 0.01 and abs(actual_a - test_a) < 0.01:
        print(f"   [OK] Settings applied correctly!")
    else:
        print(f"   [FAIL] Settings mismatch")
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Test 9: Simulate what ApplySettings VBA macro does
print("\n9. Simulating ApplySettings VBA macro...")
try:
    # Read from Excel (like VBA does)
    voltage = ctrl.range("SetVoltage").value
    current = ctrl.range("SetCurrent").value
    ocp = ctrl.range("OCPEnabled").value
    
    print(f"   Read from Excel: V={voltage}, A={current}, OCP={ocp}")
    
    # Get fresh controller (like VBA does)
    from voltampero import get_controller
    c2 = get_controller()
    
    # Attach excel
    c2.attach_excel(wb)
    print(f"   Controller attached to Excel")
    
    # Check connection
    print(f"   PSU connected before set: {c2.psu.is_connected()}")
    
    # Apply settings
    r1 = c2.set_voltage(float(voltage))
    print(f"   set_voltage returned: {r1}")
    
    r2 = c2.set_current(float(current))
    print(f"   set_current returned: {r2}")
    
    r3 = c2.set_ocp(bool(ocp))
    print(f"   set_ocp returned: {r3}")
    
    # Check final values
    final_v = c2.psu.get_voltage_setpoint()
    final_a = c2.psu.get_current_setpoint()
    print(f"   Final setpoints: V={final_v}, A={final_a}")
    
    if abs(final_v - float(voltage)) < 0.01 and abs(final_a - float(current)) < 0.01:
        print(f"   [OK] VBA simulation successful!")
    else:
        print(f"   [FAIL] Settings not applied correctly")
    
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Test 10: Check if real hardware is expected
print("\n10. Checking hardware mode...")
psu_port_value = ctrl.range("PSUPort").value
if psu_port_value and "SIM" not in str(psu_port_value).upper():
    print(f"   [WARN] WARNING: Excel is configured for REAL hardware (port: {psu_port_value})")
    print(f"   But we tested with SIMULATED hardware above")
    print(f"   The issue might be that real PSU is not connected!")
    print(f"")
    print(f"   SOLUTION: Either:")
    print(f"   1. Click 'Test (Simulated)' button in Excel to use simulated mode")
    print(f"   2. Or connect your real PSU and click 'Connect PSU' first")
else:
    print(f"   Using simulated mode")

print("\n" + "=" * 60)
print("Debug test complete!")
print("=" * 60)
