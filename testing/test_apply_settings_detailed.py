"""
Test what happens when Apply Settings is clicked (simulate it)
"""

import xlwings as xw
import sys
sys.path.insert(0, r"C:\Users\User\GitHub\voltampero")

print("="*60)
print("Simulating Apply Settings Button Click")
print("="*60)

# Read Excel values
wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
control = wb.sheets["Control"]

voltage = control.range("B16").value
current = control.range("B17").value
port = control.range("B3").value
status = control.range("D3").value

print(f"\n1. Excel state:")
print(f"   Port: {port}")
print(f"   Status: {status}")
print(f"   Desired Voltage: {voltage}")
print(f"   Desired Current: {current}")

# Now simulate what ApplySettings does
print(f"\n2. Simulating ApplySettings button...")
print(f"   This is a FRESH Python process (controller is None)")

# Import and call get_controller like the button does
from voltampero import get_controller

print(f"\n3. Calling get_controller()...")
ctrl = get_controller()

print(f"   Controller created: {ctrl}")
print(f"   PSU type: {type(ctrl.psu).__name__}")
print(f"   PSU connected: {ctrl.psu.is_connected()}")

if not ctrl.psu.is_connected():
    print(f"\n   [ERROR] Auto-reconnect FAILED!")
    print(f"   Even though Excel shows 'Connected', Python can't connect")
    print(f"   Checking why...")
    
    # Check if port is accessible
    print(f"\n4. Testing if {port} is accessible...")
    import serial
    try:
        test_ser = serial.Serial(port, 115200, timeout=1)
        print(f"   [OK] Port {port} is accessible")
        test_ser.close()
        
        print(f"\n5. Trying manual connection...")
        result = ctrl.connect_psu(port)
        print(f"   Manual connect result: {result}")
        print(f"   Connected now: {ctrl.psu.is_connected()}")
        
    except serial.SerialException as e:
        print(f"   [ERROR] Port locked: {e}")
        print(f"   Solution: Click 'Disconnect All' in Excel first")
        sys.exit(1)

if ctrl.psu.is_connected():
    print(f"\n6. [SUCCESS] PSU is connected!")
    
    # Get current setpoints
    current_v = ctrl.psu.get_voltage_setpoint()
    current_i = ctrl.psu.get_current_setpoint()
    print(f"   Current setpoints: {current_v}V, {current_i}A")
    
    # Apply new settings
    print(f"\n7. Applying settings: {voltage}V, {current}A")
    v_result = ctrl.set_voltage(voltage)
    i_result = ctrl.set_current(current)
    print(f"   set_voltage result: {v_result}")
    print(f"   set_current result: {i_result}")
    
    # Read back
    new_v = ctrl.psu.get_voltage_setpoint()
    new_i = ctrl.psu.get_current_setpoint()
    print(f"   New setpoints: {new_v}V, {new_i}A")
    
    if new_v == voltage and new_i == current:
        print(f"\n   [SUCCESS] Settings applied correctly!")
        print(f"\n   Check your PSU display - it should show {voltage}V, {current}A")
        print(f"   Click 'Output ON' to enable output")
    else:
        print(f"\n   [WARNING] Settings don't match!")
        print(f"   Expected: {voltage}V, {current}A")
        print(f"   Got: {new_v}V, {new_i}A")
    
    # Check output status
    print(f"\n8. Checking output status...")
    status = ctrl.psu.get_status()
    print(f"   Output ON: {status.output_on}")
    if not status.output_on:
        print(f"   [INFO] Output is OFF - click 'Output ON' to see voltage/current")
    
    ctrl.disconnect_psu()
    print(f"\n9. Disconnected PSU")

print("\n" + "="*60)
