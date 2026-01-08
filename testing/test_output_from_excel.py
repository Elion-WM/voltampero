"""
Test Output ON while Excel is open
This uses the existing controller instance
"""

import xlwings as xw
import time

print("=" * 60)
print("Test Output Control from Excel")
print("=" * 60)

# Get Excel workbook
print("\n1. Finding Excel workbook...")
wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("   [FAIL] VoltAmpero.xlsm not open")
    exit(1)
print(f"   [OK] Found {wb.name}")

# Import voltampero
print("\n2. Getting controller...")
from voltampero import get_controller

ctrl = get_controller()
ctrl.attach_excel(wb)

# Check connection
print("\n3. Checking PSU connection...")
print(f"   PSU connected: {ctrl.psu.is_connected()}")
if not ctrl.psu.is_connected():
    print("   [FAIL] PSU not connected!")
    print("   Click 'Connect PSU' button in Excel first")
    exit(1)
print("   [OK] PSU is connected")

# Get current settings
print("\n4. Reading current state...")
v_set = ctrl.psu.get_voltage_setpoint()
a_set = ctrl.psu.get_current_setpoint()
status = ctrl.psu.get_status()
print(f"   Setpoints: {v_set}V, {a_set}A")
print(f"   Output currently: {'ON' if status.output_on else 'OFF'}")

# Test the output_on command with raw serial
print("\n5. Testing Output ON command...")
print(f"   Calling ctrl.output_on()...")
result = ctrl.output_on()
print(f"   Returned: {result}")
time.sleep(0.5)

# Check if it worked
print("\n6. Verifying output state...")
status = ctrl.psu.get_status()
vout = ctrl.psu.get_output_voltage()
iout = ctrl.psu.get_output_current()
print(f"   Output status: {'ON' if status.output_on else 'OFF'}")
print(f"   Output readings: {vout}V, {iout}A")

if status.output_on:
    print("\n   [SUCCESS] Output is ON!")
    print("\n   WARNING: PSU output is now ACTIVE!")
    print("   Check your DMM - it should show voltage reading")
else:
    print("\n   [FAIL] Output is still OFF after command")
    
    # Debug - try sending command directly
    print("\n7. Debugging - sending raw command...")
    ser = ctrl.psu.serial
    
    # Get current status byte
    ser.reset_input_buffer()
    ser.write(b"STATUS?\r")
    time.sleep(0.1)
    resp = ser.read(10)
    print(f"   STATUS? response: {repr(resp)}")
    
    # Try OUT1
    print("\n   Trying: OUT1")
    ser.reset_input_buffer()
    ser.write(b"OUT1\r")
    time.sleep(0.3)
    
    # Check again
    ser.reset_input_buffer()
    ser.write(b"STATUS?\r")
    time.sleep(0.1)
    resp = ser.read(10)
    print(f"   STATUS? after OUT1: {repr(resp)}")
    
    # Re-check with get_status
    status = ctrl.psu.get_status()
    print(f"   Output now: {'ON' if status.output_on else 'OFF'}")

print("\n" + "=" * 60)
print("Test complete!")
print("=" * 60)
