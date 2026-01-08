"""
Debug Output ON from Excel - run with Excel open
"""

import xlwings as xw
import time

print("=" * 60)
print("Debug Excel Output ON")
print("=" * 60)

# Find Excel
print("\n1. Finding Excel...")
try:
    wb = None
    for book in xw.books:
        if "VoltAmpero" in book.name:
            wb = book
            break
    
    if not wb:
        print("   [FAIL] Excel not open")
        exit(1)
    print(f"   [OK] Found {wb.name}")
except:
    print("   [FAIL] Excel not running")
    exit(1)

# Get controller
print("\n2. Getting controller...")
from voltampero import get_controller
ctrl = get_controller()
ctrl.attach_excel(wb)

print(f"   PSU connected: {ctrl.psu.is_connected()}")
if not ctrl.psu.is_connected():
    print("   [FAIL] PSU not connected in controller")
    exit(1)

# Check current state
print("\n3. Checking current state BEFORE output_on()...")
ser = ctrl.psu.serial
print(f"   Serial port open: {ser.is_open if ser else 'None'}")
print(f"   Port: {ctrl.psu.port}")

# Read voltage
ser.reset_input_buffer()
ser.write(b"VOUT?\r")
time.sleep(0.15)
resp = ser.read(100)
v_before = resp.decode('ascii').strip() if resp else "?"
print(f"   Output voltage before: {v_before}V")

# Test the actual method that Excel VBA calls
print("\n4. Calling ctrl.output_on() (same as Excel VBA)...")
try:
    result = ctrl.output_on()
    print(f"   output_on() returned: {result}")
except Exception as e:
    print(f"   [ERROR] {e}")
    import traceback
    traceback.print_exc()

time.sleep(0.5)

# Check what actually happened
print("\n5. Checking state AFTER output_on()...")
ser.reset_input_buffer()
ser.write(b"VOUT?\r")
time.sleep(0.15)
resp = ser.read(100)
v_after = resp.decode('ascii').strip() if resp else "?"
print(f"   Output voltage after: {v_after}V")

if v_after != v_before and float(v_after) > 0.5:
    print("\n   [OK] Output turned ON successfully!")
else:
    print("\n   [FAIL] Output didn't turn on")
    
    # Debug - send command manually
    print("\n6. Manual debugging - sending OUT1 directly...")
    ser.reset_input_buffer()
    ser.write(b"OUT1\r")
    time.sleep(0.5)
    
    ser.reset_input_buffer()
    ser.write(b"VOUT?\r")
    time.sleep(0.15)
    resp = ser.read(100)
    v_manual = resp.decode('ascii').strip() if resp else "?"
    print(f"   After manual OUT1: {v_manual}V")
    
    if float(v_manual) > 0.5:
        print("\n   Manual command works! Problem is in Python code path.")
    else:
        print("\n   Manual command also fails! PSU might have lock/protection.")

print("\n" + "=" * 60)
print("Test complete")
print("=" * 60)
