"""
Test what readings show during ramp with Excel open
Run this AFTER:
1. Connect PSU
2. Output ON
3. Start Logging
4. Start Ramp
"""

import xlwings as xw
import time

print("=" * 60)
print("Testing Readings During Ramp")
print("=" * 60)

# Get Excel
wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("[FAIL] Excel not open")
    exit(1)

print("[OK] Found Excel")

# Get controller
from voltampero import get_controller
ctrl = get_controller()
ctrl.attach_excel(wb)

print(f"\nPSU connected: {ctrl.psu.is_connected()}")

if not ctrl.psu.is_connected():
    print("[FAIL] PSU not connected - click Connect PSU first!")
    exit(1)

# Direct serial test
print("\n1. Testing direct serial commands:")
ser = ctrl.psu.serial

ser.reset_input_buffer()
ser.write(b"VOUT?\r")
time.sleep(0.15)
resp = ser.read(100)
v_direct = resp.decode('ascii').strip() if resp else "no response"
print(f"   VOUT? direct: {v_direct}")

ser.reset_input_buffer()
ser.write(b"IOUT?\r")
time.sleep(0.15)
resp = ser.read(100)
a_direct = resp.decode('ascii').strip() if resp else "no response"
print(f"   IOUT? direct: {a_direct}")

# Test via methods
print("\n2. Testing via PSU methods:")
v_method = ctrl.psu.get_output_voltage()
a_method = ctrl.psu.get_output_current()
print(f"   get_output_voltage(): {v_method}V")
print(f"   get_output_current(): {a_method}A")

# Test setpoints
print("\n3. Testing setpoint methods:")
v_set = ctrl.psu.get_voltage_setpoint()
a_set = ctrl.psu.get_current_setpoint()
print(f"   get_voltage_setpoint(): {v_set}V")
print(f"   get_current_setpoint(): {a_set}A")

# Test get_readings
print("\n4. Testing get_readings():")
v, a = ctrl.psu.get_readings()
print(f"   get_readings(): {v}V, {a}A")

# Test capture_reading
print("\n5. Testing _capture_reading():")
entry = ctrl._capture_reading()
if entry:
    print(f"   psu_voltage: {entry.psu_voltage}V")
    print(f"   psu_current: {entry.psu_current}A")
    print(f"   psu_setpoint_v: {entry.psu_setpoint_v}V")
    print(f"   psu_setpoint_a: {entry.psu_setpoint_a}A")

print("\n6. Checking Data sheet last row:")
data = wb.sheets["Data"]
last_row = data.range("A1").end('down').row
if last_row > 1:
    print(f"   Last row: {last_row}")
    last_voltage = data.range(f"C{last_row}").value
    last_current = data.range(f"D{last_row}").value
    print(f"   Last voltage in data: {last_voltage}V")
    print(f"   Last current in data: {last_current}A")
else:
    print("   No data rows yet")

print("\n" + "=" * 60)
print("DIAGNOSIS:")
if v_method > 0.5:
    print("[OK] PSU is outputting voltage")
    if entry and entry.psu_voltage < 0.5:
        print("[FAIL] But _capture_reading shows 0 - this is the bug!")
    else:
        print("[OK] Readings are captured correctly")
else:
    print("[FAIL] PSU output is OFF - turn it ON first!")
print("=" * 60)
