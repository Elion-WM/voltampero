"""
Test if logging captures data during ramp
Run with Excel OPEN
"""

import xlwings as xw
import time

print("=" * 60)
print("Test Logging During Ramp")
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

print(f"PSU connected: {ctrl.psu.is_connected()}")

if not ctrl.psu.is_connected():
    print("[FAIL] PSU not connected")
    exit(1)

# Check current readings
print("\n1. Current PSU readings:")
v, a = ctrl.get_psu_readings()
print(f"   Output: {v}V, {a}A")

v_set = ctrl.psu.get_voltage_setpoint()
a_set = ctrl.psu.get_current_setpoint()
print(f"   Setpoints: {v_set}V, {a_set}A")

# Capture a reading
print("\n2. Testing _capture_reading():")
entry = ctrl._capture_reading()
if entry:
    print(f"   PSU voltage: {entry.psu_voltage}V")
    print(f"   PSU current: {entry.psu_current}A")
    print(f"   PSU setpoint V: {entry.psu_setpoint_v}V")
    print(f"   PSU setpoint A: {entry.psu_setpoint_a}A")
else:
    print("   [FAIL] No entry captured")

# Check if output is ON
if entry and entry.psu_voltage < 0.5:
    print("\n[WARN] PSU output is OFF - readings will be 0")
    print("Turn output ON to see actual voltage values!")
else:
    print("\n[OK] PSU output is ON - readings are real")

print("\n" + "=" * 60)
print("If readings show 0, turn output ON in Excel first!")
print("=" * 60)
