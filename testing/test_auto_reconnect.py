"""
Test the auto-reconnect functionality
"""

import xlwings as xw
import sys
sys.path.insert(0, r"C:\Users\User\GitHub\voltampero")

print("Testing auto-reconnect...")

# Simulate what happens when Apply Settings is clicked
print("\n1. Simulating fresh Python process (like RunPython)...")
print("   _controller is None (fresh start)")

# Import and call get_controller like ApplySettings does
from voltampero import get_controller

print("\n2. Calling get_controller()...")
ctrl = get_controller()

print(f"   Controller created: {ctrl}")
print(f"   PSU type: {type(ctrl.psu).__name__}")
print(f"   PSU connected: {ctrl.psu.is_connected()}")

if ctrl.psu.is_connected():
    print("\n3. [SUCCESS] Auto-reconnect worked!")
    print(f"   Current setpoints: {ctrl.psu.get_voltage_setpoint()}V, {ctrl.psu.get_current_setpoint()}A")
    
    # Now test set_voltage
    print("\n4. Testing set_voltage(10.0)...")
    result = ctrl.set_voltage(10.0)
    print(f"   Result: {result}")
    print(f"   New voltage setpoint: {ctrl.psu.get_voltage_setpoint()}V")
    
    print("\n5. Testing set_current(2.0)...")
    result = ctrl.set_current(2.0)
    print(f"   Result: {result}")
    print(f"   New current setpoint: {ctrl.psu.get_current_setpoint()}A")
    
    print("\n" + "="*60)
    print("[SUCCESS] Auto-reconnect is working!")
    print("Apply Settings should now work correctly.")
    print("="*60)
else:
    print("\n3. [ERROR] Auto-reconnect did not work")
    print("   Make sure:")
    print("   1. Excel is open with VoltAmpero.xlsm")
    print("   2. Cell D3 (PSUStatus) shows 'Connected'")
    print("   3. Cell B3 (PSUPort) shows 'SIM1'")
