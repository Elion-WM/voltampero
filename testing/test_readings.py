"""
Test if PSU readings work correctly
Close Excel first!
"""

from psu_korad import KoradKWR102
import time

print("=" * 60)
print("Testing PSU Readings")
print("=" * 60)

psu = KoradKWR102()
print("\nConnecting to COM4...")
psu.connect("COM4")

print("\nSetting 12V, 0.5A...")
psu.set_voltage(12.0)
psu.set_current(0.5)
time.sleep(0.2)

print("\nTurning output ON...")
psu.set_output(True)
time.sleep(0.5)

print("\n1. Reading voltage setpoint...")
v_set = psu.get_voltage_setpoint()
print(f"   Voltage setpoint: {v_set}V")

print("\n2. Reading current setpoint...")
a_set = psu.get_current_setpoint()
print(f"   Current setpoint: {a_set}A")

print("\n3. Reading OUTPUT voltage...")
v_out = psu.get_output_voltage()
print(f"   Output voltage: {v_out}V")

print("\n4. Reading OUTPUT current...")
a_out = psu.get_output_current()
print(f"   Output current: {a_out}A")

print("\n5. Using get_readings()...")
v, a = psu.get_readings()
print(f"   get_readings() returned: {v}V, {a}A")

if v_out > 0.5:
    print(f"\n[OK] Readings work correctly!")
else:
    print(f"\n[FAIL] Output voltage is 0 even though output is ON")

print("\nTurning output OFF...")
psu.set_output(False)

psu.disconnect()
print("\nDone!")
