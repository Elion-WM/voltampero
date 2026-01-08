"""
Test Output ON function - Excel must be closed first!
"""

from psu_korad import KoradKWR102
import time

print("=" * 60)
print("Test PSU Output Control")
print("=" * 60)

print("\n1. Connecting to PSU...")
psu = KoradKWR102()
if not psu.connect("COM4"):
    print("   [FAIL] Cannot connect - is Excel open? Close it first!")
    exit(1)
print("   [OK] Connected")

print("\n2. Reading current settings...")
v = psu.get_voltage_setpoint()
a = psu.get_current_setpoint()
print(f"   Setpoints: {v}V, {a}A")

print("\n3. Getting current status...")
status = psu.get_status()
print(f"   Output enabled: {status.output_on}")
print(f"   Mode: {status.mode}")

print("\n4. Testing Output ON command...")
result = psu.output_on()
print(f"   output_on() returned: {result}")
time.sleep(0.5)

print("\n5. Checking status after Output ON...")
status = psu.get_status()
print(f"   Output enabled: {status.output_on}")
vout = psu.get_output_voltage()
iout = psu.get_output_current()
print(f"   Output readings: {vout}V, {iout}A")

if status.output_on:
    print("\n   [SUCCESS] Output is ON!")
else:
    print("\n   [FAIL] Output is still OFF")
    print("   Testing alternative command formats...")
    
    # Try direct command
    import serial
    ser = psu.serial
    
    # Try different formats
    print("\n   Trying OUT1...")
    ser.reset_input_buffer()
    ser.write(b"OUT1\r")
    time.sleep(0.2)
    
    status = psu.get_status()
    print(f"   Output enabled: {status.output_on}")
    
    if not status.output_on:
        print("\n   Trying OUT:1...")
        ser.reset_input_buffer()
        ser.write(b"OUT:1\r")
        time.sleep(0.2)
        
        status = psu.get_status()
        print(f"   Output enabled: {status.output_on}")

print("\n6. Turning output OFF...")
result = psu.output_off()
print(f"   output_off() returned: {result}")
time.sleep(0.5)

status = psu.get_status()
print(f"   Output enabled: {status.output_on}")

print("\n7. Disconnecting...")
psu.disconnect()

print("\n" + "=" * 60)
print("Test complete! Check results above.")
print("=" * 60)
