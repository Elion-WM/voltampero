"""
Test the actual Python methods that Excel calls
Close Excel first!
"""

from psu_korad import KoradKWR102
import time

print("=" * 60)
print("Testing PSU Methods Directly")
print("=" * 60)

print("\n1. Creating PSU instance...")
psu = KoradKWR102()

print("\n2. Connecting to COM4...")
result = psu.connect("COM4")
print(f"   connect() returned: {result}")

if not result:
    print("   [FAIL] Cannot connect")
    exit(1)

print("\n3. Testing output_on() method...")
print(f"   Before: is_connected = {psu.is_connected()}")
result = psu.output_on()
print(f"   output_on() returned: {result}")
time.sleep(0.5)

# Check if it worked
vout = psu.get_output_voltage()
print(f"   Output voltage: {vout}V")

if vout > 0.5:
    print(f"   [OK] Output is ON")
    
    print("\n4. Testing output_off() method...")
    result = psu.output_off()
    print(f"   output_off() returned: {result}")
    time.sleep(0.5)
    
    vout = psu.get_output_voltage()
    print(f"   Output voltage: {vout}V")
    
    if vout < 0.5:
        print(f"   [OK] Output is OFF")
    else:
        print(f"   [FAIL] Output still ON")
else:
    print(f"   [FAIL] Output didn't turn ON")
    
    # Debug the set_output method
    print("\n5. Debugging set_output method...")
    print(f"   Calling psu.set_output(True)...")
    result = psu.set_output(True)
    print(f"   Returned: {result}")
    time.sleep(0.5)
    
    vout = psu.get_output_voltage()
    print(f"   Output voltage: {vout}V")

print("\n6. Disconnecting...")
psu.disconnect()

print("\n" + "=" * 60)
print("Done")
print("=" * 60)
