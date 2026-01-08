"""
Fix Output OFF - test different commands with Excel open
"""

import xlwings as xw
import time

print("=" * 60)
print("Testing Output OFF Commands")
print("=" * 60)

# Get Excel workbook
print("\n1. Finding Excel workbook...")
wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("   [FAIL] Please open VoltAmpero.xlsm first")
    exit(1)
print(f"   [OK] Found {wb.name}")

# Get controller
print("\n2. Getting controller...")
from voltampero import get_controller
ctrl = get_controller()
ctrl.attach_excel(wb)

if not ctrl.psu.is_connected():
    print("   [FAIL] PSU not connected. Click 'Connect PSU' in Excel first")
    exit(1)
print("   [OK] PSU connected")

# Get serial object
ser = ctrl.psu.serial

# Check current state
print("\n3. Checking current output state...")
vout = ctrl.psu.get_output_voltage()
print(f"   Output voltage: {vout}V")

if vout < 0.5:
    print("   Output is already OFF. Please turn it ON first:")
    print("   1. Click 'Output ON' in Excel")
    print("   2. Wait for PSU to show voltage")
    print("   3. Run this script again")
    exit(0)

print(f"   Output is ON ({vout}V)")

# Try different OFF commands
off_commands = [
    ("OUT0", "Standard format"),
    ("OUT 0", "With space"),
    ("OUT:0", "With colon"),
    ("OUTP0", "OUTP format"),
    ("OUTP 0", "OUTP with space"),
    ("OUTP:0", "OUTP with colon"),
    ("OUTPUT0", "Full word"),
    ("OUTPUT 0", "Full word with space"),
]

print("\n4. Testing different OFF commands...")
for cmd, description in off_commands:
    print(f"\n--- Testing: {cmd} ({description}) ---")
    
    # Send command
    ser.reset_input_buffer()
    ser.write((cmd + '\r').encode('ascii'))
    time.sleep(0.5)
    
    # Check if it worked
    vout = ctrl.psu.get_output_voltage()
    print(f"   Voltage after command: {vout}V")
    
    if vout < 0.5:
        print(f"   ✓✓✓ SUCCESS! '{cmd}' turns output OFF!")
        print(f"\n   FOUND WORKING COMMAND: {cmd}")
        
        # Save the result
        with open("working_off_command.txt", "w") as f:
            f.write(f"Working OFF command: {cmd}\n")
            f.write(f"Description: {description}\n")
        
        print(f"   Saved to: working_off_command.txt")
        break
    else:
        print(f"   ✗ Failed - output still ON ({vout}V)")
        
        # Turn back ON for next test if needed
        if vout < ctrl.psu.get_voltage_setpoint() - 1.0:
            print(f"   (Voltage dropped, turning back ON for next test)")
            ser.write(b"OUT1\r")
            time.sleep(0.5)
else:
    print("\n   [FAIL] None of the commands worked!")
    print("\n   Your PSU might use a different protocol.")
    print("   Please turn OFF output manually on PSU front panel.")
    print("\n   Alternative: The PSU might need a 'lock/unlock' sequence first")

print("\n" + "=" * 60)
print("Test complete!")
print("=" * 60)
