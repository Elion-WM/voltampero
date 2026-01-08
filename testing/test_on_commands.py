"""
Test different ON commands to find which one works
Close Excel first!
"""

import serial
import time

port = "COM4"
ser = serial.Serial(port, 115200, timeout=0.5)
ser.setRTS(True)
ser.setDTR(True)
time.sleep(0.2)

def send(cmd):
    print(f"Send: {cmd}")
    ser.reset_input_buffer()
    ser.write((cmd + '\r').encode('ascii'))
    time.sleep(0.3)

def check_voltage():
    ser.reset_input_buffer()
    ser.write(b"VOUT?\r")
    time.sleep(0.15)
    resp = ser.read(100)
    if resp:
        v = resp.decode('ascii').strip()
        print(f"  Voltage: {v}V")
        return float(v)
    return 0

print("=" * 60)
print("Testing ON Commands")
print("=" * 60)

# Make sure output is OFF first
print("\n1. Turning output OFF first...")
send("OUT:0")
time.sleep(0.5)
v = check_voltage()

if v > 0.5:
    print(f"  WARNING: Output still ON ({v}V) - turn it off manually on PSU")
    print("  Then run this test again")
    ser.close()
    exit(1)

print(f"  [OK] Output is OFF")

# Test different ON commands
on_commands = [
    "OUT1",      # No colon
    "OUT:1",     # With colon (to match OUT:0)
    "OUT 1",     # With space
    "OUTP1",
    "OUTP:1",
]

for cmd in on_commands:
    print(f"\n2. Testing: {cmd}")
    send(cmd)
    time.sleep(0.5)
    
    v = check_voltage()
    
    if v > 0.5:
        print(f"  *** SUCCESS! '{cmd}' turns output ON! ({v}V)")
        
        # Turn it back OFF for next test
        send("OUT:0")
        time.sleep(0.5)
        check_voltage()
        
        print(f"\n  WORKING ON COMMAND: {cmd}")
        
        # Save to file
        with open("working_on_command.txt", "w") as f:
            f.write(f"Working ON command: {cmd}\n")
        break
    else:
        print(f"  Failed - still OFF")
else:
    print(f"\n[FAIL] None of the ON commands worked!")

ser.close()
print("\n" + "=" * 60)
