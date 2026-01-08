"""
Test which OFF command works
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
    resp = ser.read(100)
    if resp:
        print(f"  Resp: {repr(resp)}")
    else:
        print(f"  (no response)")
    return resp

def check_output():
    """Check if output is on"""
    resp = send("VOUT?")
    if resp:
        try:
            v = float(resp.decode('ascii').strip())
            status = "ON" if v > 0.1 else "OFF"
            print(f"  --> Output: {status} ({v}V)")
            return v > 0.1
        except:
            pass
    return False

print("=" * 60)
print("Testing Output OFF Commands")
print("=" * 60)

# Make sure output is ON first
print("\n1. Turning output ON...")
send("OUT1")
time.sleep(0.3)
is_on = check_output()

if not is_on:
    print("   [WARN] Output is not ON, can't test OFF commands")
    ser.close()
    exit(0)

# Try different OFF command formats
print("\n2. Trying different OFF commands:\n")

formats = [
    "OUT0",
    "OUT 0", 
    "OUT:0",
    "OUTP0",
    "OUTP 0",
    "OUTPUT0",
    "OUTPUT 0",
]

for cmd in formats:
    print(f"--- Testing: {cmd} ---")
    send(cmd)
    time.sleep(0.3)
    
    if not check_output():
        print(f"   ✓ SUCCESS! {cmd} turns output OFF")
        
        # Turn it back on for next test
        print(f"   (turning back ON for next test)")
        send("OUT1")
        time.sleep(0.3)
    else:
        print(f"   ✗ Failed - output still ON")
    print()

# Final state
print("\n3. Final state:")
send("OUT0")
time.sleep(0.3)
check_output()

ser.close()
print("\n" + "=" * 60)
print("Check results above to see which command worked")
print("=" * 60)
