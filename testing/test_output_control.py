"""
Test output control commands for Korad KWR102
"""

import serial
import time

port = "COM4"

print("=" * 60)
print("Testing Output Control Commands")
print("=" * 60)

ser = serial.Serial(port, 115200, timeout=1.0)
ser.setRTS(True)
ser.setDTR(True)
time.sleep(0.2)
ser.reset_input_buffer()

def send(cmd):
    """Send command with \r terminator"""
    print(f"\nSending: {cmd}")
    ser.reset_input_buffer()
    ser.write((cmd + '\r').encode('ascii'))
    time.sleep(0.15)
    resp = ser.read(100)
    if resp:
        decoded = resp.decode('ascii', errors='ignore').strip()
        print(f"  Response: {decoded}")
        return decoded
    else:
        print(f"  No response (OK for set commands)")
        return None

# Check current status
print("\n1. Check current output status:")
send("STATUS?")
send("OUT?")
send("OUTP?")

# Try turning output ON with different formats
print("\n2. Try turning output ON:")
send("OUT1")     # Format 1
time.sleep(0.2)
send("STATUS?")

send("OUT:1")    # Format 2
time.sleep(0.2)
send("STATUS?")

send("OUTP 1")   # Format 3
time.sleep(0.2)
send("STATUS?")

send("OUTPUT 1") # Format 4
time.sleep(0.2)
send("STATUS?")

# Try turning output OFF
print("\n3. Try turning output OFF:")
send("OUT0")     # Format 1
time.sleep(0.2)
send("STATUS?")

send("OUT:0")    # Format 2
time.sleep(0.2)
send("STATUS?")

send("OUTP 0")   # Format 3
time.sleep(0.2)
send("STATUS?")

# Get final readings
print("\n4. Final state:")
send("VOUT?")
send("IOUT?")
send("STATUS?")

print("\n" + "=" * 60)
ser.close()
print("Test complete!")
print("Check which command format worked for STATUS?")
print("=" * 60)
