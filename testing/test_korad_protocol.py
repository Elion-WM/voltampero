"""
Test correct Korad protocol - now that we know \r is needed
"""

import serial
import time

port = "COM4"

print("=" * 60)
print("Korad KWR102 Protocol Test")
print("Serial: 000000166271")
print("=" * 60)

ser = serial.Serial(port, 115200, timeout=1.0)
ser.setRTS(True)
ser.setDTR(True)
time.sleep(0.2)
ser.reset_input_buffer()
ser.reset_output_buffer()

def send(cmd):
    """Send command with \r terminator"""
    print(f"\nSending: {cmd}")
    ser.reset_input_buffer()
    ser.write((cmd + '\r').encode('ascii'))
    time.sleep(0.1)
    resp = ser.read(100)
    if resp:
        decoded = resp.decode('ascii', errors='ignore').strip()
        print(f"  Response: {decoded}")
        return decoded
    else:
        print(f"  No response")
        return None

# Test ID
print("\n1. Get ID:")
send("*IDN?")

# Try different voltage command formats
print("\n2. Testing voltage commands:")
send("VSET1?")  # Query current setpoint
send("VSET?")   # Alternative
send("VMAX?")   # Maybe max voltage

# Look for documentation - try RCL (recall) commands
print("\n3. Testing alternative query formats:")
send("ISET1?")
send("VOUT1?")
send("IOUT1?")

# Maybe it needs channel selection first?
print("\n4. Testing with explicit channel:")
send("INST OUT1")
time.sleep(0.1)
send("VSET?")
send("ISET?")

# Try the original format from other Korad models
print("\n5. Testing classic Korad format:")
send("VSET1:18.00")  # Set voltage
time.sleep(0.2)
send("VSET1?")        # Query it back

# Try without channel number
send("VSET:18.00")
time.sleep(0.2)
send("VSET?")

# Try with leading zeros
send("VSET1:018.00")
time.sleep(0.2)
send("VSET1?")

# Maybe it needs a different format
print("\n6. Testing SCPI-style commands:")
send("SOUR:VOLT 18.0")
time.sleep(0.2)
send("SOUR:VOLT?")

send("VOLT 18.0")
time.sleep(0.2)
send("VOLT?")

# Try measuring actual output
print("\n7. Testing measurement commands:")
send("MEAS:VOLT?")
send("MEAS:CURR?")
send("MEAS:ALL?")

# Check status
print("\n8. Status commands:")
send("STAT?")
send("STATUS?")
send("*STB?")

# Try output control
print("\n9. Output control:")
send("OUT?")        # Query output state
send("OUTP?")      # Alternative

print("\n" + "=" * 60)
print("If voltage commands don't work, the PSU might:")
print("1. Need to be unlocked first")
print("2. Use a proprietary protocol")  
print("3. Need specific initialization")
print("=" * 60)

ser.close()
