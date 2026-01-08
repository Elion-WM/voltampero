"""
Simple test - find working output command
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
    time.sleep(0.2)
    resp = ser.read(100)
    if resp:
        print(f"  Resp: {repr(resp)}")
    else:
        print(f"  (no response)")
    return resp

print("Testing output commands:\n")

# Get ID first
send("*IDN?")

# Try output commands
print("\n--- Trying different ON commands ---")
send("OUT1")
send("OUT 1")
send("OUT:1")
send("OUTP1")
send("OUTP 1")

# Check if any worked by reading voltage
print("\n--- Check if output is on ---")
resp = send("VOUT?")
if resp:
    try:
        v = float(resp.decode('ascii').strip())
        if v > 0:
            print(f"SUCCESS! Output is ON: {v}V")
        else:
            print(f"Output still OFF: {v}V")
    except:
        pass

# Try turning off
print("\n--- Turning OFF ---")
send("OUT0")
time.sleep(0.2)
send("VOUT?")

ser.close()
print("\nDone!")
