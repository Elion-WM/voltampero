"""
Test different baudrates to find the correct one
"""

import serial
import time

port = "COM4"
baudrates = [9600, 19200, 38400, 57600, 115200]

print("=" * 60)
print("Testing Different Baudrates")
print("=" * 60)

for baud in baudrates:
    print(f"\nTesting {baud} baud...")
    try:
        ser = serial.Serial(
            port=port,
            baudrate=baud,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=0.5
        )
        time.sleep(0.1)
        ser.reset_input_buffer()
        ser.reset_output_buffer()
        
        # Try *IDN?
        ser.write(b"*IDN?")
        time.sleep(0.2)
        response = ser.read(100)
        
        if response:
            print(f"  [OK] Got response: {repr(response)}")
            try:
                decoded = response.decode('ascii', errors='ignore')
                print(f"  Decoded: {decoded}")
            except:
                pass
        else:
            print(f"  No response")
            
        ser.close()
        
    except Exception as e:
        print(f"  Error: {e}")

print("\n" + "=" * 60)
print("If all show 'No response', the PSU might:")
print("1. Not be a Korad KWR102 (wrong model?)")
print("2. Use a different protocol")
print("3. Have a hardware issue")
print("=" * 60)
