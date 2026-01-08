"""
Test raw serial commands to PSU
Debug communication protocol
"""

import serial
import time

print("=" * 60)
print("Raw PSU Communication Test")
print("=" * 60)

port = "COM4"
baudrate = 115200

print(f"\n1. Opening {port} at {baudrate} baud...")
try:
    ser = serial.Serial(
        port=port,
        baudrate=baudrate,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        timeout=1.0
    )
    print("   [OK] Port opened")
    time.sleep(0.1)
    ser.reset_input_buffer()
    ser.reset_output_buffer()
except Exception as e:
    print(f"   [FAIL] {e}")
    exit(1)

# Test different baudrates if needed
def send_and_read(cmd, description):
    """Send command and read response"""
    print(f"\n{description}")
    print(f"   Sending: {repr(cmd)}")
    ser.reset_input_buffer()
    ser.write(cmd.encode('ascii'))
    time.sleep(0.1)
    
    response = ser.read(100)
    print(f"   Raw response: {repr(response)}")
    
    if response:
        try:
            decoded = response.decode('ascii').strip()
            print(f"   Decoded: {decoded}")
            return decoded
        except:
            print(f"   (Cannot decode as ASCII)")
            return None
    else:
        print(f"   No response")
        return None

# Test commands
print("\n" + "=" * 60)
print("Testing PSU Commands")
print("=" * 60)

# Get ID
idn = send_and_read("*IDN?", "2. Get Identification (*IDN?)")

# Get voltage setpoint - try different formats
send_and_read("VSET1?", "3. Get voltage setpoint (VSET1?)")
send_and_read("VSET?", "4. Get voltage setpoint (VSET?)")

# Get current setpoint
send_and_read("ISET1?", "5. Get current setpoint (ISET1?)")
send_and_read("ISET?", "6. Get current setpoint (ISET?)")

# Set voltage to 18V - try different formats
print("\n" + "=" * 60)
print("Setting Voltage to 18V")
print("=" * 60)

send_and_read("VSET1:18.00", "7. Set voltage format 1 (VSET1:18.00)")
time.sleep(0.2)
send_and_read("VSET1?", "8. Verify voltage setpoint")

send_and_read("VSET1:018.00", "9. Set voltage format 2 (VSET1:018.00)")
time.sleep(0.2)
send_and_read("VSET1?", "10. Verify voltage setpoint")

# Set current to 0.1A
print("\n" + "=" * 60)
print("Setting Current to 0.1A")
print("=" * 60)

send_and_read("ISET1:00.100", "11. Set current format 1 (ISET1:00.100)")
time.sleep(0.2)
send_and_read("ISET1?", "12. Verify current setpoint")

send_and_read("ISET1:0.100", "13. Set current format 2 (ISET1:0.100)")
time.sleep(0.2)
send_and_read("ISET1?", "14. Verify current setpoint")

# Get output readings
print("\n" + "=" * 60)
print("Reading Output Values")
print("=" * 60)

send_and_read("VOUT1?", "15. Get output voltage (VOUT1?)")
send_and_read("IOUT1?", "16. Get output current (IOUT1?)")
send_and_read("STATUS?", "17. Get status (STATUS?)")

# Try model-specific commands
print("\n" + "=" * 60)
print("Trying Alternative Commands")
print("=" * 60)

# Some Korad models use different commands
send_and_read("VSET:", "18. Set voltage (VSET:)")
send_and_read("VSET?", "19. Get voltage")

print("\n" + "=" * 60)
print("Closing port...")
ser.close()
print("[OK] Done!")
print("=" * 60)

print("\nANALYSIS:")
print("- Check which commands returned valid responses")
print("- Look for numeric values in the 'Decoded' output")
print("- If all responses are empty, baudrate or protocol might be wrong")
