"""
Test PSU with different line terminators
Some devices need \r, \n, or \r\n at end of commands
"""

import serial
import time

port = "COM4"
baudrate = 115200

print("=" * 60)
print("Testing Command Terminators")
print("=" * 60)

print(f"\n1. Opening {port}...")
try:
    ser = serial.Serial(
        port=port,
        baudrate=baudrate,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        timeout=1.0,
        rtscts=False,
        dsrdtr=False,
        xonxoff=False
    )
    time.sleep(0.2)
    ser.reset_input_buffer()
    ser.reset_output_buffer()
    
    # Set RTS and DTR (some devices need these)
    ser.setRTS(True)
    ser.setDTR(True)
    time.sleep(0.1)
    
    print("   [OK] Port opened with flow control lines set")
except Exception as e:
    print(f"   [FAIL] {e}")
    exit(1)

def test_command(cmd, terminator_name, terminator):
    """Test a command with specific terminator"""
    print(f"\n{terminator_name}: {repr(cmd + terminator)}")
    ser.reset_input_buffer()
    ser.write((cmd + terminator).encode('ascii'))
    time.sleep(0.15)
    
    response = ser.read(100)
    if response:
        print(f"  [OK] Response: {repr(response)}")
        try:
            decoded = response.decode('ascii', errors='ignore').strip()
            if decoded:
                print(f"  Decoded: {decoded}")
                return decoded
        except:
            pass
    else:
        print(f"  No response")
    return None

# Test *IDN? with different terminators
print("\n" + "=" * 60)
print("Testing *IDN? command")
print("=" * 60)

test_command("*IDN?", "No terminator", "")
test_command("*IDN?", "With \\n", "\n")
test_command("*IDN?", "With \\r", "\r")
test_command("*IDN?", "With \\r\\n", "\r\n")

# Test VSET1? with different terminators
print("\n" + "=" * 60)
print("Testing VSET1? command")
print("=" * 60)

test_command("VSET1?", "No terminator", "")
test_command("VSET1?", "With \\n", "\n")
test_command("VSET1?", "With \\r", "\r")
test_command("VSET1?", "With \\r\\n", "\r\n")

# Try setting voltage with different terminators
print("\n" + "=" * 60)
print("Testing VSET1:18.00 command")
print("=" * 60)

result = test_command("VSET1:18.00", "No terminator", "")
time.sleep(0.2)
test_command("VSET1?", "Verify", "")

result = test_command("VSET1:18.00", "With \\n", "\n")
time.sleep(0.2)
test_command("VSET1?", "Verify", "\n")

result = test_command("VSET1:18.00", "With \\r", "\r")
time.sleep(0.2)
test_command("VSET1?", "Verify", "\r")

result = test_command("VSET1:18.00", "With \\r\\n", "\r\n")
time.sleep(0.2)
test_command("VSET1?", "Verify", "\r\n")

# Try current
print("\n" + "=" * 60)
print("Testing ISET1:0.100 command")
print("=" * 60)

test_command("ISET1:0.100", "With \\n", "\n")
time.sleep(0.2)
test_command("ISET1?", "Verify", "\n")

print("\n" + "=" * 60)
print("Closing port...")
ser.close()
print("[OK] Done!")
print("=" * 60)
