"""
Find which COM port the PSU is actually on right now
"""

import serial
import serial.tools.list_ports

print("="*60)
print("Finding Available COM Ports")
print("="*60)

# List all COM ports
print("\nAll detected COM ports:")
ports = serial.tools.list_ports.comports()

if not ports:
    print("   [ERROR] No COM ports detected!")
    print("   Make sure USB cable is plugged in")
else:
    for i, port in enumerate(ports, 1):
        print(f"\n{i}. {port.device}")
        print(f"   Description: {port.description}")
        print(f"   Hardware ID: {port.hwid}")
        
        # Try to open it
        try:
            test_ser = serial.Serial(port.device, 115200, timeout=1)
            print(f"   Status: [OK] Port is AVAILABLE (not locked)")
            test_ser.close()
        except serial.SerialException as e:
            if "PermissionError" in str(e) or "Access" in str(e):
                print(f"   Status: [LOCKED] Access denied")
            else:
                print(f"   Status: [ERROR] {e}")

print("\n" + "="*60)
print("SOLUTION:")
print("="*60)

available_ports = []
for port in ports:
    try:
        test_ser = serial.Serial(port.device, 115200, timeout=1)
        test_ser.close()
        available_ports.append(port.device)
    except:
        pass

if available_ports:
    print(f"\nAvailable (unlocked) COM ports: {', '.join(available_ports)}")
    print(f"\nIn Excel:")
    print(f"1. Change cell B3 to: {available_ports[0]}")
    print(f"2. Click 'Connect PSU'")
    print(f"3. Try 'Apply Settings'")
else:
    print("\n[ERROR] ALL COM ports are locked!")
    print("\nYou MUST restart the computer to release them.")
    print("\nAfter restart:")
    print("1. Plug in PSU USB cable")
    print("2. Run this script again to find the port")
    print("3. Update B3 in Excel with the correct port")

print("="*60)
