"""
Diagnose real PSU connection and commands
"""

import xlwings as xw
import sys
sys.path.insert(0, r"C:\Users\User\GitHub\voltampero")

print("="*60)
print("Diagnosing Real PSU Connection")
print("="*60)

# Check Excel state
wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
control = wb.sheets["Control"]

print("\n1. Excel cell values:")
port = control.range("B3").value
psu_status = control.range("D3").value
set_v = control.range("B16").value
set_i = control.range("B17").value

print(f"   B3 (PSU Port): {port}")
print(f"   D3 (PSU Status): {psu_status}")
print(f"   B16 (Set Voltage): {set_v}")
print(f"   B17 (Set Current): {set_i}")

# Test direct serial connection
print("\n2. Testing direct serial connection to PSU...")
try:
    import serial
    import serial.tools.list_ports
    
    # List all available ports
    print("   Available COM ports:")
    ports = serial.tools.list_ports.comports()
    for p in ports:
        print(f"      {p.device}: {p.description}")
    
    # Try to connect to the specified port
    if port:
        print(f"\n3. Attempting to connect to {port}...")
        try:
            ser = serial.Serial(
                port=port,
                baudrate=9600,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            print(f"   [OK] Serial port opened: {ser.is_open}")
            
            # Test identification command
            print("\n4. Testing PSU identification (*IDN? command)...")
            ser.write(b'*IDN?\n')
            import time
            time.sleep(0.1)
            response = ser.read(100)
            if response:
                print(f"   PSU Response: {response}")
            else:
                print("   [WARNING] No response from PSU")
                print("   This might not be a Korad PSU or wrong baud rate")
            
            # Try to read voltage setpoint
            print("\n5. Reading current voltage setpoint from PSU...")
            ser.write(b'VSET1?\n')
            time.sleep(0.1)
            response = ser.read(100)
            if response:
                print(f"   Current voltage setpoint: {response}")
            else:
                print("   [WARNING] No response")
            
            # Try to set voltage
            print(f"\n6. Setting voltage to {set_v}V...")
            cmd = f"VSET1:{set_v:.2f}\n".encode()
            print(f"   Sending command: {cmd}")
            ser.write(cmd)
            time.sleep(0.1)
            
            # Read back
            ser.write(b'VSET1?\n')
            time.sleep(0.1)
            response = ser.read(100)
            print(f"   Voltage setpoint now: {response}")
            
            # Try to set current
            print(f"\n7. Setting current to {set_i}A...")
            cmd = f"ISET1:{set_i:.3f}\n".encode()
            print(f"   Sending command: {cmd}")
            ser.write(cmd)
            time.sleep(0.1)
            
            # Read back
            ser.write(b'ISET1?\n')
            time.sleep(0.1)
            response = ser.read(100)
            print(f"   Current setpoint now: {response}")
            
            # Check output status
            print("\n8. Checking output status...")
            ser.write(b'STATUS?\n')
            time.sleep(0.1)
            response = ser.read(100)
            if response:
                status_byte = response[0] if response else 0
                output_on = (status_byte & 0x40) != 0
                print(f"   Output is: {'ON' if output_on else 'OFF'}")
                if not output_on:
                    print("   [INFO] Output is OFF - you won't see voltage/current")
                    print("   Click 'Output ON' button in Excel to enable")
            
            ser.close()
            print("\n   [OK] Direct serial communication test complete")
            
        except serial.SerialException as e:
            print(f"   [ERROR] Cannot open {port}: {e}")
            print("   Check:")
            print("   - Is the PSU powered on?")
            print("   - Is the USB cable connected?")
            print("   - Is the correct COM port selected?")
            print("   - Is another program using the port?")
        except Exception as e:
            print(f"   [ERROR] {e}")
            import traceback
            traceback.print_exc()
    else:
        print("   [ERROR] No port specified in B3")
        
except ImportError:
    print("   [ERROR] pyserial not installed")
    print("   Run: pip install pyserial")

# Now test via voltampero module
print("\n9. Testing via voltampero module...")
try:
    from psu_korad import KoradKWR102
    
    psu = KoradKWR102()
    print(f"   Created PSU object: {psu}")
    
    if port:
        print(f"   Connecting to {port}...")
        result = psu.connect(port)
        print(f"   Connect result: {result}")
        
        if result:
            print(f"   Setting voltage to {set_v}V...")
            psu.set_voltage(set_v)
            
            print(f"   Setting current to {set_i}A...")
            psu.set_current(set_i)
            
            # Read back
            actual_v = psu.get_voltage_setpoint()
            actual_i = psu.get_current_setpoint()
            print(f"   Readback: {actual_v}V, {actual_i}A")
            
            if actual_v == set_v and actual_i == set_i:
                print("\n   [SUCCESS] voltampero module works correctly!")
            else:
                print(f"\n   [WARNING] Values don't match:")
                print(f"   Expected: {set_v}V, {set_i}A")
                print(f"   Got: {actual_v}V, {actual_i}A")
            
            psu.disconnect()
        else:
            print("   [ERROR] Could not connect via voltampero module")
            
except Exception as e:
    print(f"   [ERROR] {e}")
    import traceback
    traceback.print_exc()

print("\n" + "="*60)
print("DIAGNOSIS COMPLETE")
print("="*60)
