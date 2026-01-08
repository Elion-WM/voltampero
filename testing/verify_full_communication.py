"""
Comprehensive communication verification with PSU
Tests EVERY command to see what actually works
"""

import serial
import time

port = "COM4"

print("=" * 70)
print("COMPREHENSIVE PSU COMMUNICATION TEST")
print("=" * 70)

print("\nOpening port...")
ser = serial.Serial(port, 115200, timeout=0.5)
ser.setRTS(True)
ser.setDTR(True)
time.sleep(0.2)
ser.reset_input_buffer()
ser.reset_output_buffer()
print("[OK] Port opened with RTS/DTR set")

def send_cmd(cmd, description="", expect_response=True):
    """Send command and show what happens"""
    print(f"\n  >> Send: '{cmd}' {description}")
    ser.reset_input_buffer()
    
    # Show exact bytes being sent
    full_cmd = cmd + '\r'
    bytes_sent = full_cmd.encode('ascii')
    print(f"    Bytes: {repr(bytes_sent)}")
    
    ser.write(bytes_sent)
    time.sleep(0.15)
    
    response = ser.read(100)
    if response:
        print(f"    << Response: {repr(response)}")
        try:
            decoded = response.decode('ascii').strip()
            print(f"    Decoded: '{decoded}'")
            return decoded
        except:
            print(f"    (Cannot decode)")
            return None
    else:
        if expect_response:
            print(f"    << No response")
        else:
            print(f"    (No response expected for SET commands)")
        return None

# Test 1: Identification
print("\n" + "=" * 70)
print("TEST 1: Basic Communication")
print("=" * 70)
send_cmd("*IDN?", "(Get ID)")

# Test 2: Read current settings
print("\n" + "=" * 70)
print("TEST 2: Read Current Settings")
print("=" * 70)
v_set = send_cmd("VSET?", "(Voltage setpoint)")
a_set = send_cmd("ISET?", "(Current setpoint)")
v_out = send_cmd("VOUT?", "(Output voltage)")
a_out = send_cmd("IOUT?", "(Output current)")

print(f"\n  Summary: Setpoints={v_set}V/{a_set}A, Output={v_out}V/{a_out}A")

# Test 3: Verify SET commands work
print("\n" + "=" * 70)
print("TEST 3: Verify SET Commands Work")
print("=" * 70)
print("  Setting voltage to 12V...")
send_cmd("VSET:12.00", "(Set 12V)", expect_response=False)
time.sleep(0.2)
new_v = send_cmd("VSET?", "(Verify)")
print(f"  Result: Voltage setpoint is now {new_v}V")

print("\n  Setting current to 0.5A...")
send_cmd("ISET:0.500", "(Set 0.5A)", expect_response=False)
time.sleep(0.2)
new_a = send_cmd("ISET?", "(Verify)")
print(f"  Result: Current setpoint is now {new_a}A")

# Test 4: Output control
print("\n" + "=" * 70)
print("TEST 4: Output Control - Comprehensive")
print("=" * 70)

# First ensure output is OFF
print("\n  Step 1: Ensure output is OFF first")
v_before = send_cmd("VOUT?")
print(f"  Current output: {v_before}V")

# Turn ON
print("\n  Step 2: Turn output ON")
send_cmd("OUT1", "(Turn ON)", expect_response=False)
time.sleep(0.5)
v_on = send_cmd("VOUT?", "(Check voltage)")
print(f"  Output voltage after OUT1: {v_on}V")

if float(v_on) > 0.5:
    print(f"  [OK] SUCCESS: Output is ON ({v_on}V)")
    
    # Now test every possible OFF command
    print("\n  Step 3: Testing ALL possible OFF commands...")
    
    off_tests = [
        "OUT0",
        "OUT 0", 
        "OUT:0",
        "OUTP0",
        "OUTP 0",
        "OUTP:0",
        "OUTPUT0",
        "OUTPUT 0",
        "OUTPUT:0",
        "SYST:REM 0",
        "OUTPut 0",
        "OUTPut:STATe 0",
    ]
    
    for test_cmd in off_tests:
        print(f"\n  >> Testing: {test_cmd}")
        
        # Make sure output is ON before each test
        v_check = send_cmd("VOUT?")
        if float(v_check) < 0.5:
            print(f"    Output was OFF, turning back ON...")
            send_cmd("OUT1", expect_response=False)
            time.sleep(0.5)
        
        # Try the OFF command
        send_cmd(test_cmd, expect_response=False)
        time.sleep(0.5)
        
        # Check if it worked
        v_after = send_cmd("VOUT?")
        if float(v_after) < 0.5:
            print(f"    *** SUCCESS! '{test_cmd}' turns output OFF!")
            print(f"\n    WORKING OFF COMMAND FOUND: {test_cmd}")
            
            # Save result
            with open("working_off_command.txt", "w") as f:
                f.write(f"Working OFF command: {test_cmd}\n")
            break
        else:
            print(f"    [X] Failed - still {v_after}V")
    else:
        print(f"\n  [X] NONE of the OFF commands worked!")
        print(f"\n  This PSU might:")
        print(f"    1. Have a hardware lock preventing remote OFF")
        print(f"    2. Require a special unlock command first")
        print(f"    3. Not support remote OFF at all")
else:
    print(f"  [X] Output didn't turn ON! Something wrong with OUT1 command")

# Test 5: Check if there's a LOCK/UNLOCK
print("\n" + "=" * 70)
print("TEST 5: Check for LOCK/UNLOCK Commands")
print("=" * 70)
send_cmd("LOCK?")
send_cmd("UNLOCK")
send_cmd("*LOCK?")
send_cmd("SYST:LOCK?")
send_cmd("SYST:COMM:REM?")

# Test 6: Try remote mode
print("\n" + "=" * 70)
print("TEST 6: Try Enabling Remote Mode")
print("=" * 70)
send_cmd("SYST:REM", expect_response=False)
time.sleep(0.2)
print("  Now try OFF after remote mode:")
send_cmd("OUT0", expect_response=False)
time.sleep(0.5)
v_final = send_cmd("VOUT?")
if float(v_final) < 0.5:
    print(f"  [OK] Remote mode helped! Output is OFF")
else:
    print(f"  Still ON: {v_final}V")

print("\n" + "=" * 70)
print("TEST COMPLETE")
print("=" * 70)
print("\nManually turn off the PSU now using front panel button")

ser.close()
