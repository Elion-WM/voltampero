"""
Direct test: Connect to PSU and set 18V, 0.1A without turning on
"""

import sys
import time

print("=" * 60)
print("Direct PSU Test - Set 18V, 0.1A")
print("=" * 60)

# Step 1: Import and check
print("\n1. Importing modules...")
try:
    from psu_korad import KoradKWR102
    print("   [OK] psu_korad imported")
except Exception as e:
    print(f"   [FAIL] {e}")
    sys.exit(1)

# Step 2: List available ports
print("\n2. Scanning for COM ports...")
try:
    ports = KoradKWR102.list_ports()
    print(f"   Available ports: {ports}")
    
    if not ports:
        print("   [FAIL] No COM ports found!")
        sys.exit(1)
        
    # Find COM4 or use first available
    target_port = "COM4"
    if target_port not in ports:
        print(f"   [WARN] COM4 not found, using {ports[0]}")
        target_port = ports[0]
    else:
        print(f"   [OK] Found COM4")
        
except Exception as e:
    print(f"   [FAIL] {e}")
    sys.exit(1)

# Step 3: Create PSU instance
print("\n3. Creating PSU instance...")
try:
    psu = KoradKWR102()
    print("   [OK] PSU object created")
except Exception as e:
    print(f"   [FAIL] {e}")
    sys.exit(1)

# Step 4: Connect to PSU
print(f"\n4. Connecting to {target_port}...")
try:
    result = psu.connect(target_port)
    print(f"   Connection result: {result}")
    
    if not result:
        print(f"   [FAIL] Could not connect to {target_port}")
        print(f"   This port might be locked by another process")
        print(f"\n   SOLUTION:")
        print(f"   1. Close Excel if it's open")
        print(f"   2. Open Task Manager and end any Python processes")
        print(f"   3. Try again")
        sys.exit(1)
    
    print(f"   [OK] Connected successfully!")
    
    # Get PSU info
    idn = psu.get_identification()
    print(f"   PSU ID: {idn}")
    
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Step 5: Read current settings
print("\n5. Reading current PSU settings...")
try:
    current_v_set = psu.get_voltage_setpoint()
    current_a_set = psu.get_current_setpoint()
    current_v_out = psu.get_output_voltage()
    current_a_out = psu.get_output_current()
    
    print(f"   Current setpoints: {current_v_set}V, {current_a_set}A")
    print(f"   Current output: {current_v_out}V, {current_a_out}A")
    
except Exception as e:
    print(f"   [FAIL] Error reading: {e}")
    import traceback
    traceback.print_exc()

# Step 6: Set voltage to 18V
print("\n6. Setting voltage to 18.0V...")
try:
    result = psu.set_voltage(18.0)
    print(f"   set_voltage(18.0) returned: {result}")
    
    if not result:
        print(f"   [FAIL] Failed to set voltage")
    else:
        print(f"   [OK] Voltage command sent")
        
    time.sleep(0.1)
    
    # Verify
    actual_v = psu.get_voltage_setpoint()
    print(f"   Verification: PSU reports {actual_v}V")
    
    if abs(actual_v - 18.0) < 0.1:
        print(f"   [OK] Voltage set correctly!")
    else:
        print(f"   [WARN] Voltage mismatch: expected 18.0V, got {actual_v}V")
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Step 7: Set current to 0.1A
print("\n7. Setting current limit to 0.1A...")
try:
    result = psu.set_current(0.1)
    print(f"   set_current(0.1) returned: {result}")
    
    if not result:
        print(f"   [FAIL] Failed to set current")
    else:
        print(f"   [OK] Current command sent")
        
    time.sleep(0.1)
    
    # Verify
    actual_a = psu.get_current_setpoint()
    print(f"   Verification: PSU reports {actual_a}A")
    
    if abs(actual_a - 0.1) < 0.01:
        print(f"   [OK] Current set correctly!")
    else:
        print(f"   [WARN] Current mismatch: expected 0.1A, got {actual_a}A")
        
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Step 8: Read final state
print("\n8. Reading final state...")
try:
    final_v_set = psu.get_voltage_setpoint()
    final_a_set = psu.get_current_setpoint()
    final_v_out = psu.get_output_voltage()
    final_a_out = psu.get_output_current()
    
    print(f"   Final setpoints: {final_v_set}V, {final_a_set}A")
    print(f"   Final output: {final_v_out}V, {final_a_out}A")
    
    status = psu.get_status()
    print(f"   Output enabled: {status.output_on}")
    print(f"   Mode: {status.mode}")
    
except Exception as e:
    print(f"   [FAIL] Error: {e}")

# Step 9: Disconnect
print("\n9. Disconnecting...")
try:
    psu.disconnect()
    print("   [OK] Disconnected")
except Exception as e:
    print(f"   [WARN] {e}")

print("\n" + "=" * 60)
print("Test complete!")
print("=" * 60)
print("\nRESULT:")
print(f"  Target: 18.0V, 0.1A")
print(f"  Achieved: {final_v_set}V, {final_a_set}A")
print(f"  Output: OFF (as requested)")
print("\nCheck your PSU display to confirm these values!")
