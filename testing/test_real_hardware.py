
import sys
import time
from voltampero import VoltAmpero

def test_hardware():
    print("=== VoltAmpero Hardware Test ===")
    
    # Initialize in REAL mode (not simulation)
    app = VoltAmpero(simulate=False)
    
    # 1. Test DMM (Auto-detect USB HID)
    print("\n[1/2] Testing Multimeter (UNI-T UT8804E)...")
    try:
        if app.connect_dmm():
            print("SUCCESS: Multimeter connected!")
            print(f"Device ID: {app.dmm.get_device_id()}")
            
            print("Reading data for 5 seconds...")
            for i in range(5):
                val = app.dmm.get_value_with_unit()
                print(f"  Reading {i+1}: {val}")
                time.sleep(1)
            
            app.disconnect_dmm()
        else:
            print("FAILURE: Could not connect to Multimeter.")
            print("Ensure USB is connected and device is powered on.")
    except Exception as e:
        print(f"ERROR: DMM Test crashed: {e}")

    # 2. Test PSU (Serial Ports)
    print("\n[2/2] Testing Power Supply (Korad KWR102)...")
    ports = app.list_com_ports()
    print(f"Available COM ports: {ports}")
    
    if not ports:
        print("FAILURE: No COM ports found. Check PSU USB connection.")
    else:
        for port in ports:
            print(f"Attempting connection on {port}...")
            if app.connect_psu(port):
                print(f"SUCCESS: Connected to PSU on {port}")
                status = app.get_psu_status()
                print(f"  Status: {status}")
                app.disconnect_psu()
                break
            else:
                print(f"  Failed to connect on {port}")

    print("\n=== Test Complete ===")

if __name__ == "__main__":
    test_hardware()
