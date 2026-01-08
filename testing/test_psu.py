
import sys
from psu_korad import KoradKWR102

def test_psu(port):
    print(f"Attempting to connect to Korad PSU on {port}...")
    psu = KoradKWR102()
    
    try:
        if psu.connect(port):
            print("SUCCESS: Connected!")
            print(f"ID: {psu.get_id()}")
            v, i = psu.get_readings()
            print(f"Readings: {v}V, {i}A")
            psu.disconnect()
        else:
            print("FAILURE: Could not connect.")
            print("Check: Is the device turned on? Is the cable secure?")
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")

if __name__ == "__main__":
    test_psu("COM4")
