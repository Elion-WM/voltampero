
import hid
import time

def test_minimal():
    print("=== Minimal CP2110 Test ===")
    h = hid.device()
    h.open(0x10c4, 0xea80)
    
    # 1. Reset (0x40)
    # 2. Get Version (0x46)
    
    print("Sending Get Version (0x46)...")
    h.write(b'\x46')
    
    # Read response
    d = h.read(64, 1000)
    if d:
        print(f"Response: {bytes(d).hex()}")
    else:
        print("No response from CP2110 control endpoint.")

if __name__ == "__main__":
    test_minimal()
