
import hid
import time

def test_write_fix():
    print("=== CP2110 Write Fix Test ===")
    h = hid.device()
    try:
        h.open(0x10c4, 0xea80)
        print("Opened.")
    except:
        print("Open failed")
        return

    # Payload to send: Reset (0x40 0x00)
    # CP2110 Output Report 0x40
    
    payloads = [
        (b'\x40\x00', "Standard (ID 40)"),
        (b'\x00\x40\x00', "Prefix 00"),
        (b'\x00\x00\x40\x00', "Prefix 00 00"),
    ]
    
    for data, label in payloads:
        print(f"Trying {label}: {data.hex()}")
        try:
            res = h.write(data)
            print(f"  -> Result: {res}")
        except Exception as e:
            print(f"  -> Error: {e}")
            
    # Also try sending Data (Report 0x01)
    print("\nTrying Data Write (0x01)...")
    # 'D' command
    data_std = b'\x01\x01D'
    data_pre = b'\x00\x01\x01D'
    
    try:
        res = h.write(data_std)
        print(f"  Std Result: {res}")
    except Exception as e:
        print(f"  Std Error: {e}")
        
    try:
        res = h.write(data_pre)
        print(f"  Pre Result: {res}")
    except Exception as e:
        print(f"  Pre Error: {e}")

if __name__ == "__main__":
    test_write_fix()
