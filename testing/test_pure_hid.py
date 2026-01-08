
import hid
import time

def test_pure_hid():
    print("=== Pure HID Read Test ===")
    try:
        dev = hid.device()
        dev.open(0x10c4, 0xea80)
        print("Device Opened")
        
        # Don't configure anything. Just try to read.
        # Some UNI-T devices stream constantly on the interrupt endpoint.
        
        print("Reading for 10 seconds...")
        start = time.time()
        while time.time() - start < 10:
            d = dev.read(64, 100)
            if d:
                print(f"DATA: {bytes(d).hex()}")
                return
            
        print("No data received in pure HID mode.")
        
        print("\nTrying Feature Report 0x60 Enable (UT61E style)...")
        # UT61E sometimes needs this magic feature report
        try:
             # Report 0x60, payload...
             payload = bytes.fromhex("00 09 00 00 00") # random guess based on UT protocols
             dev.send_feature_report(b'\x60' + payload)
        except:
             pass
             
        # Read again
        d = dev.read(64, 1000)
        if d:
             print(f"DATA after Enable: {bytes(d).hex()}")
        else:
             print("Still no data.")
             
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_pure_hid()
