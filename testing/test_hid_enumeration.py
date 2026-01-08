
import hid

def enum_hid():
    print("=== HID Enumeration ===")
    devices = hid.enumerate(0x10c4, 0xea80)
    
    if not devices:
        print("No 0x10C4:0xEA80 devices found!")
        return

    for i, d in enumerate(devices):
        print(f"\n[Device {i}]")
        print(f"  Path: {d['path']}")
        print(f"  Serial: {d['serial_number']}")
        print(f"  Usage Page: 0x{d.get('usage_page', 0):04x}")
        print(f"  Usage: 0x{d.get('usage', 0):04x}")
        print(f"  Interface: {d['interface_number']}")
        
        # Try opening
        try:
            h = hid.device()
            h.open_path(d['path'])
            print("  -> OPEN SUCCESS")
            
            # Try Write (Report 0x46 - Get Version)
            try:
                # Add report ID 0 if Windows needs it
                h.write(b'\x46') 
                res = h.read(64, 100)
                if res:
                    print(f"  -> READ SUCCESS: {bytes(res).hex()}")
                else:
                    print("  -> READ TIMEOUT (No response)")
            except Exception as e:
                print(f"  -> WRITE ERROR: {e}")
                
            h.close()
        except Exception as e:
            print(f"  -> OPEN FAILED: {e}")

if __name__ == "__main__":
    enum_hid()
