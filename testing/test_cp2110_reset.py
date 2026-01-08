
import hid
import time
import struct

def test_reset_write():
    print("=== CP2110 Reset & Output Report Test ===")
    
    try:
        h = hid.device()
        h.open(0x10c4, 0xea80)
        print("Device Opened.")
    except Exception as e:
        print(f"Failed to open: {e}")
        return

    def send_cmd(name, data):
        print(f"Sending {name} (len={len(data)}): {data.hex()}")
        try:
            # Try Write (Output Report)
            res = h.write(data)
            print(f"  -> Write Result: {res} bytes")
        except Exception as e:
            print(f"  -> Write Error: {e}")

    # 1. Reset Device (Report 0x40)
    # Payload: 0x00 (Soft reset? Spec says just report 40)
    # Some docs say Report 0x40 + 0x00
    send_cmd("RESET (0x40)", b'\x40\x00')
    time.sleep(1.0) # Wait for reset
    
    # Re-open if needed (Windows might loose handle on reset)
    try:
        h.close()
        time.sleep(0.5)
        h = hid.device()
        h.open(0x10c4, 0xea80)
        print("Device Re-Opened after Reset.")
    except:
        print("Could not re-open (might be same handle).")

    # 2. Enable UART (Report 0x41)
    # Byte 0: 0x41 (ID)
    # Byte 1: 0x01 (Enable)
    send_cmd("ENABLE UART (0x41)", b'\x41\x01')
    time.sleep(0.1)

    # 3. Configure UART (Report 0x50)
    # 19200 7O1 (Odd Parity)
    baud = 19200
    baud_bytes = struct.pack('>I', baud)
    # Parity=1(Odd), Flow=0, Data=2(7bits), Stop=0(1)
    payload = baud_bytes + bytes([1, 0, 2, 0])
    pad = b'\x00' * (63 - len(payload))
    report = b'\x50' + payload + pad
    
    send_cmd("CONFIG 19200 7O1 (0x50)", report)
    time.sleep(0.2)
    
    # 4. Flush FIFOs (Report 0x43)
    send_cmd("FLUSH (0x43)", b'\x43\x03')
    
    print("\nReading Loop (5s)...")
    start = time.time()
    while time.time() - start < 5:
        try:
            # Send Keep-Alive / Stimulus (GET_ID)
            # Wrapped in UART Write (0x01)
            # DMM CMD: AB CD 04 58 00 01 D4
            dmm_cmd = bytes.fromhex("ab cd 04 58 00 01 d4")
            wrap = b'\x01' + bytes([len(dmm_cmd)]) + dmm_cmd
            h.write(wrap)
            
            d = h.read(64, 200)
            if d:
                print(f"DATA RECEIVED: {bytes(d).hex()}")
                return
        except Exception as e:
            pass
            
    print("No data.")
    h.close()

if __name__ == "__main__":
    test_reset_write()
