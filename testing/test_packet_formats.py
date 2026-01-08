
import hid
import time
import struct

def test_packet_formats():
    print("=== Deep Packet Test for UNI-T UT8804E ===")
    
    try:
        dev = hid.device()
        dev.open(0x10c4, 0xea80)
        print("Device Opened (0x10c4:0xea80)")
    except Exception as e:
        print(f"Failed to open device: {e}")
        return

    # Helper to print reads
    def try_read(label):
        try:
            d = dev.read(64, 500)
            if d:
                print(f"[{label}] READ SUCCESS: {bytes(d).hex()}")
                return True
        except:
            pass
        return False

    # 1. Try Enable UART (0x41)
    print("\n[1] Enabling UART (Report 0x41)...")
    try:
        # Report 41, Enable(1)
        dev.write(b'\x41\x01')
        time.sleep(0.1)
    except Exception as e:
        print(f"Write failed: {e}")

    # 2. Try various baud rates with proper Config Report (0x50)
    # Struct: Baud(4), Parity(1), Flow(1), DataBits(1), StopBits(1)
    # CP2110 Config Report is 64 bytes total (Report ID 0x50 + 63 bytes payload)
    
    configs = [
        (19200, 1, 2, "19200 7O1"), 
        (9600, 1, 2, "9600 7O1"),
        (19200, 0, 3, "19200 8N1"),
        (9600, 0, 3, "9600 8N1"),
        (2400, 0, 3, "2400 8N1")
    ]

    for baud, parity, bits, label in configs:
        print(f"\n--- Testing Config: {label} ---")
        
        # Build payload
        baud_bytes = struct.pack('>I', baud)
        # Parity: 0=N, 1=O, 2=E, 3=M, 4=S
        # Flow: 0=None
        # DataBits: 0=5, 1=6, 2=7, 3=8
        # StopBits: 0=1
        payload = baud_bytes + bytes([parity, 0, bits, 0])
        pad = b'\x00' * (63 - len(payload))
        report = b'\x50' + payload + pad
        
        try:
            dev.write(report)
            time.sleep(0.2)
            
            # Flush
            dev.write(b'\x43\x03') 
            
            # Send Stimuli
            # A. GET_ID (UT8804E specific?)
            CMD_GET_ID = bytes.fromhex("abcd04580001d4") # ID=0x58
            # Wrap in Report 0x01 (Write UART) -> Len(1) + Data(...)
            # CP2110 Write Report: ID(0x01) + Len(1) + Data
            
            # Packet 1: Raw write (if direct HID)
            # dev.write(b'\x00' + CMD_GET_ID) 
            
            # Packet 2: CP2110 UART Write Wrapper
            l = len(CMD_GET_ID)
            uart_write_report = b'\x01' + bytes([l]) + CMD_GET_ID
            dev.write(uart_write_report)
            
            if try_read(f"{label} - After GET_ID"): continue

            # B. SEND_DATA (UT61E+ style)
            CMD_SEND = bytes.fromhex("ab cd 03 5e 01 d9") # ID=0x5E
            l = len(CMD_SEND)
            uart_write_report = b'\x01' + bytes([l]) + CMD_SEND
            dev.write(uart_write_report)
            
            if try_read(f"{label} - After SEND_DATA"): continue
            
            # C. D0-Series simple 'D' command
            dev.write(b'\x01\x01D')
            if try_read(f"{label} - After 'D'"): continue

        except Exception as e:
            print(f"Error in loop: {e}")

    print("\nTest Complete")
    dev.close()

if __name__ == "__main__":
    test_packet_formats()
