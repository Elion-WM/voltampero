
import hid
import time
import struct

def monitor_uart():
    print("=== UART Status Monitor ===")
    dev = hid.device()
    dev.open(0x10c4, 0xea80)
    
    # 1. Enable UART
    dev.write(b'\x41\x01')
    
    # 2. Configure (Try 19200 7O1)
    baud = 19200
    baud_bytes = struct.pack('>I', baud)
    # Parity=1(Odd), Flow=0, Data=2(7bits), Stop=0(1)
    config_payload = baud_bytes + bytes([1, 0, 2, 0])
    pad = b'\x00' * (63 - len(config_payload))
    dev.write(b'\x50' + config_payload + pad)
    
    print("Monitoring UART Status (Press Ctrl+C to stop)...")
    try:
        for i in range(10):
            # Send Get Status (0x42)
            dev.write(b'\x42')
            d = dev.read(64, 100)
            if d and len(d) >= 7:
                tx = (d[1] << 8) | d[2]
                rx = (d[3] << 8) | d[4]
                err = d[5]
                print(f"[{i}] TX:{tx} RX:{rx} Err:0x{err:02x}")
            
            # Send Wakeup
            dev.write(b'\x01\x07\xab\xcd\x04\x58\x00\x01\xd4') # Wrapped GET_ID
            
            time.sleep(1.0)
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    monitor_uart()
