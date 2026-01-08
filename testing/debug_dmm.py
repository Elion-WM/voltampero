
import hid
import time
import struct

def debug_dmm():
    print("Enumerating HID devices...")
    for d in hid.enumerate():
        if d['vendor_id'] == 0x10c4 and d['product_id'] == 0xea80:
            print(f"FOUND UT8804E: {d}")

    print("\nAttempting connection to 0x10c4:0xea80...")
    try:
        dev = hid.device()
        dev.open(0x10c4, 0xea80)
        print("Opened successfully")
        
        # Configure CP2110 UART
        print("Enabling UART...")
        uart_enable = b'\x41\x01'
        dev.send_feature_report(uart_enable)
        
        # Try 9600
        baud = 9600
        baud_bytes = struct.pack('>I', baud)
        config_payload = baud_bytes + b'\x00\x00\x03' + (b'\x00' * 11)
        config_report = b'\x50' + config_payload
        print(f"Configuring UART 9600 8N1: {config_report.hex()}")
        dev.send_feature_report(config_report)
        time.sleep(0.5)
        
        CMD_GET_ID = bytes.fromhex("abcd04580001d4")
        dev.write(b"\x00" + CMD_GET_ID)
        
        print("Reading 10 packets (9600)...")
        for i in range(10):
            data = dev.read(64, 500)
            if data:
                 print(f"Packet {i}: {bytes(data).hex()}")
        
        # Try 19200
        print("\nRetrying with 19200 baud...")
        baud = 19200
        baud_bytes = struct.pack('>I', baud)
        config_payload = baud_bytes + b'\x00\x00\x03' + (b'\x00' * 11)
        config_report = b'\x50' + config_payload
        dev.send_feature_report(config_report)
        time.sleep(0.5)
        dev.write(b"\x00" + CMD_GET_ID)
        for i in range(10):
             data = dev.read(64, 500)
             if data:
                print(f"Packet {i} (19200): {bytes(data).hex()}")
                
        # Try 115200 (Common default)
        print("\nRetrying with 115200 baud...")
        baud = 115200
        baud_bytes = struct.pack('>I', baud)
        config_payload = baud_bytes + b'\x00\x00\x03' + (b'\x00' * 11)
        config_report = b'\x50' + config_payload
        dev.send_feature_report(config_report)
        time.sleep(0.5)
        dev.write(b"\x00" + CMD_GET_ID)
        for i in range(10):
             data = dev.read(64, 500)
             if data:
                print(f"Packet {i} (115200): {bytes(data).hex()}")

        dev.close()
    except Exception as e:
        print(f"Connection failed: {e}")

if __name__ == "__main__":
    debug_dmm()
