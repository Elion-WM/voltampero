
import hid

def inspect_hid():
    print("=== CP2110 / UT8804E Inspector ===")
    
    # Open device
    try:
        dev = hid.device()
        dev.open(0x10c4, 0xea80)
        print("Device Opened Successfully")
        
        print(f"Manufacturer: {dev.get_manufacturer_string()}")
        print(f"Product: {dev.get_product_string()}")
        print(f"Serial: {dev.get_serial_number_string()}")
        
        # CP2110 / CP2114 specific: Get Version Information
        # Report ID 0x46 (Get Version)
        # Returns: ID(1) + PartNum(1) + Version(1)
        try:
            dev.write(b'\x46')
            d = dev.read(64, 100)
            if d and len(d) > 2:
                print(f"Chip Part Number: 0x{d[1]:02x} (should be 0x0A for CP2110)")
                print(f"Chip Version: 0x{d[2]:02x}")
        except Exception as e:
            print(f"Failed to read version: {e}")

        # Check UART Status (Report 0x42)
        # Returns: ID(1) + TX_FIFO(2) + RX_FIFO(2) + Error(1) + Break(1)
        try:
            dev.write(b'\x42')
            d = dev.read(64, 100)
            if d and len(d) >= 8:
                tx_count = (d[1] << 8) | d[2]
                rx_count = (d[3] << 8) | d[4]
                err_status = d[5]
                print(f"UART FIFO Status: TX={tx_count}, RX={rx_count}")
                print(f"UART Error Status: 0x{err_status:02x}")
                if err_status & 0x01: print("  - Parity Error")
                if err_status & 0x02: print("  - Frame Error")
                if err_status & 0x04: print("  - Overrun Error")
                if err_status & 0x10: print("  - Break Error")
        except Exception as e:
            print(f"Failed to read status: {e}")
            
    except Exception as e:
        print(f"Inspection failed: {e}")

if __name__ == "__main__":
    inspect_hid()
