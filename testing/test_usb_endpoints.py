
import usb.core
import usb.util
import time

def test_usb_endpoints():
    print("=== USB Endpoint Inspector ===")
    # UT8804E ID
    dev = usb.core.find(idVendor=0x10c4, idProduct=0xea80)
    
    if dev is None:
        print("Device not found")
        return

    print(f"Device found: {dev.idVendor:04x}:{dev.idProduct:04x}")
    
    for cfg in dev:
        print(f"Configuration {cfg.bConfigurationValue}")
        for intf in cfg:
            print(f"  Interface {intf.bInterfaceNumber}, Alt {intf.bAlternateSetting}")
            print(f"    Class: {intf.bInterfaceClass} (0x{intf.bInterfaceClass:02x})")
            print(f"    Subclass: {intf.bInterfaceSubClass} (0x{intf.bInterfaceSubClass:02x})")
            print(f"    Protocol: {intf.bInterfaceProtocol} (0x{intf.bInterfaceProtocol:02x})")
            
            if intf.bInterfaceClass == 0xFE and intf.bInterfaceSubClass == 0x03:
                print("    -> FOUND USBTMC Interface!")
                
            for ep in intf:
                print(f"      Endpoint 0x{ep.bEndpointAddress:02x}")
                print(f"        Attr: 0x{ep.bmAttributes:02x}")
                print(f"        MaxPacket: {ep.wMaxPacketSize}")

if __name__ == "__main__":
    # We might need libusb, but let's try with pyusb if installed, 
    # otherwise fallback to simple print
    try:
        test_usb_endpoints()
    except Exception as e:
        print(f"PyUSB Error: {e}")
        print("Note: This script requires 'pyusb' and a libusb backend.")
