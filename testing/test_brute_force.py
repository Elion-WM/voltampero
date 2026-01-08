
import hid
import time
import struct

def test_brute_force():
    print("=== CP2110 Brute Force Scanner ===")
    
    try:
        h = hid.device()
        h.open(0x10c4, 0xea80)
        print("Device Opened.")
    except Exception as e:
        print(f"Failed to open: {e}")
        return

    # Valid Bauds: 2400, 4800, 9600, 19200, 38400, 57600, 115200
    # Parity: 0(N), 1(O), 2(E), 3(M), 4(S)
    # Bits: 0(5), 1(6), 2(7), 3(8)
    
    bauds = [19200, 9600, 115200]
    parities = [0, 1, 2] 
    data_bits = [2, 3]   
    flow_controls = [0, 1] # 0=None, 1=Hardware (RTS/CTS)
    
    # Enable UART once
    try:
        h.send_feature_report(b'\x41\x01')
        time.sleep(0.1)
    except:
        pass

    total = len(bauds) * len(parities) * len(data_bits) * len(flow_controls)
    count = 0
    
    for baud in bauds:
        for parity in parities:
            for bits in data_bits:
                for flow in flow_controls:
                    count += 1
                    cfg_str = f"{baud} {'N' if parity==0 else 'O' if parity==1 else 'E'}{7 if bits==2 else 8}1 Flow={'HW' if flow else 'No'}"
                    print(f"[{count}/{total}] Trying {cfg_str}...", end="\r")
                    
                    try:
                        # Config Report
                        baud_bytes = struct.pack('>I', baud)
                        # Payload: Baud(4) + Parity(1) + Flow(1) + DataBits(1) + StopBits(1)
                        # Flow Control: 0=None, 1=Hardware
                        payload = baud_bytes + bytes([parity, flow, bits, 0])
                        pad = b'\x00' * (63 - len(payload))
                        h.send_feature_report(b'\x50' + payload + pad)
                        time.sleep(0.05)
                        
                        # Purge
                        h.send_feature_report(b'\x43\x03')
                        
                        # Stimulate (Send ID query)
                        # Wrap in UART Write (0x01)
                        cmd = bytes.fromhex("ab cd 04 58 00 01 d4") # Get ID
                        h.write(b'\x01' + bytes([len(cmd)]) + cmd)
                        
                        # Read
                        d = h.read(64, 200) # Short timeout
                        if d:
                            print(f"\n!!! MATCH FOUND !!! Config: {cfg_str}")
                            print(f"Data: {bytes(d).hex()}")
                            return
                            
                    except Exception as e:
                        # print(f"Err: {e}")
                        pass
                    
    print("\nScan complete. No data received.")

if __name__ == "__main__":
    test_brute_force()
