# Communication Protocols - Quick Reference

This directory contains complete communication protocol documentation for devices used in the VoltAmpero project.

## Available Protocol Documents

### 1. [Korad KWR102 Power Supply](KORAD_KWR102_PROTOCOL.md)
**Interface**: USB-to-Serial (CH340 chip)  
**Protocol**: ASCII commands with `\r` terminator  
**Key Features**:
- Voltage control: 0-60V
- Current limiting: 0-5A
- Output ON/OFF control
- Real-time voltage/current readings
- Thread-safe for concurrent access

**Quick Start**:
```python
import serial
ser = serial.Serial('COM3', 115200)
ser.setRTS(True)
ser.setDTR(True)
ser.write(b'VSET:12.00\r')  # Set 12V
ser.write(b'OUT:1\r')        # Output ON
```

[📖 Full Documentation →](KORAD_KWR102_PROTOCOL.md)

---

### 2. [UNI-T UT8804E Bench Multimeter](UNIT_UT8804E_PROTOCOL.md)
**Interface**: USB HID (CP2110 bridge)  
**Protocol**: Binary packets via HID with UART bridge  
**Key Features**:
- Multiple measurement modes (V, A, Ω, F, Hz, °C)
- Continuous data streaming
- Remote control commands
- No driver installation needed (HID)

**Quick Start**:
```python
import hid
device = hid.device()
device.open(0x10c4, 0xea80)
device.send_feature_report([0x41, 0x01])  # Enable UART
# Initialize and read measurements
```

[📖 Full Documentation →](UNIT_UT8804E_PROTOCOL.md)

---

## Protocol Comparison

| Feature | Korad KWR102 | UNI-T UT8804E |
|---------|--------------|---------------|
| **Interface** | Serial (Virtual COM) | HID (USB) |
| **Driver Needed** | Yes (CH340) | No (HID built-in) |
| **Protocol** | ASCII text | Binary packets |
| **Baud Rate** | 115200 | 9600 (UART bridge) |
| **Terminator** | `\r` | Packet-based |
| **Threading** | Requires lock | Single-threaded HID |
| **Read Speed** | Fast (~50ms/query) | Medium (~200ms/reading) |
| **Admin Rights** | May be needed | Not needed |

---

## Common Patterns

### Pattern 1: Initialize Device
```python
# PSU (Serial)
import serial
psu = serial.Serial('COM3', 115200, timeout=1.0)
psu.setRTS(True)
psu.setDTR(True)

# DMM (HID)
import hid
dmm = hid.device()
dmm.open(0x10c4, 0xea80)
dmm.send_feature_report([0x41, 0x01])  # Enable UART
```

### Pattern 2: Read Measurement
```python
# PSU
psu.write(b'VOUT?\r')
voltage = float(psu.read(100).decode().strip())

# DMM
data = dmm.read(64, 500)
# Parse binary packet (see full docs)
```

### Pattern 3: Thread-Safe Access
```python
import threading

lock = threading.Lock()

def safe_read():
    with lock:
        psu.write(b'VOUT?\r')
        return float(psu.read(100))
```

---

## Reverse Engineering Tools Used

Both protocols were reverse-engineered using:

1. **Wireshark + USBPcap**
   - Capture USB traffic while using official software
   - Analyze URB packets for command patterns
   - Identify protocol structure

2. **Serial/HID Monitors**
   - RealTerm (serial)
   - Python with debug output
   - Device Manager (check VID/PID)

3. **Community Resources**
   - Sigrok wiki for Korad protocols
   - CP2110 datasheet for HID bridge
   - Forum discussions

---

## Troubleshooting Quick Reference

### Korad KWR102
| Problem | Solution |
|---------|----------|
| No response | Set RTS/DTR high |
| Wrong commands | Use `VSET:` not `VSET1:` for V2.3 |
| Slow logging | Reduce timeout to 100ms, delay to 20ms |

### UNI-T UT8804E
| Problem | Solution |
|---------|----------|
| Device not found | Check VID:PID 0x10c4:0xea80 |
| No data | Send initialization: `abcd040005010a00` |
| Garbled data | Flush buffer, wait for `AB CD` header |

---

## Protocol Development Notes

### Why These Protocols?
- **Korad KWR102**: Official protocol not documented, V2.3 differs from other models
- **UT8804E**: Uses HID instead of serial, protocol completely undocumented

### Verification Method
1. Captured USB traffic with official software
2. Identified command patterns and responses
3. Tested with Python to verify behavior
4. Documented edge cases and timing requirements
5. Created reusable drivers (see `psu_korad.py` and `multimeter_unit.py`)

### Future Compatibility
- **Korad**: May work with other models (KA3005P, etc.) with command format changes
- **UNI-T**: Likely compatible with UT8803E, possibly UT61E family

---

## Usage in VoltAmpero Project

These protocols are implemented in:
- **`psu_korad.py`** - Complete Korad KWR102 driver
- **`multimeter_unit.py`** - Complete UT8804E driver  
- **`voltampero.py`** - High-level integration with logging and ramping

Example integration:
```python
from psu_korad import KoradKWR102
from multimeter_unit import UNIT_UT8804E

# Initialize devices
psu = KoradKWR102(port='COM3')
psu.connect()

dmm = UNIT_UT8804E()
dmm.connect()

# Use together
psu.set_voltage(12.0)
psu.output_on()

reading = dmm.get_reading()
print(f"PSU set: {psu.get_voltage_setpoint()}V")
print(f"DMM reads: {reading.value}{reading.unit}")
```

---

## License and Attribution

These protocol documents are provided for educational and development purposes. They were created through reverse engineering and testing with physical devices.

**If you use these protocols in your project:**
- Attribution appreciated but not required
- Share improvements/corrections back to community
- Test thoroughly with your specific hardware

**Original reverse engineering**: 2026-01-08  
**Hardware tested**: Korad KWR102 V2.3, UNI-T UT8804E

---

## Contributing

Found a bug or have improvements?
- Test with your device and document differences
- Submit corrections or additions
- Share new command discoveries
- Report firmware version differences

---

*Save yourself hours of Wireshark analysis - use these docs!* 🔌⚡
