"""
UNI-T UT8804E Multimeter Communication Module
USB HID communication via CP2110
Compatible with UT8803E protocol
Windows 11 compatible - no admin rights needed
"""

import time
import struct
from typing import Optional, List
from dataclasses import dataclass
from enum import Enum

try:
    import hid
    HID_AVAILABLE = True
except ImportError:
    HID_AVAILABLE = False


class MeasurementMode(Enum):
    DC_VOLTAGE = "DC V"
    AC_VOLTAGE = "AC V"
    DC_CURRENT_UA = "DC µA"
    DC_CURRENT_MA = "DC mA"
    DC_CURRENT_A = "DC A"
    AC_CURRENT_UA = "AC µA"
    AC_CURRENT_MA = "AC mA"
    AC_CURRENT_A = "AC A"
    RESISTANCE = "Ω"
    CAPACITANCE = "F"
    FREQUENCY = "Hz"
    DUTY_CYCLE = "%"
    TEMPERATURE_C = "°C"
    TEMPERATURE_F = "°F"
    DIODE = "Diode"
    CONTINUITY = "Cont"
    HFE = "hFE"
    UNKNOWN = "???"


@dataclass
class MultimeterReading:
    value: float
    unit: str
    mode: MeasurementMode
    timestamp: float
    range_str: str = ""
    overflow: bool = False
    underflow: bool = False
    hold: bool = False
    relative: bool = False
    auto_range: bool = True
    min_max: bool = False
    raw_data: bytes = b""


class UNIT_UT8804E:
    """Driver for UNI-T UT8804E Bench Multimeter"""
    
    VENDOR_ID = 0x10c4   # Silicon Labs CP2110
    PRODUCT_ID = 0xea80
    
    # Command bytes
    CMD_HOLD = bytes.fromhex("abcd04460001c2")
    CMD_BRIGHTNESS = bytes.fromhex("abcd04470001c3")
    CMD_SELECT = bytes.fromhex("abcd04480001c4")
    CMD_RANGE_MANUAL = bytes.fromhex("abcd04490001c5")
    CMD_RANGE_AUTO = bytes.fromhex("abcd044a0001c6")
    CMD_MINMAX = bytes.fromhex("abcd044b0001c7")
    CMD_EXIT_MINMAX = bytes.fromhex("abcd044c0001c8")
    CMD_REL = bytes.fromhex("abcd044d0001c9")
    CMD_D_VAL = bytes.fromhex("abcd044e0001ca")
    CMD_Q_VAL = bytes.fromhex("abcd044f0001cb")
    CMD_EXIT_DQR = bytes.fromhex("abcd04500001cc")
    CMD_R_VAL = bytes.fromhex("abcd04510001cd")
    CMD_GET_ID = bytes.fromhex("abcd04580001d4")
    
    # Mode byte mapping (based on reverse engineering)
    MODE_MAP = {
        0x00: (MeasurementMode.DC_VOLTAGE, "V"),
        0x01: (MeasurementMode.AC_VOLTAGE, "V"),
        0x02: (MeasurementMode.DC_CURRENT_UA, "µA"),
        0x03: (MeasurementMode.DC_CURRENT_MA, "mA"),
        0x04: (MeasurementMode.DC_CURRENT_A, "A"),
        0x05: (MeasurementMode.AC_CURRENT_UA, "µA"),
        0x06: (MeasurementMode.AC_CURRENT_MA, "mA"),
        0x07: (MeasurementMode.AC_CURRENT_A, "A"),
        0x08: (MeasurementMode.RESISTANCE, "Ω"),
        0x09: (MeasurementMode.CONTINUITY, "Ω"),
        0x0A: (MeasurementMode.DIODE, "V"),
        0x0B: (MeasurementMode.CAPACITANCE, "F"),
        0x0C: (MeasurementMode.FREQUENCY, "Hz"),
        0x0D: (MeasurementMode.DUTY_CYCLE, "%"),
        0x0E: (MeasurementMode.TEMPERATURE_C, "°C"),
        0x0F: (MeasurementMode.TEMPERATURE_F, "°F"),
        0x10: (MeasurementMode.HFE, ""),
    }
    
    # Range prefixes
    RANGE_PREFIX = {
        0: "",
        1: "m",   # milli
        2: "µ",   # micro
        3: "n",   # nano
        4: "k",   # kilo
        5: "M",   # mega
    }
    
    def __init__(self):
        self.device = None
        self.connected = False
        self.last_reading: Optional[MultimeterReading] = None
        self.device_id = ""
        
    @staticmethod
    def find_devices() -> List[dict]:
        """Find all connected UT8804E devices"""
        if not HID_AVAILABLE:
            return []
        try:
            devices = hid.enumerate(UNIT_UT8804E.VENDOR_ID, UNIT_UT8804E.PRODUCT_ID)
            return devices
        except:
            return []
        
    def connect(self) -> bool:
        """Connect to the multimeter"""
        if not HID_AVAILABLE:
            print("HID library not available. Install with: pip install hidapi")
            return False
            
        try:
            self.device = hid.device()
            self.device.open(self.VENDOR_ID, self.PRODUCT_ID)
            self.device.set_nonblocking(1)
            self.connected = True
            time.sleep(0.1)
            return True
        except Exception as e:
            print(f"Multimeter connection error: {e}")
            self.connected = False
            return False
    
    def disconnect(self):
        """Disconnect from the multimeter"""
        if self.device:
            try:
                self.device.close()
            except:
                pass
        self.device = None
        self.connected = False
        
    def is_connected(self) -> bool:
        """Check if connected"""
        return self.connected and self.device is not None
    
    def _send_command(self, cmd: bytes) -> bool:
        """Send a command to the multimeter"""
        if not self.is_connected():
            return False
        try:
            # HID report with report ID 0
            self.device.write(b'\x00' + cmd)
            return True
        except Exception as e:
            print(f"Multimeter command error: {e}")
            return False
    
    def _read_data(self, timeout_ms: int = 100) -> Optional[bytes]:
        """Read data from the multimeter"""
        if not self.is_connected():
            return None
        try:
            data = self.device.read(64, timeout_ms)
            return bytes(data) if data else None
        except Exception as e:
            print(f"Multimeter read error: {e}")
            return None
    
    def _parse_reading(self, data: bytes) -> Optional[MultimeterReading]:
        """Parse raw HID data into a reading"""
        if not data or len(data) < 8:
            return None
            
        try:
            # Check for valid header (0xAB 0xCD)
            if len(data) >= 2 and data[0] == 0xAB and data[1] == 0xCD:
                pass  # Valid packet
            elif len(data) >= 2 and data[0] != 0xAB:
                # Try to find header in data
                idx = data.find(b'\xab\xcd')
                if idx >= 0:
                    data = data[idx:]
                else:
                    return None
            
            if len(data) < 10:
                return None
                
            # Packet structure (approximate, based on UT8803E):
            # [0-1]: Header 0xABCD
            # [2]: Packet type/length
            # [3]: Mode byte
            # [4-7]: Value (32-bit signed int, little-endian)
            # [8]: Decimal point position
            # [9]: Range/unit modifier
            # [10]: Status flags
            
            mode_byte = data[3] & 0x1F
            mode_info = self.MODE_MAP.get(mode_byte, (MeasurementMode.UNKNOWN, ""))
            mode = mode_info[0]
            base_unit = mode_info[1]
            
            # Extract value
            if len(data) >= 8:
                raw_value = struct.unpack('<i', data[4:8])[0]
            else:
                raw_value = 0
                
            # Decimal position
            decimal_pos = data[8] if len(data) > 8 else 0
            if decimal_pos > 10:
                decimal_pos = 4  # Default
                
            value = raw_value / (10 ** decimal_pos)
            
            # Range/prefix
            range_byte = data[9] if len(data) > 9 else 0
            prefix = self.RANGE_PREFIX.get(range_byte & 0x0F, "")
            unit = prefix + base_unit
            
            # Status flags
            flags = data[10] if len(data) > 10 else 0
            overflow = (raw_value >= 59999 or raw_value <= -59999) or bool(flags & 0x01)
            hold = bool(flags & 0x02)
            relative = bool(flags & 0x04)
            auto_range = bool(flags & 0x08) or not bool(flags & 0x10)
            min_max = bool(flags & 0x20)
            
            # Build range string
            range_str = f"{prefix}{base_unit}" if prefix else base_unit
            
            return MultimeterReading(
                value=value,
                unit=unit,
                mode=mode,
                timestamp=time.time(),
                range_str=range_str,
                overflow=overflow,
                hold=hold,
                relative=relative,
                auto_range=auto_range,
                min_max=min_max,
                raw_data=data[:16] if len(data) >= 16 else data
            )
            
        except Exception as e:
            print(f"Parse error: {e}")
            return None
    
    def get_reading(self) -> Optional[MultimeterReading]:
        """Get a reading from the multimeter (streams continuously)"""
        # Device streams at ~3 readings/sec, just read latest
        data = self._read_data(500)
        if data:
            reading = self._parse_reading(data)
            if reading:
                self.last_reading = reading
                return reading
        return self.last_reading
    
    def get_value(self) -> float:
        """Get just the numeric value"""
        reading = self.get_reading()
        return reading.value if reading else 0.0
    
    def get_value_with_unit(self) -> str:
        """Get value formatted with unit"""
        reading = self.get_reading()
        if reading:
            if reading.overflow:
                return f"OL {reading.unit}"
            return f"{reading.value:.4f} {reading.unit}"
        return "--- ---"
    
    def get_device_id(self) -> str:
        """Get device identification"""
        if self._send_command(self.CMD_GET_ID):
            time.sleep(0.2)
            # Read multiple times to get response
            for _ in range(5):
                data = self._read_data(200)
                if data and len(data) > 4:
                    try:
                        # ID response after header
                        id_str = data[4:].decode('ascii', errors='ignore').strip('\x00')
                        if id_str and len(id_str) > 2:
                            self.device_id = id_str
                            return id_str
                    except:
                        pass
        return self.device_id or "UT8804E"
    
    def toggle_hold(self) -> bool:
        """Toggle hold mode"""
        return self._send_command(self.CMD_HOLD)
    
    def set_auto_range(self) -> bool:
        """Set auto range"""
        return self._send_command(self.CMD_RANGE_AUTO)
    
    def next_manual_range(self) -> bool:
        """Switch to next manual range"""
        return self._send_command(self.CMD_RANGE_MANUAL)
    
    def toggle_relative(self) -> bool:
        """Toggle relative mode"""
        return self._send_command(self.CMD_REL)
    
    def toggle_minmax(self) -> bool:
        """Toggle min/max mode"""
        return self._send_command(self.CMD_MINMAX)
    
    def exit_minmax(self) -> bool:
        """Exit min/max mode"""
        return self._send_command(self.CMD_EXIT_MINMAX)
    
    def change_brightness(self) -> bool:
        """Change display brightness"""
        return self._send_command(self.CMD_BRIGHTNESS)


class SimulatedMultimeter:
    """Simulated multimeter for testing without hardware"""
    
    def __init__(self):
        self.connected = False
        self.last_reading = None
        self._mode = MeasurementMode.DC_VOLTAGE
        self._base_value = 5.0
        import random
        self._random = random
        
    @staticmethod
    def find_devices() -> List[dict]:
        return [{"path": "SIMULATED", "product_string": "Simulated UT8804E"}]
        
    def connect(self) -> bool:
        self.connected = True
        return True
        
    def disconnect(self):
        self.connected = False
        
    def is_connected(self) -> bool:
        return self.connected
        
    def get_reading(self) -> MultimeterReading:
        value = self._base_value + self._random.uniform(-0.05, 0.05)
        reading = MultimeterReading(
            value=round(value, 4),
            unit="V",
            mode=self._mode,
            timestamp=time.time(),
            range_str="V"
        )
        self.last_reading = reading
        return reading
    
    def get_value(self) -> float:
        reading = self.get_reading()
        return reading.value
        
    def get_value_with_unit(self) -> str:
        reading = self.get_reading()
        return f"{reading.value:.4f} {reading.unit}"
        
    def get_device_id(self) -> str:
        return "SIMULATED-UT8804E"
        
    def toggle_hold(self) -> bool:
        return True
        
    def set_auto_range(self) -> bool:
        return True
        
    def set_base_value(self, value: float):
        """For simulation: set the base value"""
        self._base_value = value
