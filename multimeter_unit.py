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
            
            # Initialize CP2110 UART bridge
            self._init_uart()
            
            return True
        except Exception as e:
            print(f"Multimeter connection error: {e}")
            self.connected = False
            return False
    
    def _init_uart(self):
        """Initialize CP2110 UART for communication with multimeter"""
        if not self.device:
            return
        
        # Enable UART (Report 0x41)
        self.device.send_feature_report([0x41, 0x01])
        
        # Set UART config: 9600 baud, 8N1 (Report 0x50)
        # Format: [ReportID, Baud(4 bytes BE), Parity, FlowCtrl, DataBits, StopBits]
        config = [0x50, 0x00, 0x00, 0x25, 0x80, 0x00, 0x00, 0x03, 0x00]
        self.device.send_feature_report(config)
        
        time.sleep(0.1)
        
        # Send initialization command to start data streaming
        # Command from captured traffic: abcd040005010a00
        init_cmd = bytes.fromhex("abcd040005010a00")
        write_data = bytes([len(init_cmd)]) + init_cmd
        self.device.write(write_data)
        
        time.sleep(0.2)
        
        # Flush initial response data
        for _ in range(20):
            self.device.read(64, 50)
    
    def _wake_up_device(self):
        """Send wake-up sequence to initialize communication"""
        if not self.device:
            return
            
        # Flush any stale data in the buffer
        for _ in range(5):
            try:
                self.device.read(64, 50)
            except:
                pass
        
        time.sleep(0.05)
        
        # Send GET_ID command to wake up the device
        try:
            self.device.write(b'\x00' + self.CMD_GET_ID)
        except:
            pass
        
        time.sleep(0.1)
        
        # Read and discard wake-up response, try multiple times
        for _ in range(3):
            try:
                self.device.read(64, 100)
            except:
                pass
        
        time.sleep(0.05)
    
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
        """Send a command to the multimeter via UART"""
        if not self.is_connected():
            return False
        try:
            # CP2110 UART write format: first byte is length (1-63), followed by data
            write_data = bytes([len(cmd)]) + cmd
            self.device.write(write_data)
            return True
        except Exception as e:
            print(f"Multimeter command error: {e}")
            return False
    
    def _read_data(self, timeout_ms: int = 500) -> Optional[bytes]:
        """Read data from the multimeter, accumulating UART bytes"""
        if not self.is_connected():
            return None
        try:
            # CP2110 returns data in chunks: first byte is count, rest is UART data
            # We need to accumulate bytes to get a complete packet (at least 14 bytes)
            result = bytearray()
            deadline = time.time() + timeout_ms / 1000.0
            
            while time.time() < deadline:
                data = self.device.read(64, 50)
                if data:
                    count = data[0]
                    if 0 < count < 64:
                        result.extend(data[1:count+1])
                        # Check if we have a complete packet (starts with ABCD, need 14+ bytes)
                        idx = result.find(b'\xab\xcd')
                        if idx >= 0 and len(result) >= idx + 14:
                            return bytes(result[idx:idx+20])
            
            # Try to return partial packet if we have enough data
            if result:
                idx = result.find(b'\xab\xcd')
                if idx >= 0 and len(result) >= idx + 14:
                    return bytes(result[idx:idx+20])
            
            return None
        except Exception as e:
            print(f"Multimeter read error: {e}")
            return None
    
    def _parse_reading(self, data: bytes) -> Optional[MultimeterReading]:
        """Parse raw HID data into a reading
        
        UT8804E packet structure (from reverse engineering):
        [0-1]: Header 0xABCD
        [2]: Packet type (0x21)
        [3]: Unknown (0x00)
        [4-5]: Mode/range info (0x02 0x08 for DC Voltage)
        [6-7]: Flags (0x01 0x10)
        [8-9]: More flags (0x31 0x01 or 0x31 0x02)
        [10-13]: Main value as IEEE 754 float (little-endian, negated)
        [14]: Range indicator
        ...
        """
        if not data or len(data) < 14:
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
            
            if len(data) < 14:
                return None
            
            # Extract mode from bytes 4-5
            # UT8804E uses different mode encoding than UT8803E
            # Observed: 0x02 0x08 = DC Voltage
            mode_byte1 = data[4]
            mode_byte2 = data[5]
            
            # Map based on observed values
            if mode_byte1 == 0x02 and mode_byte2 == 0x08:
                mode = MeasurementMode.DC_VOLTAGE
                base_unit = "V"
            elif mode_byte1 == 0x02 and mode_byte2 == 0x00:
                mode = MeasurementMode.DC_VOLTAGE
                base_unit = "V"
            elif mode_byte1 == 0x03:
                mode = MeasurementMode.AC_VOLTAGE
                base_unit = "V"
            elif mode_byte1 == 0x04:
                mode = MeasurementMode.DC_CURRENT_MA
                base_unit = "mA"
            elif mode_byte1 == 0x08:
                mode = MeasurementMode.RESISTANCE
                base_unit = "Ω"
            else:
                # Fallback to old mapping
                mode_info = self.MODE_MAP.get(mode_byte1, (MeasurementMode.DC_VOLTAGE, "V"))
                mode = mode_info[0]
                base_unit = mode_info[1]
            
            # Extract value as float from bytes 10-13 (little-endian, negated)
            raw_float = struct.unpack('<f', data[10:14])[0]
            value = abs(raw_float)  # Value is stored as negative
            
            # Determine unit prefix based on range byte
            range_byte = data[14] if len(data) > 14 else 0
            prefix = ""
            if mode == MeasurementMode.DC_VOLTAGE or mode == MeasurementMode.AC_VOLTAGE:
                # Voltage mode - value is in volts
                if value < 0.001:
                    prefix = "m"
                    value *= 1000
            
            unit = prefix + base_unit if base_unit else "V"
            
            # Status flags from byte 9
            flags = data[9] if len(data) > 9 else 0
            hold = bool(flags & 0x02)
            relative = bool(flags & 0x04)
            auto_range = True
            min_max = False
            overflow = value > 99999
            
            # Build range string
            range_str = unit
            
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
                raw_data=data[:20] if len(data) >= 20 else data
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
