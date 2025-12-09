"""
Korad KWR102 Power Supply Communication Module

Refactored to use PyMeasure for a consistent, testable driver interface.
"""

import time
from typing import Optional, Tuple, List
from dataclasses import dataclass

import serial
import serial.tools.list_ports

try:
    from pymeasure.adapters import SerialAdapter
    from pymeasure.instruments import Instrument
    PYMEASURE_AVAILABLE = True
except Exception:
    PYMEASURE_AVAILABLE = False


@dataclass
class PSUStatus:
    voltage: float
    current: float
    voltage_setpoint: float
    current_setpoint: float
    output_on: bool
    ocp_on: bool
    ovp_on: bool
    mode: str


class _KoradInstrument(Instrument):
    """PyMeasure Instrument for Korad KWR102 over serial."""

    def __init__(self, port: str, baudrate: int = 115200, timeout: float = 1.0):
        adapter = SerialAdapter(port, baudrate=baudrate, timeout=timeout)
        super().__init__(adapter)

    # Low-level helpers kept similar to the legacy driver
    def _send(self, cmd: str) -> Optional[str]:
        try:
            # Korad typically does not require terminators; keep as-is
            if '?' in cmd:
                return self.ask(cmd).strip()
            else:
                self.write(cmd)
                return ""
        except Exception as e:
            print(f"PSU command error: {e}")
            return None


class KoradKWR102:
    """Wrapper driver exposing the legacy API backed by PyMeasure."""

    def __init__(self, port: str = "", baudrate: int = 115200, timeout: float = 1.0):
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self._inst: Optional[_KoradInstrument] = None
        self._ocp_enabled = False
        self._ovp_enabled = False

    @staticmethod
    def list_ports() -> List[str]:
        ports = serial.tools.list_ports.comports()
        return [p.device for p in ports]

    def connect(self, port: str = "") -> bool:
        if port:
            self.port = port
        if not self.port:
            return False
        try:
            if not PYMEASURE_AVAILABLE:
                # Fallback to raw serial for environments without pymeasure
                ser = serial.Serial(
                    port=self.port,
                    baudrate=self.baudrate,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    timeout=self.timeout,
                )
                # Wrap minimal subset inside a tiny adapter for compatibility
                class _Fallback(_KoradInstrument):
                    def __init__(self, ser):
                        self.adapter = type("_A", (), {"connection": ser, "write": ser.write, "read": ser.read})()
                        Instrument.__init__(self, self.adapter)
                self._inst = _Fallback(ser)
            else:
                self._inst = _KoradInstrument(self.port, self.baudrate, self.timeout)
            time.sleep(0.1)
            return True
        except Exception as e:
            print(f"PSU connection error: {e}")
            self._inst = None
            return False

    def disconnect(self):
        try:
            if self._inst is not None:
                conn = getattr(self._inst.adapter, "connection", None)
                if conn is not None:
                    try:
                        conn.close()
                    except Exception:
                        pass
        finally:
            self._inst = None

    def is_connected(self) -> bool:
        conn = getattr(self._inst.adapter, "connection", None) if self._inst else None
        return bool(conn) and getattr(conn, "is_open", True)

    def _send_command(self, cmd: str) -> Optional[str]:
        if not self._inst:
            return None
        return self._inst._send(cmd)
    
    def get_identification(self) -> str:
        """Get device identification string"""
        return self._send_command("*IDN?") or "Unknown"
    
    def set_voltage(self, voltage: float) -> bool:
        """Set output voltage (V)"""
        voltage = max(0, min(voltage, 60))  # Clamp to safe range
        cmd = f"VSET1:{voltage:05.2f}"
        return self._send_command(cmd) is not None
    
    def get_voltage_setpoint(self) -> float:
        """Get voltage setpoint (V)"""
        response = self._send_command("VSET1?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def get_output_voltage(self) -> float:
        """Get actual output voltage (V)"""
        response = self._send_command("VOUT1?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def set_current(self, current: float) -> bool:
        """Set current limit (A)"""
        current = max(0, min(current, 30))  # Clamp to safe range
        cmd = f"ISET1:{current:05.3f}"
        return self._send_command(cmd) is not None
    
    def get_current_setpoint(self) -> float:
        """Get current setpoint (A)"""
        response = self._send_command("ISET1?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def get_output_current(self) -> float:
        """Get actual output current (A)"""
        response = self._send_command("IOUT1?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def set_output(self, on: bool) -> bool:
        """Turn output on or off"""
        cmd = "OUT1" if on else "OUT0"
        return self._send_command(cmd) is not None
    
    def output_on(self) -> bool:
        """Turn output on"""
        return self.set_output(True)
    
    def output_off(self) -> bool:
        """Turn output off"""
        return self.set_output(False)
    
    def set_ocp(self, on: bool) -> bool:
        """Turn Over Current Protection on or off"""
        cmd = "OCP1" if on else "OCP0"
        result = self._send_command(cmd) is not None
        if result:
            self._ocp_enabled = on
        return result
    
    def set_ovp(self, on: bool) -> bool:
        """Turn Over Voltage Protection on or off"""
        cmd = "OVP1" if on else "OVP0"
        result = self._send_command(cmd) is not None
        if result:
            self._ovp_enabled = on
        return result
    
    def get_status(self) -> PSUStatus:
        """Get full status of the power supply"""
        status_response = self._send_command("STATUS?")
        
        output_on = False
        mode = "CV"
        
        if status_response:
            try:
                status = ord(status_response[0]) if len(status_response) > 0 else 0
                output_on = bool(status & 0x40)
                mode = "CC" if (status & 0x01) else "CV"
            except:
                pass
        
        return PSUStatus(
            voltage=self.get_output_voltage(),
            current=self.get_output_current(),
            voltage_setpoint=self.get_voltage_setpoint(),
            current_setpoint=self.get_current_setpoint(),
            output_on=output_on,
            ocp_on=self._ocp_enabled,
            ovp_on=self._ovp_enabled,
            mode=mode
        )
    
    def get_readings(self) -> Tuple[float, float]:
        """Get voltage and current readings as tuple (V, A)"""
        return (self.get_output_voltage(), self.get_output_current())


class VoltageRamp:
    """Voltage ramping functionality for PSU with multi-cycle support"""
    
    def __init__(self, psu: KoradKWR102):
        self.psu = psu
        self.running = False
        self.paused = False
        self.current_voltage = 0.0
        self.current_cycle = 0
        self.total_cycles = 1
        self.progress_callback = None
        
    def configure(self, start_v: float, end_v: float, duration_s: float,
                  cycles: int = 1, delay_between_s: float = 0.0,
                  ping_pong: bool = False):
        """Configure ramp parameters"""
        self.start_v = start_v
        self.end_v = end_v
        self.duration_s = duration_s
        self.cycles = cycles  # 0 = infinite
        self.delay_between_s = delay_between_s
        self.ping_pong = ping_pong
        
    def start(self, start_v: float, end_v: float, duration_s: float,
              cycles: int = 1, delay_between_s: float = 0.0,
              ping_pong: bool = False, step_interval: float = 0.1,
              progress_callback=None):
        """
        Start voltage ramp
        cycles: number of repetitions (0 = infinite)
        ping_pong: if True, alternates direction each cycle
        progress_callback(cycle, total_cycles, current_v, progress_pct)
        """
        self.configure(start_v, end_v, duration_s, cycles, delay_between_s, ping_pong)
        self.progress_callback = progress_callback
        self.running = True
        self.paused = False
        self.current_cycle = 0
        
        cycle_count = 0
        direction = 1  # 1 = forward, -1 = reverse
        
        while self.running:
            cycle_count += 1
            self.current_cycle = cycle_count
            
            # Determine start/end for this cycle
            if ping_pong and direction == -1:
                cycle_start = end_v
                cycle_end = start_v
            else:
                cycle_start = start_v
                cycle_end = end_v
            
            # Execute single ramp
            self._run_single_ramp(cycle_start, cycle_end, duration_s, 
                                   step_interval, cycle_count, cycles)
            
            if not self.running:
                break
                
            # Check if we've completed all cycles
            if cycles > 0 and cycle_count >= cycles:
                break
                
            # Flip direction for ping-pong
            if ping_pong:
                direction *= -1
                
            # Delay between cycles
            if delay_between_s > 0 and self.running:
                delay_steps = int(delay_between_s / 0.1)
                for _ in range(delay_steps):
                    if not self.running:
                        break
                    while self.paused and self.running:
                        time.sleep(0.1)
                    time.sleep(0.1)
        
        self.running = False
        
    def _run_single_ramp(self, start_v: float, end_v: float, duration_s: float,
                         step_interval: float, current_cycle: int, total_cycles: int):
        """Execute a single voltage ramp"""
        if duration_s <= 0:
            self.psu.set_voltage(end_v)
            return
            
        steps = int(duration_s / step_interval)
        if steps < 1:
            steps = 1
            
        voltage_step = (end_v - start_v) / steps
        self.current_voltage = start_v
        self.psu.set_voltage(self.current_voltage)
        
        for i in range(steps + 1):
            if not self.running:
                break
                
            while self.paused and self.running:
                time.sleep(0.1)
                
            if i > 0:
                time.sleep(step_interval)
                self.current_voltage = start_v + (voltage_step * i)
                self.current_voltage = round(self.current_voltage, 3)
                self.psu.set_voltage(self.current_voltage)
            
            if self.progress_callback:
                progress_pct = (i / steps) * 100
                self.progress_callback(current_cycle, total_cycles, 
                                       self.current_voltage, progress_pct)
                
    def stop(self):
        """Stop the voltage ramp"""
        self.running = False
        self.paused = False
        
    def pause(self):
        """Pause the voltage ramp"""
        self.paused = True
        
    def resume(self):
        """Resume the voltage ramp"""
        self.paused = False
        
    def is_running(self) -> bool:
        """Check if ramp is running"""
        return self.running


# Simulated PSU for testing without hardware
class SimulatedPSU:
    """Simulated PSU for testing without hardware"""
    
    def __init__(self):
        self._connected = False
        self._voltage_set = 0.0
        self._current_set = 0.0
        self._output_on = False
        self._ocp_on = False
        self._ovp_on = False
        
    @staticmethod
    def list_ports() -> List[str]:
        return ["SIM1", "SIM2"]
        
    def connect(self, port: str = "") -> bool:
        self._connected = True
        return True
        
    def disconnect(self):
        self._connected = False
        
    def is_connected(self) -> bool:
        return self._connected
        
    def get_identification(self) -> str:
        return "SIMULATED-KWR102"
        
    def set_voltage(self, voltage: float) -> bool:
        self._voltage_set = voltage
        return True
        
    def get_voltage_setpoint(self) -> float:
        return self._voltage_set
        
    def get_output_voltage(self) -> float:
        if self._output_on:
            import random
            return self._voltage_set + random.uniform(-0.01, 0.01)
        return 0.0
        
    def set_current(self, current: float) -> bool:
        self._current_set = current
        return True
        
    def get_current_setpoint(self) -> float:
        return self._current_set
        
    def get_output_current(self) -> float:
        if self._output_on:
            import random
            return min(self._current_set, 0.1 + random.uniform(-0.01, 0.01))
        return 0.0
        
    def set_output(self, on: bool) -> bool:
        self._output_on = on
        return True
        
    def output_on(self) -> bool:
        return self.set_output(True)
        
    def output_off(self) -> bool:
        return self.set_output(False)
        
    def set_ocp(self, on: bool) -> bool:
        self._ocp_on = on
        return True
        
    def set_ovp(self, on: bool) -> bool:
        self._ovp_on = on
        return True
        
    def get_status(self) -> PSUStatus:
        return PSUStatus(
            voltage=self.get_output_voltage(),
            current=self.get_output_current(),
            voltage_setpoint=self._voltage_set,
            current_setpoint=self._current_set,
            output_on=self._output_on,
            ocp_on=self._ocp_on,
            ovp_on=self._ovp_on,
            mode="CV"
        )
        
    def get_readings(self) -> Tuple[float, float]:
        return (self.get_output_voltage(), self.get_output_current())
