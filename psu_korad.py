"""
Korad KWR102 Power Supply Communication Module
Serial communication using SCPI-like commands
Windows 11 compatible - no admin rights needed
"""

import serial
import serial.tools.list_ports
import time
import threading
from typing import Optional, Tuple, List
from dataclasses import dataclass


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


class KoradKWR102:
    """Driver for Korad KWR102 Power Supply"""
    
    def __init__(self, port: str = "", baudrate: int = 115200, timeout: float = 1.0):
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.serial: Optional[serial.Serial] = None
        self._ocp_enabled = False
        self._ovp_enabled = False
        self._serial_lock = threading.Lock()  # Thread-safe serial access
        
    @staticmethod
    def list_ports() -> List[str]:
        """List available COM ports"""
        ports = serial.tools.list_ports.comports()
        return [p.device for p in ports]
    
    def connect(self, port: str = "") -> bool:
        """Connect to the power supply"""
        if port:
            self.port = port
        if not self.port:
            return False
        try:
            self.serial = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=self.timeout
            )
            # KWR102 needs RTS/DTR set
            self.serial.setRTS(True)
            self.serial.setDTR(True)
            time.sleep(0.2)
            self.serial.reset_input_buffer()
            self.serial.reset_output_buffer()
            return True
        except Exception as e:
            print(f"PSU connection error: {e}")
            return False
    
    def disconnect(self):
        """Disconnect from the power supply"""
        if self.serial and self.serial.is_open:
            try:
                self.serial.close()
            except:
                pass
        self.serial = None
            
    def is_connected(self) -> bool:
        """Check if connected"""
        return self.serial is not None and self.serial.is_open
    
    def _send_command(self, cmd: str, fast: bool = False) -> Optional[str]:
        """Send command and optionally read response (thread-safe)
        
        Args:
            cmd: Command to send
            fast: If True, uses minimal delay (for ramping)
        """
        if not self.is_connected():
            return None
        
        with self._serial_lock:  # Ensure thread-safe access
            try:
                self.serial.reset_input_buffer()
                # KWR102 requires \r (carriage return) terminator
                self.serial.write((cmd + '\r').encode('ascii'))
                
                if '?' in cmd:
                    # For queries, use shorter timeout to speed up logging
                    old_timeout = self.serial.timeout
                    self.serial.timeout = 0.1  # Fast timeout for queries
                    time.sleep(0.02)  # Minimal delay for queries
                    response = self.serial.read(100).decode('ascii').strip()
                    self.serial.timeout = old_timeout
                    return response
                else:
                    # For set commands
                    if fast:
                        time.sleep(0.01)  # Minimal delay for fast operations (ramp)
                    else:
                        time.sleep(0.1)  # Normal delay for regular operations
                return ""
            except Exception as e:
                print(f"PSU command error: {e}")
                return None
    
    def get_identification(self) -> str:
        """Get device identification string"""
        return self._send_command("*IDN?") or "Unknown"
    
    def set_voltage(self, voltage: float, fast: bool = False) -> bool:
        """Set output voltage (V)
        
        Args:
            voltage: Voltage to set
            fast: If True, uses minimal delay (for ramping)
        """
        voltage = max(0, min(voltage, 60))  # Clamp to safe range
        # KWR102 uses VSET: not VSET1:
        cmd = f"VSET:{voltage:05.2f}"
        return self._send_command(cmd, fast=fast) is not None
    
    def get_voltage_setpoint(self) -> float:
        """Get voltage setpoint (V)"""
        # KWR102 uses VSET? not VSET1?
        response = self._send_command("VSET?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def get_output_voltage(self) -> float:
        """Get actual output voltage (V)"""
        # KWR102 uses VOUT? not VOUT1?
        response = self._send_command("VOUT?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def set_current(self, current: float) -> bool:
        """Set current limit (A)"""
        current = max(0, min(current, 30))  # Clamp to safe range
        # KWR102 uses ISET: not ISET1:
        cmd = f"ISET:{current:05.3f}"
        return self._send_command(cmd) is not None
    
    def get_current_setpoint(self) -> float:
        """Get current setpoint (A)"""
        # KWR102 uses ISET? not ISET1?
        response = self._send_command("ISET?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def get_output_current(self) -> float:
        """Get actual output current (A)"""
        # KWR102 uses IOUT? not IOUT1?
        response = self._send_command("IOUT?")
        try:
            return float(response) if response else 0.0
        except ValueError:
            return 0.0
    
    def set_output(self, on: bool) -> bool:
        """Turn output on or off"""
        if on:
            # KWR102 uses OUT:1 (WITH colon) for ON - consistent with OUT:0 for OFF
            cmd = "OUT:1"
        else:
            # KWR102 uses OUT:0 (WITH colon) for OFF
            cmd = "OUT:0"
        
        result = self._send_command(cmd) is not None
        # Give PSU time to process output state change
        time.sleep(0.3)
        return result
    
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
        # KWR102 STATUS? sometimes hangs, check output voltage instead
        output_on = False
        mode = "CV"
        
        # Try STATUS? with short timeout
        try:
            old_timeout = self.serial.timeout
            self.serial.timeout = 0.3  # Short timeout
            status_response = self._send_command("STATUS?")
            self.serial.timeout = old_timeout
            
            if status_response:
                try:
                    status = ord(status_response[0]) if len(status_response) > 0 else 0
                    output_on = bool(status & 0x40)
                    mode = "CC" if (status & 0x01) else "CV"
                except:
                    pass
        except:
            pass
        
        # Fallback: check if output voltage > 0 to determine if output is on
        if not output_on:
            vout = self.get_output_voltage()
            if vout > 0.1:  # If we read voltage, output is likely on
                output_on = True
        
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
        """Execute a single voltage ramp with accurate timing"""
        if duration_s <= 0:
            self.psu.set_voltage(end_v)
            return
            
        steps = int(duration_s / step_interval)
        if steps < 1:
            steps = 1
            
        voltage_step = (end_v - start_v) / steps
        self.current_voltage = start_v
        
        # Debug logging
        import os
        ramp_log = os.path.join(os.path.dirname(__file__), "ramp_debug.txt")
        with open(ramp_log, "a") as f:
            f.write(f"\n=== RAMP START ===\n")
            f.write(f"Start: {start_v}V, End: {end_v}V, Duration: {duration_s}s\n")
            f.write(f"Steps: {steps}, Step interval: {step_interval}s, Voltage step: {voltage_step:.6f}V\n")
        
        # Start timing from first voltage set
        ramp_start_time = time.time()
        self.psu.set_voltage(self.current_voltage, fast=True)
        
        for i in range(steps + 1):
            if not self.running:
                break
                
            while self.paused and self.running:
                time.sleep(0.1)
                
            if i > 0:
                # Calculate target time for this step
                target_time = ramp_start_time + (i * step_interval)
                
                # Set next voltage with fast mode
                self.current_voltage = start_v + (voltage_step * i)
                self.current_voltage = round(self.current_voltage, 3)
                cmd_start = time.time()
                self.psu.set_voltage(self.current_voltage, fast=True)
                cmd_duration = time.time() - cmd_start
                
                # Wait until target time (accounts for command overhead)
                remaining = target_time - time.time()
                if remaining > 0.001:  # Only wait if >1ms remains
                    time.sleep(remaining)
            
            if self.progress_callback:
                progress_pct = (i / steps) * 100
                self.progress_callback(current_cycle, total_cycles, 
                                       self.current_voltage, progress_pct)
        
        # Debug: log completion
        actual_duration = time.time() - ramp_start_time
        with open(ramp_log, "a") as f:
            f.write(f"RAMP COMPLETE: Final voltage: {self.current_voltage}V, Actual duration: {actual_duration:.1f}s\n")
                
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
