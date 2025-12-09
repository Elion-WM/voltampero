"""
VoltAmpero - Lab Instrument Control Software
Main controller with Excel/xlwings integration
Controls Korad KWR102 PSU and UNI-T UT8804E Multimeter
"""

import time
import threading
import csv
import os
from datetime import datetime
from typing import Optional, List, Tuple
from dataclasses import dataclass

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    print("xlwings not installed. Install with: pip install xlwings")

from psu_korad import KoradKWR102, VoltageRamp, SimulatedPSU
from multimeter_unit import UNIT_UT8804E, SimulatedMultimeter, MultimeterReading


@dataclass
class LogEntry:
    timestamp: datetime
    elapsed_s: float
    psu_voltage: float
    psu_current: float
    psu_setpoint_v: float
    psu_setpoint_a: float
    dmm_value: float
    dmm_unit: str
    dmm_mode: str


class VoltAmpero:
    """Main controller class for VoltAmpero system"""
    
    def __init__(self, simulate: bool = False):
        self.simulate = simulate
        
        # Initialize devices
        if simulate:
            self.psu = SimulatedPSU()
            self.dmm = SimulatedMultimeter()
        else:
            self.psu = KoradKWR102()
            self.dmm = UNIT_UT8804E()
            
        self.voltage_ramp = VoltageRamp(self.psu)
        
        # Logging state
        self.logging_active = False
        self.log_data: List[LogEntry] = []
        self.log_start_time: Optional[datetime] = None
        self.log_interval_ms = 300
        self._log_thread: Optional[threading.Thread] = None
        self._stop_logging = threading.Event()
        
        # Excel integration
        self.wb: Optional[xw.Book] = None
        self.data_sheet = None
        self.control_sheet = None
        self._excel_update_row = 2
        
        # Ramp state
        self._ramp_thread: Optional[threading.Thread] = None
        
    # ========== Connection Methods ==========
    
    def list_com_ports(self) -> List[str]:
        """List available COM ports"""
        if self.simulate:
            return ["SIM1", "SIM2"]
        return KoradKWR102.list_ports()
    
    def connect_psu(self, port: str) -> bool:
        """Connect to power supply"""
        result = self.psu.connect(port)
        self._update_excel_status("PSU", "Connected" if result else "Disconnected")
        return result
    
    def disconnect_psu(self):
        """Disconnect from power supply"""
        self.psu.disconnect()
        self._update_excel_status("PSU", "Disconnected")
    
    def connect_dmm(self) -> bool:
        """Connect to multimeter"""
        result = self.dmm.connect()
        self._update_excel_status("DMM", "Connected" if result else "Disconnected")
        return result
    
    def disconnect_dmm(self):
        """Disconnect from multimeter"""
        self.dmm.disconnect()
        self._update_excel_status("DMM", "Disconnected")
        
    def connect_all(self, psu_port: str) -> Tuple[bool, bool]:
        """Connect to both devices"""
        psu_ok = self.connect_psu(psu_port)
        dmm_ok = self.connect_dmm()
        return (psu_ok, dmm_ok)
    
    def disconnect_all(self):
        """Disconnect from all devices"""
        self.stop_logging()
        self.stop_ramp()
        self.disconnect_psu()
        self.disconnect_dmm()
        
    # ========== PSU Control Methods ==========
    
    def set_voltage(self, voltage: float) -> bool:
        """Set PSU output voltage"""
        return self.psu.set_voltage(voltage)
    
    def set_current(self, current: float) -> bool:
        """Set PSU current limit"""
        return self.psu.set_current(current)
    
    def output_on(self) -> bool:
        """Turn PSU output on"""
        return self.psu.output_on()
    
    def output_off(self) -> bool:
        """Turn PSU output off"""
        return self.psu.output_off()
    
    def set_ocp(self, enabled: bool) -> bool:
        """Enable/disable Over Current Protection"""
        return self.psu.set_ocp(enabled)
    
    def set_ovp(self, enabled: bool) -> bool:
        """Enable/disable Over Voltage Protection"""
        return self.psu.set_ovp(enabled)
    
    def get_psu_readings(self) -> Tuple[float, float]:
        """Get PSU voltage and current"""
        return self.psu.get_readings()
    
    def get_psu_status(self):
        """Get full PSU status"""
        return self.psu.get_status()
    
    # ========== Voltage Ramp Methods ==========
    
    def start_ramp(self, start_v: float, end_v: float, duration_s: float,
                   cycles: int = 1, delay_between_s: float = 0.0,
                   ping_pong: bool = False):
        """Start voltage ramp in background thread"""
        if self._ramp_thread and self._ramp_thread.is_alive():
            self.stop_ramp()
            time.sleep(0.2)
            
        def ramp_with_callback():
            self.voltage_ramp.start(
                start_v=start_v,
                end_v=end_v,
                duration_s=duration_s,
                cycles=cycles,
                delay_between_s=delay_between_s,
                ping_pong=ping_pong,
                progress_callback=self._ramp_progress_callback
            )
            self._update_excel_status("Ramp", "Stopped")
            
        self._update_excel_status("Ramp", "Running")
        self._ramp_thread = threading.Thread(target=ramp_with_callback, daemon=True)
        self._ramp_thread.start()
    
    def stop_ramp(self):
        """Stop voltage ramp"""
        self.voltage_ramp.stop()
        self._update_excel_status("Ramp", "Stopped")
    
    def pause_ramp(self):
        """Pause voltage ramp"""
        self.voltage_ramp.pause()
        self._update_excel_status("Ramp", "Paused")
    
    def resume_ramp(self):
        """Resume voltage ramp"""
        self.voltage_ramp.resume()
        self._update_excel_status("Ramp", "Running")
        
    def is_ramp_running(self) -> bool:
        """Check if ramp is active"""
        return self.voltage_ramp.is_running()
    
    def _ramp_progress_callback(self, cycle: int, total_cycles: int, 
                                 voltage: float, progress_pct: float):
        """Called during ramp to update Excel"""
        if self.control_sheet:
            try:
                self.control_sheet.range("RampCycle").value = f"{cycle}/{total_cycles if total_cycles > 0 else '∞'}"
                self.control_sheet.range("RampVoltage").value = voltage
                self.control_sheet.range("RampProgress").value = progress_pct / 100
            except:
                pass
                
    # ========== DMM Methods ==========
    
    def get_dmm_reading(self) -> Optional[MultimeterReading]:
        """Get current multimeter reading"""
        return self.dmm.get_reading()
    
    def get_dmm_value(self) -> float:
        """Get multimeter numeric value"""
        return self.dmm.get_value()
    
    def get_dmm_display(self) -> str:
        """Get formatted multimeter display"""
        return self.dmm.get_value_with_unit()
        
    # ========== Logging Methods ==========
    
    def start_logging(self, interval_ms: int = 300):
        """Start data logging"""
        if self.logging_active:
            return
            
        self.log_interval_ms = interval_ms
        self.log_data = []
        self.log_start_time = datetime.now()
        self._stop_logging.clear()
        self.logging_active = True
        self._excel_update_row = 2
        
        # Clear previous data in Excel
        if self.data_sheet:
            try:
                last_row = self.data_sheet.range("A1").end('down').row
                if last_row > 1:
                    self.data_sheet.range(f"A2:I{last_row}").clear_contents()
            except:
                pass
        
        self._log_thread = threading.Thread(target=self._logging_loop, daemon=True)
        self._log_thread.start()
        self._update_excel_status("Logging", "Active")
    
    def stop_logging(self):
        """Stop data logging"""
        self._stop_logging.set()
        self.logging_active = False
        if self._log_thread:
            self._log_thread.join(timeout=1.0)
        self._update_excel_status("Logging", "Stopped")
    
    def _logging_loop(self):
        """Background logging loop"""
        while not self._stop_logging.is_set():
            try:
                entry = self._capture_reading()
                if entry:
                    self.log_data.append(entry)
                    self._write_entry_to_excel(entry)
            except Exception as e:
                print(f"Logging error: {e}")
                
            # Wait for next interval
            self._stop_logging.wait(self.log_interval_ms / 1000.0)
    
    def _capture_reading(self) -> Optional[LogEntry]:
        """Capture a single reading from both devices"""
        now = datetime.now()
        elapsed = (now - self.log_start_time).total_seconds() if self.log_start_time else 0
        
        # Read PSU
        psu_v, psu_a = 0.0, 0.0
        psu_set_v, psu_set_a = 0.0, 0.0
        if self.psu.is_connected():
            psu_v, psu_a = self.psu.get_readings()
            psu_set_v = self.psu.get_voltage_setpoint()
            psu_set_a = self.psu.get_current_setpoint()
        
        # Read DMM
        dmm_val, dmm_unit, dmm_mode = 0.0, "", ""
        if self.dmm.is_connected():
            reading = self.dmm.get_reading()
            if reading:
                dmm_val = reading.value
                dmm_unit = reading.unit
                dmm_mode = reading.mode.value
        
        return LogEntry(
            timestamp=now,
            elapsed_s=elapsed,
            psu_voltage=psu_v,
            psu_current=psu_a,
            psu_setpoint_v=psu_set_v,
            psu_setpoint_a=psu_set_a,
            dmm_value=dmm_val,
            dmm_unit=dmm_unit,
            dmm_mode=dmm_mode
        )
    
    def _write_entry_to_excel(self, entry: LogEntry):
        """Write a log entry to Excel"""
        if not self.data_sheet:
            return
        try:
            row = self._excel_update_row
            self.data_sheet.range(f"A{row}").value = [
                entry.timestamp.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
                entry.elapsed_s,
                entry.psu_voltage,
                entry.psu_current,
                entry.psu_setpoint_v,
                entry.psu_setpoint_a,
                entry.dmm_value,
                entry.dmm_unit,
                entry.dmm_mode
            ]
            self._excel_update_row += 1
            
            # Update live display cells
            if self.control_sheet:
                self.control_sheet.range("LiveVoltage").value = entry.psu_voltage
                self.control_sheet.range("LiveCurrent").value = entry.psu_current
                self.control_sheet.range("LiveDMM").value = f"{entry.dmm_value:.4f} {entry.dmm_unit}"
        except Exception as e:
            print(f"Excel write error: {e}")
    
    def export_csv(self, filepath: str) -> bool:
        """Export log data to CSV file"""
        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([
                    "Timestamp", "Elapsed_s", "PSU_Voltage_V", "PSU_Current_A",
                    "PSU_Setpoint_V", "PSU_Setpoint_A", "DMM_Value", "DMM_Unit", "DMM_Mode"
                ])
                for entry in self.log_data:
                    writer.writerow([
                        entry.timestamp.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
                        f"{entry.elapsed_s:.3f}",
                        f"{entry.psu_voltage:.4f}",
                        f"{entry.psu_current:.4f}",
                        f"{entry.psu_setpoint_v:.2f}",
                        f"{entry.psu_setpoint_a:.3f}",
                        f"{entry.dmm_value:.6f}",
                        entry.dmm_unit,
                        entry.dmm_mode
                    ])
            return True
        except Exception as e:
            print(f"CSV export error: {e}")
            return False
    
    def clear_log(self):
        """Clear log data"""
        self.log_data = []
        self._excel_update_row = 2
        if self.data_sheet:
            try:
                last_row = self.data_sheet.range("A1").end('down').row
                if last_row > 1:
                    self.data_sheet.range(f"A2:I{last_row}").clear_contents()
            except:
                pass
                
    # ========== Excel Integration ==========
    
    def attach_excel(self, workbook: xw.Book = None):
        """Attach to Excel workbook"""
        if not XLWINGS_AVAILABLE:
            print("xlwings not available")
            return False
            
        try:
            if workbook:
                self.wb = workbook
            else:
                self.wb = xw.Book.caller()
                
            # Get sheets
            self.control_sheet = self.wb.sheets["Control"]
            self.data_sheet = self.wb.sheets["Data"]
            return True
        except Exception as e:
            print(f"Excel attach error: {e}")
            return False
    
    def _update_excel_status(self, component: str, status: str):
        """Update status indicator in Excel"""
        if self.control_sheet:
            try:
                self.control_sheet.range(f"{component}Status").value = status
            except:
                pass


# ========== Global Instance ==========
_controller: Optional[VoltAmpero] = None

def get_controller(simulate: bool = False) -> VoltAmpero:
    """Get or create the global controller instance"""
    global _controller
    if _controller is None:
        _controller = VoltAmpero(simulate=simulate)
    return _controller


# ========== xlwings UDF Functions (callable from Excel) ==========

if XLWINGS_AVAILABLE:
    
    @xw.func
    def va_list_ports() -> str:
        """List available COM ports"""
        ctrl = get_controller()
        ports = ctrl.list_com_ports()
        return ", ".join(ports) if ports else "No ports found"
    
    @xw.func
    def va_connect_psu(port: str) -> str:
        """Connect to PSU"""
        ctrl = get_controller()
        if ctrl.connect_psu(port):
            return f"Connected to {port}"
        return "Connection failed"
    
    @xw.func
    def va_connect_dmm() -> str:
        """Connect to multimeter"""
        ctrl = get_controller()
        if ctrl.connect_dmm():
            return "DMM Connected"
        return "DMM Connection failed"
    
    @xw.func
    def va_get_voltage() -> float:
        """Get PSU output voltage"""
        ctrl = get_controller()
        v, _ = ctrl.get_psu_readings()
        return v
    
    @xw.func
    def va_get_current() -> float:
        """Get PSU output current"""
        ctrl = get_controller()
        _, a = ctrl.get_psu_readings()
        return a
    
    @xw.func
    def va_get_dmm() -> str:
        """Get DMM reading with unit"""
        ctrl = get_controller()
        return ctrl.get_dmm_display()
    
    @xw.sub
    def va_set_voltage(voltage: float):
        """Set PSU voltage"""
        ctrl = get_controller()
        ctrl.set_voltage(voltage)
    
    @xw.sub
    def va_set_current(current: float):
        """Set PSU current limit"""
        ctrl = get_controller()
        ctrl.set_current(current)
    
    @xw.sub
    def va_output_on():
        """Turn PSU output on"""
        ctrl = get_controller()
        ctrl.output_on()
    
    @xw.sub
    def va_output_off():
        """Turn PSU output off"""
        ctrl = get_controller()
        ctrl.output_off()
    
    @xw.sub
    def va_set_ocp(enabled: bool):
        """Enable/disable OCP"""
        ctrl = get_controller()
        ctrl.set_ocp(enabled)
    
    @xw.sub
    def va_start_logging():
        """Start data logging"""
        ctrl = get_controller()
        ctrl.attach_excel()
        interval = 300
        try:
            interval = int(ctrl.control_sheet.range("LogInterval").value or 300)
        except:
            pass
        ctrl.start_logging(interval)
    
    @xw.sub
    def va_stop_logging():
        """Stop data logging"""
        ctrl = get_controller()
        ctrl.stop_logging()
    
    @xw.sub
    def va_export_csv():
        """Export data to CSV"""
        ctrl = get_controller()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join(os.path.dirname(ctrl.wb.fullname), f"voltampero_log_{timestamp}.csv")
        if ctrl.export_csv(filepath):
            ctrl.control_sheet.range("ExportStatus").value = f"Exported: {filepath}"
        else:
            ctrl.control_sheet.range("ExportStatus").value = "Export failed"
    
    @xw.sub
    def va_clear_data():
        """Clear log data"""
        ctrl = get_controller()
        ctrl.clear_log()
    
    @xw.sub
    def va_start_ramp():
        """Start voltage ramp from Excel settings"""
        ctrl = get_controller()
        ctrl.attach_excel()
        try:
            start_v = float(ctrl.control_sheet.range("RampStartV").value or 0)
            end_v = float(ctrl.control_sheet.range("RampEndV").value or 0)
            duration = float(ctrl.control_sheet.range("RampDuration").value or 10)
            cycles = int(ctrl.control_sheet.range("RampCycles").value or 1)
            delay = float(ctrl.control_sheet.range("RampDelay").value or 0)
            ping_pong = bool(ctrl.control_sheet.range("RampPingPong").value)
            
            ctrl.start_ramp(start_v, end_v, duration, cycles, delay, ping_pong)
        except Exception as e:
            print(f"Ramp start error: {e}")
    
    @xw.sub
    def va_stop_ramp():
        """Stop voltage ramp"""
        ctrl = get_controller()
        ctrl.stop_ramp()
    
    @xw.sub
    def va_pause_ramp():
        """Pause voltage ramp"""
        ctrl = get_controller()
        ctrl.pause_ramp()
    
    @xw.sub
    def va_resume_ramp():
        """Resume voltage ramp"""
        ctrl = get_controller()
        ctrl.resume_ramp()
    
    @xw.sub
    def va_disconnect_all():
        """Disconnect all devices"""
        ctrl = get_controller()
        ctrl.disconnect_all()
    
    @xw.sub
    def va_init_simulated():
        """Initialize with simulated devices for testing"""
        global _controller
        _controller = VoltAmpero(simulate=True)
        _controller.attach_excel()
        _controller.connect_psu("SIM1")
        _controller.connect_dmm()


# ========== Standalone Mode ==========

def main():
    """Run in standalone mode (without Excel)"""
    print("VoltAmpero - Lab Instrument Control")
    print("=" * 40)
    
    # Try simulated mode for testing
    ctrl = VoltAmpero(simulate=True)
    
    print("\nAvailable COM ports:", ctrl.list_com_ports())
    
    print("\nConnecting to simulated devices...")
    ctrl.connect_psu("SIM1")
    ctrl.connect_dmm()
    
    print("\nPSU Status:", ctrl.get_psu_status())
    
    ctrl.set_voltage(5.0)
    ctrl.set_current(1.0)
    ctrl.output_on()
    
    print("\nStarting 5-second test log...")
    ctrl.start_logging(500)
    time.sleep(5)
    ctrl.stop_logging()
    
    print(f"\nCaptured {len(ctrl.log_data)} readings")
    
    # Export
    ctrl.export_csv("test_log.csv")
    print("Exported to test_log.csv")
    
    ctrl.disconnect_all()
    print("\nDone!")


if __name__ == "__main__":
    main()
