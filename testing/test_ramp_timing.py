"""
Test ramp timing directly
"""
import time
import xlwings as xw

wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("Excel not open")
    exit(1)

from voltampero import get_controller

ctrl = get_controller()
ctrl.attach_excel(wb)

print("Testing ramp timing...")
print("Ramp: 12V → 20V in 60 seconds")
print("Expected: 8V change in 60s = 0.133V/s")

# Check ramp settings
start_v = 12.0
end_v = 20.0
duration = 60.0

# Calculate expected steps
step_interval = 0.1  # default
steps = int(duration / step_interval)
voltage_step = (end_v - start_v) / steps

print(f"\nRamp calculation:")
print(f"  Duration: {duration}s")
print(f"  Step interval: {step_interval}s")
print(f"  Total steps: {steps}")
print(f"  Voltage step: {voltage_step:.6f}V")
print(f"  Expected rate: {(end_v-start_v)/duration:.4f}V/s")

# Simulate timing
print(f"\nAt 60 seconds (step {steps}):")
print(f"  Expected voltage: {start_v + (voltage_step * steps):.2f}V (should be {end_v}V)")

# Check if ramp is actually running
if ctrl.voltage_ramp.running:
    print(f"\n[INFO] Ramp is RUNNING")
    print(f"  Current voltage: {ctrl.voltage_ramp.current_voltage:.2f}V")
    print(f"  Current cycle: {ctrl.voltage_ramp.current_cycle}")
else:
    print(f"\n[INFO] Ramp is STOPPED")

# Get actual PSU voltage
psu_v, psu_a = ctrl.psu.get_readings()
print(f"\nActual PSU reading: {psu_v:.2f}V, {psu_a:.3f}A")
