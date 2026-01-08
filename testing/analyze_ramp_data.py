"""
Analyze the Data tab to see if current limiting occurred during ramp
"""
import xlwings as xw

# Find Excel
wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("Excel not open")
    exit(1)

data_sheet = wb.sheets["Data"]

# Get last row
last_row = data_sheet.range("A1").end('down').row
print(f"Total data rows: {last_row - 1}")

if last_row <= 1:
    print("No data collected")
    exit(1)

# Read voltage and current columns
print("\nAnalyzing data...")
voltages = data_sheet.range(f"C2:C{last_row}").value
currents = data_sheet.range(f"D2:D{last_row}").value
setpoint_v = data_sheet.range(f"E2:E{last_row}").value

# Handle single value case
if not isinstance(voltages, list):
    voltages = [voltages]
    currents = [currents]
    setpoint_v = [setpoint_v]

print(f"\nRamp analysis (first 10 and last 10 samples):")
print(f"{'Index':<8} {'Setpoint':<10} {'Actual V':<10} {'Current':<10} {'V Diff':<10} {'Status'}")
print("-" * 70)

for i in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, -10, -9, -8, -7, -6, -5, -4, -3, -2, -1]:
    if abs(i) < len(voltages):
        idx = i if i >= 0 else len(voltages) + i
        v_set = setpoint_v[idx] if setpoint_v[idx] is not None else 0
        v_actual = voltages[idx] if voltages[idx] is not None else 0
        i_actual = currents[idx] if currents[idx] is not None else 0
        v_diff = v_actual - v_set
        
        status = "OK (CV)"
        if abs(v_diff) > 0.5:  # More than 0.5V difference
            status = "WARNING: CC MODE (Current limiting!)"
        
        print(f"{idx:<8} {v_set:<10.2f} {v_actual:<10.2f} {i_actual:<10.3f} {v_diff:+10.2f} {status}")

# Statistics
voltage_errors = [abs(voltages[i] - setpoint_v[i]) for i in range(len(voltages)) 
                  if voltages[i] is not None and setpoint_v[i] is not None]
avg_error = sum(voltage_errors) / len(voltage_errors) if voltage_errors else 0
max_error = max(voltage_errors) if voltage_errors else 0

print("\n" + "=" * 70)
print("SUMMARY:")
print(f"  Average voltage error: {avg_error:.3f}V")
print(f"  Maximum voltage error: {max_error:.3f}V")

if max_error > 0.5:
    print(f"\nWARNING: PSU entered CURRENT LIMITING (CC) mode!")
    print(f"  The load drew too much current, preventing voltage from rising.")
    print(f"  Consider: Increasing current limit or reducing load.")
else:
    print(f"\nOK: PSU operated in VOLTAGE mode (CV) throughout ramp")

print("=" * 70)
