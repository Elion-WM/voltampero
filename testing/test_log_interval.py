"""
Test log interval reading from Excel
Run with Excel open
"""

import xlwings as xw

print("=" * 60)
print("Test Log Interval")
print("=" * 60)

# Get Excel
wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("[FAIL] Excel not open")
    exit(1)

ctrl = wb.sheets["Control"]

print("\n1. Reading LogInterval from Excel...")
try:
    interval_value = ctrl.range("LogInterval").value
    print(f"   Cell B8 value: {interval_value}")
    print(f"   Type: {type(interval_value)}")
    
    interval_int = int(interval_value or 300)
    print(f"   Converted to int: {interval_int}ms")
    print(f"   That's {interval_int/1000} seconds between readings")
    
except Exception as e:
    print(f"   [FAIL] Error: {e}")
    import traceback
    traceback.print_exc()

# Test what va_start_logging would do
print("\n2. Simulating va_start_logging logic...")
from voltampero import get_controller
c = get_controller()
c.attach_excel(wb)

try:
    interval = 300
    interval = int(c.control_sheet.range("LogInterval").value or 300)
    print(f"   Retrieved interval: {interval}ms")
except Exception as e:
    print(f"   Error getting interval: {e}")
    print(f"   Using default: 300ms")

print("\n" + "=" * 60)
print(f"Expected: Data every {interval}ms")
print(f"Actual: You report data every 5000ms (5 seconds)")
print(f"Ratio: {5000/interval}x slower than expected")
print("=" * 60)
