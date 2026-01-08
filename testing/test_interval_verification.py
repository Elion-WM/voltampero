"""
Automated test to verify logging intervals
Tests 800ms and 1200ms settings
"""

import xlwings as xw
import time
import sys

print("=" * 70)
print("AUTOMATED INTERVAL VERIFICATION TEST")
print("=" * 70)

# Find Excel
wb = None
for book in xw.books:
    if "VoltAmpero" in book.name:
        wb = book
        break

if not wb:
    print("\n[FAIL] Excel not open with VoltAmpero.xlsm")
    print("Please open Excel first!")
    sys.exit(1)

print("\n[OK] Found Excel workbook")

ctrl = wb.sheets["Control"]

# Check if PSU is connected
psu_status = ctrl.range("PSUStatus").value
print(f"PSU Status: {psu_status}")

if psu_status != "Connected":
    print("\n[FAIL] PSU not connected")
    print("Please connect PSU first!")
    sys.exit(1)

print("[OK] PSU is connected")

# Test intervals
test_intervals = [800, 1200]

for interval in test_intervals:
    print("\n" + "=" * 70)
    print(f"Testing {interval}ms interval")
    print("=" * 70)
    
    # Set interval
    print(f"\n1. Setting LogInterval to {interval}ms...")
    ctrl.range("LogInterval").value = interval
    time.sleep(0.5)
    
    # Start logging
    print("2. Starting logging...")
    xw.apps.active.macro("VoltAmpero.StartLogging")()
    
    # Wait for data collection
    print(f"3. Collecting data for 12 seconds...")
    time.sleep(12)
    
    # Stop logging
    print("4. Stopping logging...")
    xw.apps.active.macro("VoltAmpero.StopLogging")()
    time.sleep(1)
    
    print(f"[OK] Test complete for {interval}ms")

print("\n" + "=" * 70)
print("ALL TESTS COMPLETE")
print("=" * 70)
print("\nChecking results from timing_debug.txt...")

try:
    with open(r"C:\Users\User\GitHub\voltampero\timing_debug.txt", "r") as f:
        lines = f.readlines()
    
    print("\n" + "=" * 70)
    print("RESULTS ANALYSIS")
    print("=" * 70)
    
    current_setting = None
    intervals_800 = []
    intervals_1200 = []
    
    for line in lines:
        if "setting=800ms" in line:
            current_setting = 800
            # Parse total time
            parts = line.split(",")
            total_str = parts[0].split("total=")[1].strip("s")
            intervals_800.append(float(total_str))
        elif "setting=1200ms" in line:
            current_setting = 1200
            parts = line.split(",")
            total_str = parts[0].split("total=")[1].strip("s")
            intervals_1200.append(float(total_str))
    
    if intervals_800:
        avg_800 = sum(intervals_800[1:]) / len(intervals_800[1:]) if len(intervals_800) > 1 else intervals_800[0]
        print(f"\n800ms setting:")
        print(f"  Samples: {len(intervals_800)}")
        print(f"  Average actual interval: {avg_800:.3f}s ({avg_800*1000:.0f}ms)")
        print(f"  Target: 800ms")
        print(f"  Accuracy: {(800/(avg_800*1000))*100:.1f}%")
        deviation = abs(avg_800 * 1000 - 800)
        if deviation < 100:
            print(f"  Status: ✓ EXCELLENT (within 100ms)")
        elif deviation < 200:
            print(f"  Status: ✓ GOOD (within 200ms)")
        else:
            print(f"  Status: ⚠ FAIR (deviation: {deviation:.0f}ms)")
    
    if intervals_1200:
        avg_1200 = sum(intervals_1200[1:]) / len(intervals_1200[1:]) if len(intervals_1200) > 1 else intervals_1200[0]
        print(f"\n1200ms setting:")
        print(f"  Samples: {len(intervals_1200)}")
        print(f"  Average actual interval: {avg_1200:.3f}s ({avg_1200*1000:.0f}ms)")
        print(f"  Target: 1200ms")
        print(f"  Accuracy: {(1200/(avg_1200*1000))*100:.1f}%")
        deviation = abs(avg_1200 * 1000 - 1200)
        if deviation < 100:
            print(f"  Status: ✓ EXCELLENT (within 100ms)")
        elif deviation < 200:
            print(f"  Status: ✓ GOOD (within 200ms)")
        else:
            print(f"  Status: ⚠ FAIR (deviation: {deviation:.0f}ms)")
    
    print("\n" + "=" * 70)
    print("VERDICT:")
    if intervals_800 and intervals_1200:
        avg_800_ms = avg_800 * 1000
        avg_1200_ms = avg_1200 * 1000
        if abs(avg_800_ms - 800) < 150 and abs(avg_1200_ms - 1200) < 150:
            print("✓ PASSED - Intervals are accurate and stable")
        else:
            print("⚠ NEEDS IMPROVEMENT - Some deviation from target")
    print("=" * 70)
    
except FileNotFoundError:
    print("\n[WARN] timing_debug.txt not found")
    print("Check if logging actually ran")
except Exception as e:
    print(f"\n[ERROR] Failed to analyze results: {e}")

print("\nDone!")
