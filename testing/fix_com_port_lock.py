"""
Fix COM port lock issue
"""

import xlwings as xw

print("="*60)
print("Fixing COM4 Port Lock Issue")
print("="*60)

print("\nThe issue: COM4 is locked (Access Denied)")
print("This happens when:")
print("1. Excel/Python is still holding the port from previous connection")
print("2. Another program is using COM4")
print("3. The PSU wasn't properly disconnected")

print("\n" + "="*60)
print("SOLUTION: Update Excel to show 'Disconnected'")
print("="*60)

wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
control = wb.sheets["Control"]

# Set status to disconnected so auto-reconnect won't try
control.range("D3").value = "Disconnected"
control.range("D4").value = "Disconnected"

wb.save()

print("\n[OK] Excel status set to 'Disconnected'")

print("\n" + "="*60)
print("STEPS TO FIX:")
print("="*60)
print("1. CLOSE EXCEL COMPLETELY")
print("   - File > Close")
print("   - Make sure Excel.exe is not running (check Task Manager)")
print("")
print("2. Check if other programs are using COM4:")
print("   - Device Manager > Ports (COM & LPT)")
print("   - Right-click COM4 > Properties")
print("   - If locked, reboot the PSU or unplug/replug USB cable")
print("")
print("3. Restart Excel:")
print("   - Double-click: start_excel_with_python.bat")
print("")
print("4. Connect to PSU:")
print("   - Make sure B3 = COM4")
print("   - Click 'Connect PSU'")
print("   - D3 should change to 'Connected'")
print("")
print("5. Now test Apply Settings:")
print("   - Set voltage in B16")
print("   - Set current in B17")
print("   - Click 'Apply Settings'")
print("   - Click 'Output ON'")
print("="*60)

# Also let's check if we need to add better cleanup
print("\n" + "="*60)
print("IMPROVING DISCONNECT HANDLING")
print("="*60)

# I should improve the DisconnectAll to make sure it really closes the port
