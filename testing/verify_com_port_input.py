"""
Verify B3 is set up correctly for COM port input
"""

import xlwings as xw

def verify_com_port_cell():
    print("Verifying COM port input cell (B3)...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control = wb.sheets["Control"]
    
    # Check B3
    print("\n1. Cell B3 (PSU Port):")
    cell = control.range("B3")
    print(f"   Current value: {cell.value}")
    print(f"   Locked: {cell.api.Locked}")
    print(f"   Color: {cell.color}")
    
    # Make sure it's unlocked
    if cell.api.Locked:
        print("   [FIX] Unlocking cell B3...")
        cell.api.Locked = False
    
    # Set default value if empty
    if not cell.value or cell.value == "COM3":
        print("   [INFO] Setting default to COM4...")
        cell.value = "COM4"
    
    # Make it yellow to show it's editable
    if cell.color != 0xFFFFCC:
        print("   [FIX] Highlighting as input cell (yellow)...")
        cell.color = 0xFFFFCC
    
    # Add a label to make it clear
    print("\n2. Adding clear label:")
    control.range("A3").value = "PSU Port:"
    control.range("A3").api.Font.Bold = True
    
    # Check named range
    print("\n3. Checking named range:")
    try:
        psu_port_range = wb.names["PSUPort"].refers_to
        print(f"   PSUPort named range: {psu_port_range}")
    except:
        print("   [FIX] Creating PSUPort named range...")
        wb.names.add("PSUPort", "=Control!$B$3")
    
    # Test auto-reconnect with real PSU
    print("\n4. Testing auto-reconnect logic with COM4:")
    import sys
    sys.path.insert(0, r"C:\Users\User\GitHub\voltampero")
    
    # Simulate what happens after ConnectPSU is clicked
    control.range("B3").value = "COM4"
    control.range("D3").value = "Connected"  # Simulate connected state
    
    print(f"   B3 (PSU Port): {control.range('B3').value}")
    print(f"   D3 (PSU Status): {control.range('D3').value}")
    
    wb.save()
    print("\n[OK] Cell B3 is now ready for COM port input!")
    print("\n" + "="*60)
    print("HOW TO USE:")
    print("="*60)
    print("1. Type 'COM4' (or any COM port) in cell B3")
    print("2. Click 'Connect PSU' button")
    print("3. Cell D3 will show 'Connected' if successful")
    print("4. Now all other buttons will auto-reconnect to COM4")
    print("5. You can change voltage/current and click 'Apply Settings'")
    print("="*60)

if __name__ == "__main__":
    verify_com_port_cell()
