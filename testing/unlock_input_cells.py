"""
Unlock all input cells in Control sheet
"""

import xlwings as xw

def unlock_input_cells():
    print("Unlocking input cells...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control_sheet = wb.sheets["Control"]
    
    # List of input cells that users should be able to edit
    input_cells = [
        "B3",   # PSU Port
        "B8",   # Log Interval
        "B16",  # Set Voltage
        "B17",  # Set Current
        "B18",  # OCP Enabled
        "B21",  # Ramp Start V
        "B22",  # Ramp End V
        "B23",  # Ramp Duration
        "B24",  # Ramp Cycles
        "B25",  # Ramp Delay
        "B26",  # Ramp Ping-Pong
    ]
    
    print("\nUnlocking cells:")
    for cell_addr in input_cells:
        cell = control_sheet.range(cell_addr)
        cell.api.Locked = False
        print(f"   {cell_addr}: Unlocked")
    
    # Also make sure these cells have proper formatting
    print("\nSetting number format for numeric cells:")
    try:
        control_sheet.range("B16").api.NumberFormat = "0.00"  # Voltage
        control_sheet.range("B17").api.NumberFormat = "0.000" # Current
        control_sheet.range("B8").api.NumberFormat = "0"      # Log interval
        control_sheet.range("B21:B23").api.NumberFormat = "0.00" # Ramp values
        control_sheet.range("B24:B25").api.NumberFormat = "0"    # Cycles and delay
        print("   [OK] Number formats set")
    except Exception as e:
        print(f"   [WARNING] Could not set number formats: {e}")
        print("   (This is OK, cells are still unlocked)")
    
    # Set default values if needed
    print("\nSetting default values:")
    if control_sheet.range("B16").value is None:
        control_sheet.range("B16").value = 5.0
        print("   B16 (Voltage): Set to 5.0")
    
    if control_sheet.range("B17").value is None:
        control_sheet.range("B17").value = 1.0
        print("   B17 (Current): Set to 1.0")
    
    # Make input cells have a light yellow background to indicate they're editable
    print("\nHighlighting input cells:")
    input_color = 0xFFFFCC  # Light yellow (RGB: 255, 255, 204)
    for cell_addr in input_cells:
        if cell_addr != "B18":  # Skip checkbox
            control_sheet.range(cell_addr).color = input_color
    
    wb.save()
    print("\n[OK] All input cells unlocked and formatted!")
    print("\nYou can now edit voltage and current values in Excel.")

if __name__ == "__main__":
    unlock_input_cells()
