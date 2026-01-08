"""
Fix button macro assignments to point to VoltAmpero module
"""

import xlwings as xw

def fix_button_assignments():
    print("Fixing button assignments...")
    
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    control_sheet = wb.sheets["Control"]
    
    # Map of expected button captions to macro names
    button_macros = {
        "Connect PSU": "VoltAmpero.ConnectPSU",
        "Connect DMM": "VoltAmpero.ConnectDMM",
        "Disconnect All": "VoltAmpero.DisconnectAll",
        "Test (Simulated)": "VoltAmpero.InitSimulated",
        "Output ON": "VoltAmpero.OutputOn",
        "Output On": "VoltAmpero.OutputOn",  # Handle case variation
        "Output OFF": "VoltAmpero.OutputOff",
        "Output Off": "VoltAmpero.OutputOff",  # Handle case variation
        "Apply Settings": "VoltAmpero.ApplySettings",
        "Start Logging": "VoltAmpero.StartLogging",
        "Stop Logging": "VoltAmpero.StopLogging",
        "Stop Loggin": "VoltAmpero.StopLogging",  # Handle typo
        "Start Ramp": "VoltAmpero.StartRamp",
        "Stop Ramp": "VoltAmpero.StopRamp",
        "Pause Ramp": "VoltAmpero.PauseRamp",
        "Export CSV": "VoltAmpero.ExportCSV",
        "Clear Data": "VoltAmpero.ClearData",
        "Clear Data ": "VoltAmpero.ClearData",  # Handle extra space
    }
    
    print("\n1. Current buttons on Control sheet:")
    fixed_count = 0
    
    try:
        # Access all shapes (buttons are shapes)
        for shape in control_sheet.api.Shapes:
            if shape.Type == 8:  # msoFormControl (button)
                try:
                    caption = shape.TextFrame.Characters().Text
                    old_macro = shape.OnAction
                    
                    print(f"\n   Button: '{caption}'")
                    print(f"   Current macro: {old_macro}")
                    
                    # Find matching macro
                    if caption in button_macros:
                        new_macro = button_macros[caption]
                        shape.OnAction = new_macro
                        print(f"   New macro: {new_macro} [FIXED]")
                        fixed_count += 1
                    else:
                        print(f"   [WARNING] No mapping found for this button")
                        
                except Exception as e:
                    print(f"   [ERROR] Failed to update button: {e}")
                    
    except Exception as e:
        print(f"[ERROR] Failed to access buttons: {e}")
        return False
    
    print(f"\n2. Fixed {fixed_count} button assignments")
    
    # Save
    print("\n3. Saving workbook...")
    wb.save()
    print("   [OK] Saved")
    
    print("\n" + "="*60)
    print(f"SUCCESS! Fixed {fixed_count} buttons.")
    print("Close and reopen Excel, then try the buttons again.")
    print("="*60)
    
    return True

if __name__ == "__main__":
    fix_button_assignments()
