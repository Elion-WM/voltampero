"""
Fix Named Ranges in VoltAmpero.xlsm
This script recreates all named ranges with correct cell references
"""

import xlwings as xw

def fix_named_ranges():
    """Recreate all named ranges with correct references"""
    
    # Open the workbook
    wb = xw.Book(r"C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm")
    
    print("Deleting old named ranges...")
    # Delete existing names (except Excel built-in names starting with _xl)
    names_to_delete = []
    for name in wb.names:
        if not name.name.startswith("_xl"):
            names_to_delete.append(name.name)
    
    for name in names_to_delete:
        try:
            wb.names[name].delete()
            print(f"  Deleted: {name}")
        except Exception as e:
            print(f"  Could not delete {name}: {e}")
    
    print("\nCreating new named ranges...")
    # Create Control sheet named ranges
    named_ranges = {
        "PSUPort": "Control!$B$3",
        "PSUStatus": "Control!$D$3",
        "DMMStatus": "Control!$D$4",
        "LoggingStatus": "Control!$D$5",
        "RampStatus": "Control!$D$6",
        "LogInterval": "Control!$B$8",
        "LiveVoltage": "Control!$B$11",
        "LiveCurrent": "Control!$B$12",
        "LiveDMM": "Control!$B$13",
        "SetVoltage": "Control!$B$16",
        "SetCurrent": "Control!$B$17",
        "OCPEnabled": "Control!$B$18",
        "RampStartV": "Control!$B$21",
        "RampEndV": "Control!$B$22",
        "RampDuration": "Control!$B$23",
        "RampCycles": "Control!$B$24",
        "RampDelay": "Control!$B$25",
        "RampPingPong": "Control!$B$26",
        "RampCycle": "Control!$D$21",
        "RampVoltage": "Control!$D$22",
        "RampProgress": "Control!$D$23",
        "ExportStatus": "Control!$B$30",
    }
    
    for name, ref in named_ranges.items():
        try:
            wb.names.add(name, f"={ref}")
            print(f"  Created: {name} = {ref}")
        except Exception as e:
            print(f"  Error creating {name}: {e}")
    
    print("\nSaving workbook...")
    wb.save()
    print("Done! Named ranges have been fixed.")
    print("\nYou can verify by going to Formulas > Name Manager in Excel")
    
    # Keep workbook open for user to verify
    # wb.close()

if __name__ == "__main__":
    try:
        fix_named_ranges()
    except Exception as e:
        print(f"Error: {e}")
        print("\nMake sure:")
        print("1. VoltAmpero.xlsm is closed before running this script")
        print("2. xlwings is installed: pip install xlwings")
