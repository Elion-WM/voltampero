# Fix for Apply Settings Not Working

## Problem
The "Apply Settings" button in Excel was not applying voltage/current settings to the PSU because the VBA macro wasn't properly attaching to the Excel workbook before trying to communicate with Python.

## What Was Fixed

### 1. VBA Code (VoltAmpero.bas)
- Added `c.attach_excel()` call in the `ApplySettings` macro
- This ensures the Python controller can read Excel cells and reconnect to the PSU if needed

### 2. Python Code (voltampero.py)
- Added auto-reconnect logic to `output_on()`, `output_off()`, and `set_ocp()` methods
- Now these methods will automatically try to reconnect to the PSU if disconnected
- Matches the behavior already present in `set_voltage()` and `set_current()`

## How to Update Your Excel File

### Option 1: Update the VBA Code Manually (Quick)
1. Open `VoltAmpero.xlsm` in Excel
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer (left panel), find `VBAProject (VoltAmpero.xlsm)` → `Modules` → `VoltAmpero`
4. Double-click on `VoltAmpero` to open the code
5. Find the `ApplySettings` subroutine (around line 38)
6. Change this line:
   ```vb
   RunPython "from voltampero import get_controller; c=get_controller(); c.set_voltage(" & Replace(voltage, ",", ".") & "); c.set_current(" & Replace(current, ",", ".") & "); c.set_ocp(" & IIf(ocp, "True", "False") & ")"
   ```
   To:
   ```vb
   RunPython "from voltampero import get_controller; c=get_controller(); c.attach_excel(); c.set_voltage(" & Replace(voltage, ",", ".") & "); c.set_current(" & Replace(current, ",", ".") & "); c.set_ocp(" & IIf(ocp, "True", "False") & ")"
   ```
   (Just add `c.attach_excel();` after `c=get_controller();`)
7. Save the file (`Ctrl + S`)
8. Close VBA Editor and return to Excel

### Option 2: Re-import the Entire VBA Module (Complete)
1. Open `VoltAmpero.xlsm` in Excel
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer, right-click on the `VoltAmpero` module
4. Select `Remove VoltAmpero...` → Choose `No` when asked to export (we already have the updated file)
5. Right-click on `VBAProject (VoltAmpero.xlsm)` → `Import File...`
6. Browse to `VoltAmpero.bas` (the updated file in this directory)
7. Click `Open`
8. Save the workbook (`Ctrl + S`)
9. Close VBA Editor

## Testing the Fix

### Test with Simulated Hardware (Recommended First)
1. Open `VoltAmpero.xlsm`
2. Click the **"Test (Simulated)"** button to initialize simulated devices
3. Enter test values in the Control sheet:
   - Set Voltage: `5.0`
   - Set Current: `1.0`
   - OCP Enabled: `FALSE`
4. Click **"Apply Settings"**
5. Check that no errors appear
6. The PSU Status should show "Connected" and the settings should be applied

### Test with Real Hardware
1. Make sure your Korad KWR102 PSU is connected via USB/Serial
2. Open `VoltAmpero.xlsm`
3. Enter your COM port (e.g., `COM3`) in the PSU Port field
4. Click **"Connect PSU"**
5. Wait for "Connected" status
6. Enter your desired settings:
   - Set Voltage: (your desired voltage)
   - Set Current: (your desired current limit)
7. Click **"Apply Settings"**
8. The PSU should now be configured with your settings
9. Click **"Output ON"** to enable the output

## What to Expect
- **Before Fix**: Clicking "Apply Settings" would fail silently or give an error
- **After Fix**: Settings are properly applied to the PSU, and the PSU will auto-reconnect if needed

## Troubleshooting
If you still have issues after applying the fix:

1. **Check Python Connection**: 
   - Open Command Prompt in this directory
   - Run: `python -c "import voltampero; print('OK')"`
   - Should print "OK" with no errors

2. **Check COM Port**:
   - Run: `python -c "from voltampero import KoradKWR102; print(KoradKWR102.list_ports())"`
   - Verify your PSU's COM port is listed

3. **Test Python Directly**:
   ```cmd
   python
   >>> from voltampero import get_controller
   >>> c = get_controller(simulate=True)
   >>> c.connect_psu("SIM1")
   >>> c.set_voltage(5.0)
   >>> c.set_current(1.0)
   >>> print(c.get_psu_readings())
   ```

4. **Check xlwings**:
   - In Excel, press `Alt + F11`
   - In VBA Immediate Window (`Ctrl + G`), type: `?Application.Version`
   - Make sure xlwings is properly configured for your Python version

## Additional Improvements Made
- All PSU control methods now auto-reconnect if disconnected
- More robust error handling throughout
- Consistent behavior across all control buttons

---
**Fixed on**: 2026-01-08
**Files Modified**: 
- `VoltAmpero.bas` (VBA code)
- `voltampero.py` (Python controller)
