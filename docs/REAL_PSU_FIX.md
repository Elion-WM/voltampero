# Fix for Real PSU Not Applying Settings from Excel

## Root Cause Found
The diagnostic test revealed the problem:
```
PSU connected: False
PermissionError(13, 'Odmowa dostępu', None, 5) - Access Denied
```

**Issue**: When you click "Connect PSU" in Excel, it creates a Python controller instance that opens COM4. When you later click "Apply Settings", it gets a **different** Python instance that can't access COM4 because it's already locked by the first instance.

## The Fix

### Files Updated:
1. `voltampero.py` - Improved global controller management and auto-reconnect
2. `VoltAmpero.bas` - Updated VBA macros to properly reset controller

### What Changed:

**1. Global Controller Management**
- The `get_controller()` function now properly checks if PSU is already connected
- Avoids trying to reconnect if the port is already open
- Updates Excel status if reconnection fails

**2. VBA Macros**
- `ConnectPSU()` now resets the global controller before connecting
- This ensures a fresh connection without COM port conflicts
- `InitSimulated()` also resets controller properly

**3. Auto-Reconnect Logic**
- Added better error handling in `set_voltage()` and `set_current()`
- Now prints debug messages when auto-reconnect is attempted
- Returns False if reconnection fails (preventing silent failures)

## How to Apply the Fix

### Step 1: Update VBA Code
1. Open `VoltAmpero.xlsm`
2. Press `Alt + F11` (VBA Editor)
3. Right-click the `VoltAmpero` module → **Remove VoltAmpero**
4. Right-click `VBAProject` → **Import File**
5. Select the updated `VoltAmpero.bas` file
6. Press `Ctrl + S` to save
7. Close VBA Editor

### Step 2: Close and Reopen Excel
- This ensures Python starts fresh
- **Important**: Close Excel completely, don't just close the workbook

### Step 3: Test with Real PSU

**Test Sequence:**
1. Make sure your Korad KWR102 PSU is connected via USB
2. Open `VoltAmpero.xlsm`
3. In the PSUPort cell (B3), enter your COM port: `COM4`
4. Click **"Connect PSU"** button
5. Wait for status to change to "Connected"
6. Enter test settings:
   - Set Voltage: `5.0`
   - Set Current: `1.0`
7. Click **"Apply Settings"**
8. You should see a success popup
9. **Check your PSU display** - it should show 5.0V / 1.0A setpoints

### Step 4: Verify Settings Were Applied
Look at the PSU display:
- Press the V-SET button on your PSU to see voltage setpoint
- Press the I-SET button to see current setpoint
- They should match what you entered in Excel

## Troubleshooting

### If you still get "Access Denied" error:

**Option A: Full Reset (Recommended)**
1. Close Excel completely
2. Open Task Manager (`Ctrl + Shift + Esc`)
3. Look for any Python processes and end them
4. Disconnect and reconnect the PSU USB cable
5. Wait 5 seconds
6. Reopen Excel and try again

**Option B: Test Direct Connection**
Run this Python test with Excel closed:
```cmd
cd C:\Users\User\GitHub\voltampero
python -c "from psu_korad import KoradKWR102; psu=KoradKWR102(); print('Connecting...'); result=psu.connect('COM4'); print(f'Result: {result}'); psu.set_voltage(5.0); psu.set_current(1.0); print(f'Set to: {psu.get_voltage_setpoint()}V, {psu.get_current_setpoint()}A'); psu.disconnect()"
```

This tests if Python can access COM4 at all.

### If "Connect PSU" button says "Connected" but settings don't apply:

1. Click **"Disconnect All"** first
2. Wait 2 seconds
3. Click **"Connect PSU"** again
4. Then try **"Apply Settings"**

### If PSU shows wrong COM port number:

Check Windows Device Manager:
1. Press `Win + X` → Device Manager
2. Expand "Ports (COM & LPT)"
3. Find your USB Serial device
4. Note the COM number (e.g., COM4, COM3, etc.)
5. Update Excel with the correct COM port

## Debug Script

If issues persist, run the diagnostic with Excel open and PSU connected:
```cmd
python test_real_psu_apply.py
```

This will show exactly where the connection is failing.

## Expected Behavior After Fix

### Connect PSU:
- Click button → Status changes to "Connected"
- No error messages
- PSU is ready to accept commands

### Apply Settings:
- Click button → **Popup shows "Settings applied: 5.0V, 1.0A, OCP=True"**
- PSU display updates to show new setpoints
- Settings persist even if you change them in Excel later

### Disconnect All:
- Click button → Status changes to "Disconnected"
- COM port is released
- Can reconnect without errors

## Technical Details

### Why This Happened
Python's `xlwings` runs each `RunPython` command in isolation. Each VBA button click was creating a **new** Python interpreter context, but the COM port was locked by the previous context. The fix ensures:

1. **Global Controller**: Single instance persists across VBA calls
2. **Smart Reset**: ConnectPSU resets the global before connecting
3. **Status Sync**: Excel status is kept in sync with actual connection state

### COM Port Locking
Windows serial ports are **exclusive** - only one process can open them at a time. The fix ensures proper cleanup and reuse of the same controller instance.

---

**Last Updated**: 2026-01-08  
**Files Modified**:
- `voltampero.py` 
- `VoltAmpero.bas`
- Created: `test_real_psu_apply.py` (diagnostic tool)
