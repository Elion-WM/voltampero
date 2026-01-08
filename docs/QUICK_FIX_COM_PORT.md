# QUICK FIX: COM4 Port Locked

## THE PROBLEM
COM4 is locked because a previous Python process is still holding the port.
Each time you click a button, Python starts fresh, but the old process didn't release COM4.

## THE SOLUTION (3 Steps)

### Step 1: Close Everything
1. **Close Excel completely** (File > Exit)
2. Open **Task Manager** (Ctrl+Shift+Esc)
3. Look for any **python.exe** or **EXCEL.EXE** processes
4. End them if found

### Step 2: Physically Reset the Connection
**Option A - Unplug/Replug:**
- Unplug the USB cable from the PSU
- Wait 3 seconds  
- Plug it back in

**Option B - Power Cycle PSU:**
- Turn off the PSU power switch
- Wait 3 seconds
- Turn it back on

### Step 3: Restart Excel Fresh
1. Double-click: **start_excel_with_python.bat**
2. Cell B3 should show: **COM4**
3. Cell D3 should show: **Disconnected**

### Step 4: Test the Connection
1. Click **"Connect PSU"** button
   - D3 should change to "Connected"
   - If error, check PSU is powered on

2. Set values:
   - B16 = 5.0 (voltage)
   - B17 = 1.0 (current)

3. Click **"Apply Settings"**
   - E12 should show 5.0
   - E13 should show 1.0

4. Click **"Output ON"**
   - B11 should show 5.0V
   - B12 should show current

## WHY THIS HAPPENS

The auto-reconnect feature tries to reconnect on EVERY button click.
But if the port is locked, it fails silently and you don't see any change.

When you click "Disconnect All", it should release the port, but sometimes
Python processes don't close properly.

## ALTERNATIVE: Use Disconnect All Button

If the above doesn't work:
1. In Excel, click **"Disconnect All"** button
2. Wait 2 seconds
3. Click **"Connect PSU"** again
4. Try Apply Settings

This should release the port cleanly.

## PERMANENT FIX

I can modify the code to:
1. Detect when port is locked
2. Show error message
3. Not try to auto-reconnect if port is locked

Would you like me to implement this permanent fix?
