# VoltAmpero - Quick User Guide

## How to Start VoltAmpero

**IMPORTANT:** Always start Excel using the batch file!

1. Double-click **start_excel_with_python.bat**
2. Keep the command window open while using Excel
3. Excel will open automatically with VoltAmpero.xlsm

## Understanding the Display

### Input Cells (Yellow Background - Row 16-17)
- **B16: Set Voltage** - Type your desired voltage here
- **B17: Set Current** - Type your desired current limit here
- **B18: OCP Enabled** - Checkbox for overcurrent protection

### PSU Setpoints Display (Column E, Rows 12-13)
- **E12: Set Voltage** - Shows what voltage is programmed into PSU
- **E13: Set Current** - Shows what current limit is programmed into PSU

### Live Readings (Column B, Rows 11-12)
- **B11: Voltage (V)** - Shows ACTUAL output voltage (only when Output is ON)
- **B12: Current (A)** - Shows ACTUAL output current (only when Output is ON)
- **B13: DMM** - Shows multimeter reading

## How to Control the PSU

### Step-by-Step Workflow:

1. **Initialize Simulated Mode** (for testing)
   - Click **"Test (Simulated)"** button
   - This connects to simulated PSU and DMM

2. **Set Desired Values**
   - Type voltage in cell **B16** (e.g., 5.0)
   - Type current in cell **B17** (e.g., 1.0)
   - Check/uncheck **B18** for OCP

3. **Apply Settings**
   - Click **"Apply Settings"** button
   - You'll see a confirmation message
   - **E12** and **E13** will update to show the setpoints

4. **Turn Output ON**
   - Click **"Output ON"** button
   - Now **B11** and **B12** will show actual voltage/current
   - The PSU is now delivering power

5. **Turn Output OFF** (when done)
   - Click **"Output OFF"** button
   - Output voltage/current will go to 0

## Important Notes

### Why do I see 0V even after Apply Settings?
- The PSU **output is OFF** by default
- **Setpoints** (E12, E13) show what's programmed
- **Live readings** (B11, B12) show actual output (0V when off)
- Click **"Output ON"** to see voltage/current

### Difference Between Setpoints and Live Readings
```
B16, B17 (Input)  →  Apply Settings  →  E12, E13 (Setpoints)
                                              ↓
                                        Output ON/OFF
                                              ↓
                                      B11, B12 (Live Output)
```

## Troubleshooting

### Buttons don't work / Error 62
- Make sure you started Excel using **start_excel_with_python.bat**
- Keep the command window open
- If closed, restart Excel using the batch file

### Can't edit voltage/current cells
- Cells B16, B17 should have yellow background
- If locked, run: `unlock_input_cells.py`

### PSU doesn't change settings
- Check if you clicked **"Apply Settings"** first
- Check **E12, E13** - these should show your values
- Then click **"Output ON"** to enable output
- Check **B11, B12** - these will show actual voltage/current

## Logging Data

1. Click **"Start Logging"** to begin recording
2. Data appears in the **Data** sheet
3. Click **"Stop Logging"** when done
4. Click **"Export CSV"** to save to file
5. Click **"Clear Data"** to erase logged data

## Voltage Ramp

1. Set ramp parameters in rows 21-26
2. Click **"Start Ramp"** to begin
3. Click **"Pause Ramp"** to pause
4. Click **"Stop Ramp"** to stop

## Real Hardware (Not Simulated)

When connecting to real hardware:

1. **Connect PSU:**
   - Enter COM port in **B3** (e.g., COM3)
   - Click **"Connect PSU"**

2. **Connect DMM:**
   - Click **"Connect DMM"** (auto-detects USB device)

3. Use the same workflow as simulated mode

4. **Disconnect All** when done
