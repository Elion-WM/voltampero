# Live Readings & Ramp Data - Simple Solution

## The Real Issue

The live readings and ramp data WORK correctly, but you need to:
1. **Turn output ON** before starting ramp
2. **Start logging** to see data in Data tab

## Why Data Shows 0's

When PSU output is OFF:
- PSU correctly returns 0V/0A (this is correct behavior!)
- Voltage SETPOINT changes during ramp (5V → 10V → 15V)
- But actual OUTPUT voltage stays 0V (because output is OFF)

**Solution**: Turn output ON before ramp!

---

## Correct Procedure for Ramp with Live Data

### Step-by-Step:

1. **Connect PSU**
   - Click "Connect PSU"
   - Wait for "Connected" status

2. **Configure Ramp**
   - Set Start Voltage: 5V
   - Set End Voltage: 15V
   - Set Duration: 30 seconds

3. **Apply Initial Settings**
   - Click "Apply Settings" (this sets 5V as starting point)

4. **Turn Output ON** ← CRITICAL!
   - Click "Output ON"
   - PSU display shows voltage
   - DMM should read ~5V

5. **Start Logging** ← CRITICAL!
   - Click "Start Logging"
   - Data tab starts filling with readings

6. **Start Ramp**
   - Click "Start Ramp"
   - Voltage ramps from 5V to 15V
   - Data tab records the changing voltage

7. **Watch Data Tab**
   - Column C (PSU_Voltage_V): Shows 5.0, 5.5, 6.0, 6.5... 15.0
   - Column D (PSU_Current_A): Shows actual current draw
   - NOT all zeros!

---

## Live Readings in Control Tab

### Built-in Live Updates (No Auto-Refresh Needed!)

When **Logging is Active**:
- Live readings update automatically every 300ms (default)
- Control tab shows: LiveVoltage, LiveCurrent, LiveDMM
- This is built into the logging loop!

When **Ramp is Running with Logging**:
- Readings update as voltage ramps
- You see real-time values changing

### Manual Refresh (When Not Logging)

To manually update live readings:
1. Press `Alt + F8`
2. Select "RefreshReadings"
3. Click "Run"
4. Live values update once

Or create a "Refresh" button:
1. Insert → Shape → Button
2. Assign to "RefreshReadings" macro
3. Click button anytime to refresh

---

## Summary

### For Live Readings in Control Tab:
✅ **Automatic**: Start logging → readings update every 300ms
✅ **Manual**: Alt+F8 → RefreshReadings → updates once
❌ **Auto-refresh timer**: Removed (was causing Excel crashes)

### For Ramp Data in Data Tab:
✅ **Must do**: Start Logging BEFORE starting ramp
✅ **Must do**: Turn output ON before ramp
✅ **Result**: Data tab fills with real voltage/current values

---

## Example Test

**Test 1: Ramp with Data Collection**

```
1. Connect PSU
2. Set: Start=5V, End=15V, Duration=30s
3. Apply Settings (sets 5V)
4. Output ON (enables PSU output)
5. Start Logging (begins data collection)
6. Start Ramp (voltage begins ramping)
7. Watch Data tab → voltage goes 5→6→7...→15V
8. Stop Logging when done
9. Output OFF
```

**Test 2: Monitor Live Readings Without Ramp**

```
1. Connect PSU
2. Set voltage/current
3. Apply Settings
4. Output ON
5. Start Logging
6. Watch Control tab → Live values update
7. Stop Logging when done
```

---

## What's Fixed

✅ VBA no longer crashes (removed problematic auto-refresh auto-start)
✅ Logging works (built-in live updates)
✅ Ramp works (changes voltage)
✅ Data collection works during logging
⚠️ **Must turn output ON** to see non-zero voltage values
⚠️ **Must start logging** to populate Data tab

---

## Troubleshooting

### Ramp shows 0's in Data tab

**Check:**
- [ ] Is logging started? (Click "Start Logging" BEFORE "Start Ramp")
- [ ] Is output ON? (Click "Output ON" BEFORE "Start Ramp")
- [ ] Is PSU connected? (Status shows "Connected")

**If still 0's:**
- Output is probably OFF
- PSU correctly reports 0V when output is disabled
- Turn output ON and try again

### Live readings don't update in Control tab

**Solution 1**: Start logging
- Logging automatically updates live readings

**Solution 2**: Manual refresh
- Press Alt+F8 → Run "RefreshReadings"

---

The key point: **Logging + Output ON = Live data!**
