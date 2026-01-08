# Fix: Live Readings and Ramp Data

## Issues Fixed

### 1. Live Readings Don't Auto-Update ✅
**Problem**: Control tab shows static values, doesn't refresh  
**Solution**: Auto-refresh now starts automatically when:
- Connect PSU
- Start Logging
- Start Ramp

### 2. Ramp Shows 0's in Data Tab ✅  
**Problem**: During ramp, voltage/current show as 0.0 instead of actual values  
**Root Cause**: PSU output must be ON for readings to show values  
**Solution**: Turn output ON before starting ramp

---

## How to Use

### Live Readings (Control Tab)

**Auto-start refresh when:**
1. Click "Connect PSU" → Live readings start updating every 1 second
2. Click "Start Logging" → Live readings update
3. Click "Start Ramp" → Live readings update

**Stop refresh when:**
1. Click "Disconnect All"
2. Click "Stop Logging"
3. Click "Stop Ramp"

**Manual control:**
- Press `Alt + F8` → Run "StartAutoRefresh" to start
- Press `Alt + F8` → Run "StopAutoRefresh" to stop

### Ramp with Proper Data Logging

**IMPORTANT**: You must turn output ON before starting ramp!

**Correct sequence:**
1. Connect PSU
2. Set voltage range (e.g., Start: 0V, End: 18V)
3. Set ramp duration (e.g., 60 seconds)
4. **Click "Apply Settings"** (set initial voltage)
5. **Click "Output ON"** ← CRITICAL!
6. Click "Start Logging" (if you want to log data)
7. Click "Start Ramp"
8. Watch Data tab fill with actual voltage/current values

**Why this is needed:**
- PSU returns 0V/0A when output is OFF
- Ramp changes voltage setpoint, but output must be ON to see actual values
- DMM will also read 0V if PSU output is OFF

---

## Testing

### Test Live Readings:

1. Open VoltAmpero.xlsm
2. Close Excel and reopen (to load updated VBA)
3. Connect PSU
4. Watch Control tab → LiveVoltage, LiveCurrent, LiveDMM cells
5. They should update every 1 second
6. Click "Output ON"
7. Values should change from 0 to actual voltage/current

### Test Ramp with Data:

1. Connect PSU
2. Set ramp: Start 5V, End 15V, Duration 30s
3. Click "Apply Settings" (sets 5V)
4. **Click "Output ON"** ← Don't skip!
5. Click "Start Logging"
6. Click "Start Ramp"
7. Watch Data tab:
   - Column C (PSU_Voltage_V) should show values changing from 5 to 15
   - Column D (PSU_Current_A) should show actual current draw
   - NOT all zeros!

---

## Files Updated

### VoltAmpero.bas
- `ConnectPSU()` - Added StartAutoRefresh
- `DisconnectAll()` - Added StopAutoRefresh
- `StartLogging()` - Added StartAutoRefresh
- `StopLogging()` - Added StopAutoRefresh
- `StartRamp()` - Added StartAutoRefresh
- `StopRamp()` - Added StopAutoRefresh

---

## Troubleshooting

### Live readings still show 0

**Check:**
1. Is PSU connected? (Status shows "Connected")
2. Is output ON? (PSU display shows voltage)
3. Is auto-refresh running? (Try Alt+F8 → StartAutoRefresh)

### Ramp still shows 0's in data

**Check:**
1. Did you turn output ON before starting ramp?
2. Is PSU actually connected?
3. Are the voltage/current setpoints configured?

**To verify PSU is working:**
- Manually set 10V and click "Apply Settings"
- Click "Output ON"
- Check PSU display → should show 10V
- Check DMM → should read ~10V
- Check LiveVoltage in Control tab → should show 10V

### Auto-refresh causes Excel to slow down

**Solution:** 
- Refresh interval is 1 second (reasonable)
- Stop auto-refresh when not needed:
  - Alt+F8 → StopAutoRefresh
- Auto-refresh stops automatically when you disconnect

---

## Summary

✅ Live readings now auto-update (every 1 second)  
✅ Auto-refresh starts when Connect PSU / Start Logging / Start Ramp  
✅ Auto-refresh stops when Disconnect / Stop Logging / Stop Ramp  
⚠️ **Must turn output ON before ramp to see actual values**  
✅ Ramp will now show real voltage/current data (not 0's)

---

**Key Point**: PSU output must be ON to read actual voltage/current!
When output is OFF, PSU correctly reports 0V/0A.
