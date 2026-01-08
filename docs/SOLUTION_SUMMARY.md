# VoltAmpero - Complete Fix Summary

## Status: WORKING (with minor issue on Output OFF)

### What's Fixed ✅

1. **Apply Settings** - ✅ WORKING
   - Fixed command format: `VSET:` instead of `VSET1:`
   - Added `\r` terminator to all commands
   - Set RTS/DTR control lines
   - **Result**: Can now set voltage and current from Excel

2. **Output ON** - ✅ WORKING
   - Command: `OUT1\r`
   - PSU successfully turns on output
   - DMM should read voltage when output is on

3. **Communication** - ✅ WORKING
   - Correct baudrate: 115200
   - Correct protocol discovered for Korad KWR102 V2.3
   - Connection from Excel works

### Current Issue ⚠️

**Output OFF** - Partially working
- Command sent: `OUT0\r`
- Issue: PSU may not immediately respond or needs different command
- Workaround added: 0.3s delay after command

### Root Cause of All Issues

The Korad KWR102 uses a **different protocol** than standard Korad models:

| Feature | Standard Korad | KWR102 V2.3 |
|---------|---------------|-------------|
| Set Voltage | `VSET1:18.00` | `VSET:18.00` |
| Query Voltage | `VSET1?` | `VSET?` |
| Set Current | `ISET1:0.100` | `ISET:0.100` |
| Query Current | `ISET1?` | `ISET?` |
| Output Voltage | `VOUT1?` | `VOUT?` |
| Output Current | `IOUT1?` | `IOUT?` |
| **Terminator** | None or `\n` | `\r` (required!) |
| **Flow Control** | Not needed | RTS/DTR must be set |

## Files Modified

### 1. `psu_korad.py`
**Changes:**
- Added `\r` terminator to all commands
- Changed command format from `VSET1:` to `VSET:`
- Set RTS/DTR control lines in `connect()`
- Increased command delay to 0.1s
- Fixed `get_status()` with timeout and fallback
- Added 0.3s delay after output state change

**Key functions:**
```python
def _send_command(self, cmd: str):
    self.serial.write((cmd + '\r').encode('ascii'))
    time.sleep(0.1)
    
def connect(self, port: str):
    self.serial.setRTS(True)
    self.serial.setDTR(True)
    
def set_voltage(self, voltage: float):
    cmd = f"VSET:{voltage:05.2f}"  # Not VSET1:
```

### 2. `voltampero.py`
**Changes:**
- Fixed global controller instance management
- Added auto-reconnect with error handling
- Better status synchronization with Excel

### 3. `VoltAmpero.bas`
**Changes:**
- Added PSU connection check before applying settings
- Reset global controller on connection to avoid COM port locks
- Added error handling and success messages
- Shows popup confirmation when settings applied

## How to Use (Working Procedure)

### Initial Setup
1. Connect Korad KWR102 PSU to USB (COM4)
2. Open `VoltAmpero.xlsm`
3. Click **"Connect PSU"** button
4. Wait for "Connected" status

### Set Voltage/Current
1. Enter desired voltage (e.g., `18`)
2. Enter desired current limit (e.g., `0.1`)
3. Click **"Apply Settings"**
4. Should see success popup: "Settings applied: 18V, 0.1A"
5. **Verify**: Check PSU display shows 18.0V / 0.1A

### Turn Output ON
1. Click **"Output ON"** button
2. PSU output activates
3. **Verify**: DMM connected to PSU output should read ~18V

### Turn Output OFF
1. Click **"Output OFF"** button
2. Wait 1-2 seconds
3. PSU output should turn off
4. **Note**: If output doesn't turn off immediately, click again or manually turn off on PSU front panel

## Troubleshooting

### Issue: "Apply Settings" doesn't work
**Solution:**
1. Click "Disconnect All"
2. Wait 2 seconds
3. Click "Connect PSU" again
4. Try "Apply Settings" again

### Issue: COM4 Access Denied
**Solution:**
1. Close Excel completely
2. Open Task Manager → End any Python processes
3. Reopen Excel
4. Click "Connect PSU"

### Issue: Output OFF doesn't turn off output
**Temporary Workaround:**
1. Click "Output OFF" button
2. Wait 2 seconds
3. If still on, click again
4. Or manually turn off using PSU front panel button

**Permanent Fix (if needed):**
The PSU might need a different OFF command. If this persists, we can test:
- `OUT 0` (with space)
- `OUTP0`
- Different timing

### Issue: DMM doesn't show reading
**Check:**
1. Is DMM connected to PSU output terminals?
2. Is PSU output ON? (check PSU display for "OUT" indicator)
3. Is DMM in correct mode (voltage measurement)?
4. Are voltage/current settings applied? (check PSU display)

## Testing Done

✅ Direct Python test - Set 18V, 0.1A - SUCCESS
✅ Excel "Connect PSU" - SUCCESS  
✅ Excel "Apply Settings" - SUCCESS
✅ Excel "Output ON" - SUCCESS
⚠️ Excel "Output OFF" - Needs verification

## Next Steps for User

1. **Test Output OFF**: 
   - With output ON and 18V applied
   - Click "Output OFF" in Excel
   - Check if PSU turns off
   - Report if it works or not

2. **If Output OFF fails**:
   - Tell me what happens (any error? no response? delay?)
   - I'll provide additional fix

3. **Once working**:
   - Test complete workflow several times
   - Verify DMM readings match PSU settings
   - Test with different voltage/current values

## Key Learnings

1. **Korad KWR102 is different** from other Korad models
2. **Protocol discovery** was key - testing revealed `\r` terminator requirement
3. **COM port locking** caused most issues - fixed with proper controller management
4. **STATUS? command hangs** - workaround with voltage-based detection
5. **Timing matters** - PSU needs delays to process commands

---

**PSU Model**: Korad KWR102 V2.3 SN:000000166271  
**Last Updated**: 2026-01-08  
**Status**: 95% Working - Output OFF needs final verification
