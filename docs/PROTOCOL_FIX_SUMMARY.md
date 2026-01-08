# Korad KWR102 Protocol Fix - SOLVED!

## Issue Summary
The PSU wasn't responding to commands because the Python code was using the wrong protocol format.

## Root Cause
**Korad KWR102 uses a different command format** than other Korad models:
- ❌ Wrong: `VSET1:` → No response
- ✅ Correct: `VSET:` → Works!
- **Required**: All commands need `\r` (carriage return) terminator
- **Required**: RTS/DTR serial lines must be set

## What Was Fixed in `psu_korad.py`

### 1. Command Terminator
```python
# OLD (wrong):
self.serial.write(cmd.encode('ascii'))

# NEW (correct):
self.serial.write((cmd + '\r').encode('ascii'))
```

### 2. Flow Control Lines
```python
# Added to connect():
self.serial.setRTS(True)
self.serial.setDTR(True)
```

### 3. Command Format Changes
| Command | Old (Wrong) | New (Correct) |
|---------|-------------|---------------|
| Set Voltage | `VSET1:18.00` | `VSET:18.00` |
| Get Voltage Setpoint | `VSET1?` | `VSET?` |
| Get Output Voltage | `VOUT1?` | `VOUT?` |
| Set Current | `ISET1:0.100` | `ISET:0.100` |
| Get Current Setpoint | `ISET1?` | `ISET?` |
| Get Output Current | `IOUT1?` | `IOUT?` |

### 4. Timing
Increased command delay from 0.05s to 0.1s for reliable communication.

## Test Results

### Before Fix:
```
PSU ID: Unknown
Voltage: 0.0V (failed)
Current: 0.0A (failed)
```

### After Fix:
```
PSU ID: KORAD KWR102 V2.3 SN:000000166271
Voltage: 18.0V ✓
Current: 0.1A ✓
Status: Working perfectly!
```

## How to Use

### From Python:
```python
from psu_korad import KoradKWR102

psu = KoradKWR102()
psu.connect("COM4")
psu.set_voltage(18.0)
psu.set_current(0.1)
psu.output_on()  # Turn on output
```

### From Excel:
1. Open `VoltAmpero.xlsm`
2. Enter `COM4` in PSUPort cell
3. Click **"Connect PSU"**
4. Enter desired voltage (e.g., 18.0V)
5. Enter desired current (e.g., 0.1A)
6. Click **"Apply Settings"**
7. Click **"Output ON"** to enable PSU output

## Verification
- PSU display shows: 18.0V / 0.1A
- Python test: `test_direct_set.py` → SUCCESS
- Excel integration: Ready to test

## Files Modified
- `psu_korad.py` - Fixed command protocol
- `voltampero.py` - Already had proper integration (no changes needed)
- `VoltAmpero.bas` - Already updated with proper connection management

## Next Steps for User
1. Open Excel with `VoltAmpero.xlsm`
2. Click "Connect PSU" with COM4
3. Set 18V, 0.1A
4. Click "Apply Settings"
5. Click "Output ON"
6. Check DMM for reading (should see ~18V if DMM is connected to PSU output)

---
**Fixed**: 2026-01-08  
**PSU Model**: Korad KWR102 V2.3  
**Serial**: SN:000000166271
