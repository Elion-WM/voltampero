# Korad KWR102 Power Supply - Communication Protocol

## Hardware Connection

### Physical Interface
- **Connection Type**: USB to Serial (CH340 chip)
- **Port**: Appears as standard COM port in Windows
- **Cable**: USB-A to USB-B (standard printer cable)

### Serial Port Settings
```
Baudrate:    115200
Data bits:   8
Parity:      None
Stop bits:   1
Flow control: None
RTS:         True (must be set!)
DTR:         True (must be set!)
Timeout:     1.0s (1000ms) for normal operations
             0.1s (100ms) for fast query operations
```

**CRITICAL**: The KWR102 requires RTS and DTR control lines to be set HIGH. Without these, the device will not respond to commands.

```python
import serial

ser = serial.Serial(
    port='COM3',
    baudrate=115200,
    bytesize=serial.EIGHTBITS,
    parity=serial.PARITY_NONE,
    stopbits=serial.STOPBITS_ONE,
    timeout=1.0
)
# REQUIRED for KWR102:
ser.setRTS(True)
ser.setDTR(True)
```

---

## Protocol Specification

### Command Format
- **Encoding**: ASCII
- **Terminator**: `\r` (Carriage Return, 0x0D)
- **Commands**: Case-sensitive
- **Delay**: Minimum 10ms between commands (50-100ms recommended for reliable operation)

### Response Format
- **Queries** (`?` suffix): Device responds with ASCII value
- **Set commands**: No response (silent success)
- **Errors**: Device does not send error messages, simply ignores invalid commands

---

## Command Set

### Device Identification

#### Get Device ID
```
Command: *IDN?
Response: KORAD KWR102 V2.3 (or similar)
Example:
  >>> *IDN?\r
  <<< KORAD KWR102 V2.3
```

---

### Voltage Control

#### Set Voltage
```
Command: VSET:<voltage>
Format:  VSET:XX.XX (5 characters, 2 decimal places)
Range:   0.00V to 60.00V
Example:
  >>> VSET:12.50\r  (set to 12.5V)
  >>> VSET:05.00\r  (set to 5.0V)
```

**Important Notes:**
- KWR102 V2.3 uses `VSET:` (with colon, NO channel number)
- Other Korad models may use `VSET1:` format
- Always format as 5 characters with 2 decimals: `05.00`, not `5.0`

#### Query Voltage Setpoint
```
Command: VSET?
Response: XX.XX (voltage setpoint in volts)
Example:
  >>> VSET?\r
  <<< 12.50
```

#### Query Output Voltage (Actual)
```
Command: VOUT?
Response: XX.XXX (actual output voltage)
Example:
  >>> VOUT?\r
  <<< 12.487
```

---

### Current Control

#### Set Current Limit
```
Command: ISET:<current>
Format:  ISET:X.XXX (5 characters, 3 decimal places)
Range:   0.000A to 5.000A (model dependent)
Example:
  >>> ISET:1.500\r  (set to 1.5A)
  >>> ISET:0.100\r  (set to 100mA)
```

**Important Notes:**
- KWR102 V2.3 uses `ISET:` (with colon, NO channel number)
- Format as 5 characters with 3 decimals: `0.100`, not `0.1`

#### Query Current Setpoint
```
Command: ISET?
Response: X.XXX (current limit in amps)
Example:
  >>> ISET?\r
  <<< 1.500
```

#### Query Output Current (Actual)
```
Command: IOUT?
Response: X.XXXX (actual output current)
Example:
  >>> IOUT?\r
  <<< 0.0987
```

---

### Output Control

#### Turn Output ON
```
Command: OUT:1
Note: Uses colon format (OUT:1), not OUT1
Example:
  >>> OUT:1\r
```

#### Turn Output OFF
```
Command: OUT:0
Note: Uses colon format (OUT:0), not OUT0
Example:
  >>> OUT:0\r
```

**Important Notes:**
- KWR102 V2.3 requires `OUT:1` and `OUT:0` (with colon)
- Other models may use `OUT1`/`OUT0` (without colon)
- After changing output state, wait 300ms for PSU to settle

---

### Protection Features

#### Over Current Protection (OCP)
```
Enable:  OCP1\r
Disable: OCP0\r
```

#### Over Voltage Protection (OVP)
```
Enable:  OVP1\r
Disable: OVP0\r
```

---

### Status Query

#### Get Status
```
Command: STATUS?
Response: 1-byte status code
Bit 0: 0=CV mode, 1=CC mode
Bit 6: 0=Output OFF, 1=Output ON
Example:
  >>> STATUS?\r
  <<< @ (0x40 = output ON, CV mode)
```

**Note**: STATUS? command can sometimes hang on KWR102 V2.3. Use with short timeout (300ms) and fallback to checking VOUT? for output state.

---

## Threading and Performance

### Thread-Safe Operation
When multiple threads access the PSU (e.g., ramping voltage while logging readings), use a threading lock:

```python
import threading

class KoradKWR102:
    def __init__(self):
        self._serial_lock = threading.Lock()
    
    def _send_command(self, cmd: str):
        with self._serial_lock:
            self.serial.reset_input_buffer()
            self.serial.write((cmd + '\r').encode('ascii'))
            time.sleep(0.02)  # Query delay
            if '?' in cmd:
                return self.serial.read(100).decode('ascii').strip()
            else:
                time.sleep(0.01)  # Set command delay
                return ""
```

### Timing Recommendations
- **Query commands**: 20-50ms delay (fast reading for data logging)
- **Set commands**: 10-100ms delay (slower for reliability during ramp)
- **Output commands**: 300ms delay after OUT:1 or OUT:0
- **Between commands**: Minimum 10ms

### Optimizations for Speed
- Use shorter timeouts for queries (100ms vs 1000ms)
- Minimize delays for fast operations (ramping)
- Use threading locks to prevent conflicts
- Batch read operations when possible

---

## Common Issues and Solutions

### Issue: No Response from Device
**Symptoms**: Commands sent but no response  
**Causes**:
1. RTS/DTR not set
2. Wrong baudrate
3. Wrong terminator (using `\n` instead of `\r`)

**Solutions**:
```python
ser.setRTS(True)
ser.setDTR(True)
time.sleep(0.2)  # Wait for handshake
```

### Issue: Wrong Command Format
**Symptoms**: Commands ignored  
**Problem**: Using `VSET1:` format on V2.3 firmware  
**Solution**: Use `VSET:` (no channel number) for KWR102 V2.3

### Issue: Slow Data Logging
**Symptoms**: 5-10 second intervals instead of 500ms  
**Cause**: Long timeouts and delays accumulating  
**Solution**:
- Reduce query timeout to 100ms
- Reduce delay to 20-50ms
- Use threading lock to prevent blocking

### Issue: Ramp Takes Too Long
**Symptoms**: 60s ramp takes 180s  
**Cause**: Each voltage set command has 100ms delay  
**Solution**:
- Use "fast mode" for ramp operations (10ms delay)
- Calculate target times and compensate for command overhead
- Use threading lock to prevent logging interference

---

## Comparison with Other Korad Models

| Feature | KWR102 V2.3 | Other Models (e.g., KA3005P) |
|---------|-------------|------------------------------|
| VSET command | `VSET:` | `VSET1:` |
| ISET command | `ISET:` | `ISET1:` |
| Output ON | `OUT:1` | `OUT1` |
| Output OFF | `OUT:0` | `OUT0` |
| VSET query | `VSET?` | `VSET1?` |
| ISET query | `ISET?` | `ISET1?` |

**Always verify your specific model's protocol!** Use Wireshark or serial monitor to capture actual commands.

---

## Example: Complete Communication Session

```python
import serial
import time

# Connect
ser = serial.Serial('COM3', 115200, timeout=1.0)
ser.setRTS(True)
ser.setDTR(True)
time.sleep(0.2)

# Get ID
ser.write(b'*IDN?\r')
time.sleep(0.05)
print(ser.read(100))  # KORAD KWR102 V2.3

# Set 12V, 1A limit
ser.write(b'VSET:12.00\r')
time.sleep(0.05)
ser.write(b'ISET:1.000\r')
time.sleep(0.05)

# Turn output ON
ser.write(b'OUT:1\r')
time.sleep(0.3)  # Wait for output to stabilize

# Read actual values
ser.write(b'VOUT?\r')
time.sleep(0.05)
voltage = float(ser.read(100))

ser.write(b'IOUT?\r')
time.sleep(0.05)
current = float(ser.read(100))

print(f"Output: {voltage}V, {current}A")

# Turn output OFF
ser.write(b'OUT:0\r')
time.sleep(0.3)

ser.close()
```

---

## Testing and Debugging

### Serial Monitor Test
Use a serial monitor (RealTerm, PuTTY, or Python) to manually test commands:

```python
import serial

ser = serial.Serial('COM3', 115200, timeout=1.0)
ser.setRTS(True)
ser.setDTR(True)

while True:
    cmd = input("Command: ")
    ser.write((cmd + '\r').encode())
    time.sleep(0.1)
    response = ser.read(100)
    print(f"Response: {response}")
```

### Wireshark Capture
To reverse-engineer unknown protocols:
1. Install USBPcap for Windows
2. Capture USB traffic while using official software
3. Filter by USB device address
4. Look for URB_BULK packets with ASCII data
5. Identify command patterns and responses

---

## References

- Korad official documentation (limited)
- Community reverse-engineered protocols: https://sigrok.org/wiki/Korad_KAxxxxP_series
- This protocol was verified via Wireshark USB capture on KWR102 V2.3 (2026-01-08)

---

*Last verified: 2026-01-08 with Korad KWR102 V2.3 firmware*
