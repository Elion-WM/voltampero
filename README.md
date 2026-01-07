# VoltAmpero - Lab Instrument Control Software

Control **Korad KWR102** power supply and **UNI-T UT8804E** multimeter with Excel-based interface.

## Features

- **Data Logging** with start/stop switch, CSV export
- **Voltage Ramp** with multiple cycles, ping-pong mode
- **OCP Control** (Over Current Protection)
- **Parallel Timestamps** for synchronized readings
- **Live Charts** in Excel (auto-updating)
- **No Admin Rights** required on Windows 11

## Requirements

- Windows 10/11
- Python 3.8+ (can install from Microsoft Store - no admin needed)
- Microsoft Excel (with macros enabled)
- Korad KWR102 USB cable
- UNI-T UT8804E USB cable

## Installation (No Admin Required)

### 1. Install Python from Microsoft Store

1. Open Microsoft Store
2. Search "Python 3.11"
3. Click Install (no admin needed)

### 2. Install Dependencies

Open Command Prompt (Win+R, type `cmd`, Enter):

```cmd
pip install --user pyserial hidapi xlwings
```

### 3. Install xlwings Excel Add-in

```cmd
xlwings addin install
```

### 4. Set Up Excel Workbook

See **EXCEL_SETUP.md** for detailed instructions, or:

1. Open Excel, create new workbook
2. Save as `VoltAmpero.xlsm` (macro-enabled)
3. Follow the setup guide to create Control and Data sheets

## Quick Start

### Test with Simulated Devices

1. Open `VoltAmpero.xlsm`
2. Click "Test (Simulated)" button
3. Click "Start Logging"
4. Watch data appear in real-time
5. Click "Stop Logging"
6. Click "Export CSV"

### Connect Real Hardware

1. Connect Korad KWR102 via USB
2. Connect UNI-T UT8804E via USB
3. Open Device Manager to find COM port (e.g., COM3)
4. Enter COM port in Excel (PSUPort cell)
5. Click "Connect PSU"
6. Click "Connect DMM"

## Usage

### Basic Operation

1. Set voltage and current in Control sheet
2. Click "Apply Settings"
3. Click "Output ON"
4. Click "Start Logging" to record data

### Voltage Ramp

Configure in Excel:
- **Start V**: Starting voltage
- **End V**: Target voltage
- **Duration**: Time in seconds
- **Cycles**: Number of repetitions (0 = infinite)
- **Delay**: Pause between cycles
- **Ping-Pong**: Alternate direction each cycle

Click "Start Ramp" to begin.

### Data Export

- Click "Export CSV" to save timestamped file
- Or use Data sheet directly for Excel charts

## File Structure

```
voltampero/
├── README.md              # This file
├── EXCEL_SETUP.md         # Excel configuration guide
├── requirements.txt       # Python dependencies
├── psu_korad.py          # Korad KWR102 driver
├── multimeter_unit.py    # UNI-T UT8804E driver
├── voltampero.py         # Main controller
└── VoltAmpero.xlsm       # Excel workbook (you create)
```

## Standalone Mode (No Excel)

Run directly from command line:

```cmd
python voltampero.py
```

This runs a test with simulated devices.

## Troubleshooting

### PSU not connecting
- Check COM port in Device Manager
- Try different USB cable
- Verify baud rate (115200)

### DMM not found
- Install hidapi: `pip install hidapi`
- Check USB connection
- Device should appear as HID device

### Excel errors
- Enable macros in Trust Center
- Check xlwings.conf PYTHONPATH
- Run `xlwings addin install`

### Permission errors
- No admin rights needed
- Use `pip install --user` for packages

## Communication Protocols

### Korad KWR102 (Serial)
- Baud: 9600, 8N1
- Commands: `VSET:xx.xx`, `ISET:x.xxx`, `OUT1`/`OUT0`, `OCP1`/`OCP0`
- Query: `*IDN?`, `VSET?`, `ISET?`, `VOUT?`, `IOUT?`, `STATUS?`

### UNI-T UT8804E (USB HID)
- USB HID via CP2110 USB-to-UART bridge
- Vendor ID: 0x10C4 (Silicon Labs)
- Product ID: 0xEA80
- UART: 9600 baud, 8N1

#### Connection Sequence
1. Enable UART: Send Feature Report 0x41 with value 0x01
2. Configure UART: Send Feature Report 0x50 (9600 baud, 8N1)
3. Send init command: `abcd040005010a00` (starts data streaming)
4. Read data continuously from HID interrupt endpoint

#### Packet Format
```
Offset  Size  Description
------  ----  -----------
0-1     2     Header: 0xAB 0xCD
2       1     Packet type (0x21)
3       1     Reserved (0x00)
4-5     2     Mode/Range (0x02 0x08 = DC Voltage)
6-9     4     Flags
10-13   4     Value: IEEE 754 float, little-endian, NEGATED (use abs())
14+     ...   Additional data
```

#### Mode Bytes (offset 4-5)
| Byte 4 | Byte 5 | Mode |
|--------|--------|------|
| 0x02   | 0x08   | DC Voltage |
| 0x02   | 0x00   | DC Voltage |
| 0x03   | -      | AC Voltage |
| 0x04   | -      | DC Current mA |
| 0x08   | -      | Resistance |

## Timing Characteristics

### PSU Response Time
Based on voltage sweep testing with continuous DMM monitoring:

| Transition | Settle Time |
|------------|-------------|
| Initial → 5V | ~464ms |
| 5V → 10V | ~761ms |
| 10V → 15V | ~939ms |
| 15V → 10V | ~643ms |
| 10V → 5V | ~838ms |
| 5V → 18V | ~733ms |
| 18V → 8V | ~887ms |

**Key Findings:**
- **Transition detection**: 130-155ms (DMM detects voltage change)
- **Typical settle time**: 600-950ms (PSU reaches stable output)
- **Recommended delay**: **1 second** after voltage change for reliable readings
- **DMM sampling rate**: ~3 readings/second

### Measurement Accuracy
Voltage sweep test results (PSU set vs DMM reading):

| PSU Setting | DMM Reading | Error |
|-------------|-------------|-------|
| 5.00V | 5.000V | 0.000V |
| 6.00V | 6.001V | +0.001V |
| 7.00V | 7.000V | 0.000V |
| 8.00V | 8.000V | 0.000V |
| 9.00V | 8.997V | -0.003V |
| 10.00V | 9.998V | -0.002V |
| 12.00V | 11.998V | -0.002V |
| 15.00V | 15.000V | 0.000V |
| 18.00V | 17.997V | -0.003V |

**Accuracy**: Within ±0.003V across 5-18V range

## License

MIT License - Free for personal and commercial use.

## Support

For issues with:
- **Hardware**: Contact device manufacturer
- **Software**: Check GitHub issues or create new one
