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

This project uses PyMeasure as the standard framework for lab equipment drivers to ensure a consistent interface and easy integration into the Mainframe.

## Installation (No Admin Required)

### 1. Install Python from Microsoft Store

1. Open Microsoft Store
2. Search "Python 3.11"
3. Click Install (no admin needed)

### 2. Install Dependencies

Open Command Prompt (Win+R, type `cmd`, Enter):

```cmd
pip install --user pyserial==3.5 hidapi==0.14.0.post4 xlwings==0.33.17 pymeasure==0.13.0
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
├── requirements.txt       # Pinned Python dependencies (incl. PyMeasure)
├── xlwings.conf.template  # Template; copy to xlwings.conf and adjust paths
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
- Baud: 115200, 8N1
- Commands: VSET1:xx.xx, ISET1:x.xxx, OUT1/0, OCP1/0

### UNI-T UT8804E (USB HID)
- Vendor ID: 0x10c4 (Silicon Labs CP2110)
- Product ID: 0xea80
- Streams data at ~3 readings/sec

## License

MIT License - Free for personal and commercial use.

## Support

For issues with:
- **Hardware**: Contact device manufacturer
- **Software**: Check GitHub issues or create new one
