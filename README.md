# VoltAmpero

**Excel-based Control and Data Logging for Korad KWR102 Power Supply**

> Professional PSU control, automated voltage ramping, and high-precision data logging through an intuitive Excel interface.

---

## 🎯 What is VoltAmpero?

VoltAmpero is a Windows application that provides comprehensive control over laboratory power supplies through Microsoft Excel. It enables researchers, engineers, and hobbyists to:

- **Control PSU settings** directly from Excel (voltage, current, output on/off)
- **Automate voltage ramping** with precise timing and logging
- **Log real-time data** at configurable intervals (500ms - unlimited)
- **Export measurements** to CSV for analysis
- **Monitor live readings** during experiments

The software features thread-safe communication, accurate timing, and a user-friendly Excel interface requiring no programming knowledge.

**Developed by**: Elion-Circular  
**Created with**: Factory.ai assistance  
**License**: Non-commercial use only

---

## ✨ Key Features

### Power Supply Control
- ✅ Set voltage (0-60V) and current limit (0-5A)
- ✅ Output ON/OFF control
- ✅ Real-time voltage/current monitoring
- ✅ Over-current protection (OCP) control
- ✅ Automatic reconnection handling

### Data Logging
- ✅ Configurable logging intervals (500ms minimum)
- ✅ Accurate timing (±50ms precision)
- ✅ Real-time display updates
- ✅ CSV export for analysis
- ✅ Timestamped measurements
- ✅ Records: voltage, current, setpoints

### Voltage Ramping
- ✅ Automated voltage sweeps with accurate timing
- ✅ Configurable start/end voltage and duration
- ✅ Concurrent data logging during ramp
- ✅ Multi-cycle support with ping-pong mode
- ✅ Real-time progress tracking

### Technical Excellence
- ✅ Thread-safe serial communication
- ✅ No conflicts between ramp and logging operations
- ✅ Handles CC/CV mode transitions
- ✅ Robust error handling and reconnection
- ✅ Excel VBA integration via xlwings

---

## 🖥️ Interface

VoltAmpero provides an intuitive Excel-based interface with:
- **Control Tab**: PSU connection, voltage/current settings, output control, live readings
- **Data Tab**: Real-time data logging with timestamps and measurements
- **Named Ranges**: All inputs/outputs clearly labeled in Excel

See [USER_GUIDE.md](USER_GUIDE.md) for detailed interface description and usage examples.

---

## 🔧 Hardware Requirements

### Required Equipment

#### Power Supply
- **Model**: Korad KWR102 (tested with V2.3 firmware)
- **Connection**: USB (appears as virtual COM port)
- **Driver**: CH340 USB-to-Serial driver
- **Protocol**: Custom serial protocol (see docs/KORAD_KWR102_PROTOCOL.md)

#### Multimeter (Optional - Future Support)
- **Model**: UNI-T UT8804E Bench Multimeter
- **Connection**: USB HID (CP2110 bridge)
- **Protocol**: See docs/UNIT_UT8804E_PROTOCOL.md

### Software Requirements

- **OS**: Windows 10/11
- **Python**: 3.11+ (included in setup)
- **Excel**: Microsoft Excel with macro support enabled
- **Drivers**: CH340 USB-to-Serial driver (for PSU)

---

## 📦 Installation

### Quick Start (5 minutes)

1. **Clone or download this repository**
   ```bash
   git clone https://github.com/yourusername/voltampero.git
   cd voltampero
   ```

2. **Install Python dependencies**
   ```bash
   python -m venv python
   python\Scripts\pip install -r requirements.txt
   ```

3. **Open the Excel workbook**
   ```
   VoltAmpero.xlsm
   ```

4. **Enable macros** when prompted

5. **Connect your PSU** and start using!

### Detailed Setup

See [QUICK_SETUP.md](QUICK_SETUP.md) for detailed installation instructions including:
- Python virtual environment setup
- Excel macro security settings
- COM port identification
- Troubleshooting

---

## 🚀 Quick Usage Guide

### Basic PSU Control

1. **Connect**
   - Enter COM port (e.g., COM3) in Excel
   - Click "Connect PSU"

2. **Set Parameters**
   - Enter desired voltage and current limit
   - Click "Apply Settings"

3. **Enable Output**
   - Click "Output ON"
   - PSU now outputs set voltage

### Data Logging

1. **Configure Logging**
   - Set log interval (e.g., 1000ms for 1-second intervals)
   
2. **Start Logging**
   - Click "Start Logging"
   - Data populates in "Data" tab automatically

3. **Export Data**
   - Click "Export CSV"
   - Timestamped CSV file created

### Voltage Ramping

1. **Configure Ramp**
   - Set: Start voltage, End voltage, Duration (seconds)
   
2. **Run Ramp**
   - Click "Apply Settings" (sets starting voltage)
   - Click "Output ON"
   - Click "Start Logging" (to record data)
   - Click "Start Ramp"
   - Voltage automatically sweeps from start to end

3. **Monitor Progress**
   - Watch real-time voltage on PSU display
   - Data collected in Data tab

**For complete usage instructions, see [USER_GUIDE.md](USER_GUIDE.md)**

---

## 📚 Documentation

### User Documentation
- **[USER_GUIDE.md](USER_GUIDE.md)** - Complete user manual
- **[QUICK_SETUP.md](QUICK_SETUP.md)** - Fast installation guide
- **[SRS.md](SRS.md)** - Software requirements specification

### Protocol Documentation (Reusable!)
- **[Korad KWR102 Protocol](docs/KORAD_KWR102_PROTOCOL.md)** - Complete PSU communication protocol
- **[UNI-T UT8804E Protocol](docs/UNIT_UT8804E_PROTOCOL.md)** - Complete DMM communication protocol
- **[Protocols Index](docs/PROTOCOLS_INDEX.md)** - Quick reference and comparison

### Development Documentation
- **[REPOSITORY_STRUCTURE.md](REPOSITORY_STRUCTURE.md)** - Project organization
- **[/docs/](docs/)** - Detailed fix documentation and development history
- **[/testing/](testing/)** - Test scripts and diagnostic tools (89 files)

---

## 🏗️ Architecture

```
VoltAmpero
│
├── VoltAmpero.xlsm          Excel UI (buttons, controls, display)
├── VoltAmpero.bas           VBA macros (Excel → Python bridge)
│
├── voltampero.py            Main application logic
│   ├── Logging system       (thread-safe, configurable intervals)
│   ├── Ramp controller      (accurate timing)
│   └── Excel integration    (xlwings communication)
│
├── psu_korad.py             Korad KWR102 driver
│   ├── Serial protocol      (thread-safe with Lock)
│   ├── Command formatting   (VSET:, ISET:, OUT:)
│   └── Query optimization   (fast timeouts)
│
└── multimeter_unit.py       UNI-T UT8804E driver (future)
    ├── HID communication    (USB, no driver needed)
    └── CP2110 UART bridge   (initialization)
```

---

## 🔬 Technical Highlights

### Thread-Safe Communication
- Threading locks prevent conflicts between ramp (writing) and logging (reading)
- Concurrent operations work flawlessly

### Accurate Timing
- Logging intervals: ±50ms accuracy (e.g., 1000ms → 950-1050ms actual)
- Ramp duration: ±0.5s accuracy over 240 seconds
- Overhead compensation ensures target intervals are met

### Protocol Reverse Engineering
- Korad KWR102 V2.3 protocol fully documented (differs from other models!)
- Discovered via USB packet capture (Wireshark)
- Uses `VSET:` format (not `VSET1:`), requires `\r` terminator
- RTS/DTR control lines must be set

### Excel Integration
- xlwings for Python-Excel communication
- VBA macros for user interaction
- Named ranges for all inputs/outputs
- No programming knowledge required for end users

---

## 📋 System Requirements

| Component | Requirement |
|-----------|------------|
| **Operating System** | Windows 10/11 |
| **Python** | 3.11+ |
| **Excel** | Microsoft Excel (2016 or later) |
| **PSU** | Korad KWR102 (V2.3 tested) |
| **USB Driver** | CH340 USB-to-Serial |
| **Dependencies** | xlwings, pyserial, hidapi |

---

## 🛠️ Dependencies

```txt
xlwings>=0.30.0
pyserial>=3.5
hidapi>=0.14.0
```

Install with:
```bash
pip install -r requirements.txt
```

---

## 📖 Usage Example

```python
# Python API (for advanced users)
from psu_korad import KoradKWR102

# Connect to PSU
psu = KoradKWR102(port='COM3')
psu.connect()

# Set 12V, 1A limit
psu.set_voltage(12.0)
psu.set_current(1.0)
psu.output_on()

# Read actual values
voltage, current = psu.get_readings()
print(f"Output: {voltage}V, {current}A")

# Ramp voltage
from voltampero import VoltAmpero
ctrl = VoltAmpero()
ctrl.connect_psu('COM3')
ctrl.start_logging(interval_ms=1000)
ctrl.start_ramp(start_v=5.0, end_v=15.0, duration_s=60)
```

**Most users will use the Excel interface instead!**

---

## 🐛 Troubleshooting

### PSU Not Responding
- Verify COM port in Device Manager
- Ensure CH340 driver installed
- Check PSU is powered on
- Try different USB port

### Settings Not Applying
- Confirm "Connected" status in Excel
- Check output is ON
- Verify voltage/current values are within PSU limits

### Data Shows Zeros
- Ensure output is ON before starting ramp/logging
- Check PSU connection status
- Verify current limit isn't too low (CC mode)

### Logging Interval Inaccurate
- Minimum realistic interval: 500ms (PSU query overhead)
- For <800ms intervals, expect some variation
- 1000ms+ intervals are highly accurate (±50ms)

**For more troubleshooting, see [USER_GUIDE.md](USER_GUIDE.md)**

---

## 🤝 Contributing

Contributions are welcome! This project is open for non-commercial use.

### How to Contribute:
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Test your changes thoroughly
4. Commit with clear messages
5. Push and create a Pull Request

### Contribution Ideas:
- Support for other Korad PSU models
- Additional DMM support
- Unit tests
- UI improvements
- Bug fixes
- Documentation improvements

**By contributing, you agree that your contributions will be under the same non-commercial license.**

---

## 📄 License

Copyright © 2026 Elion-Circular

This software is licensed for **non-commercial use only** (personal, educational, research, non-profit).

**Commercial use is STRICTLY PROHIBITED.**

### Attribution Required:
This software was developed with the assistance of [Factory.ai](https://factory.ai).

All copies must include:
- The LICENSE file
- Factory.ai attribution
- Original copyright notice

**See [LICENSE](LICENSE) for full terms. See [docs/LICENSE_EXPLAINED.md](docs/LICENSE_EXPLAINED.md) for plain-English explanation.**

---

## ⚠️ Safety Disclaimer

**This software controls electrical power equipment. Users are responsible for:**
- Ensuring proper electrical safety measures
- Understanding equipment specifications and limitations
- Following applicable safety regulations
- Any damages or injuries resulting from use

**Elion-Circular and Factory.ai assume no liability for equipment damage, personal injury, or any consequences of using this software.**

---

## 🙏 Acknowledgments

- **Factory.ai** - AI-powered development assistance
- **Korad** - KWR102 power supply hardware
- **UNI-T** - UT8804E multimeter hardware (future support)
- **xlwings** - Excellent Python-Excel integration library
- **Open-source community** - Sigrok wiki and protocol documentation efforts

---

## 📞 Support

- **Documentation**: See `/docs/` folder
- **Issues**: Use GitHub Issues for bug reports
- **Protocol questions**: See protocol documentation in `/docs/`
- **Usage help**: See USER_GUIDE.md

---

## 🗺️ Roadmap

### Current Version (v1.0)
- ✅ PSU control via Korad KWR102
- ✅ Data logging with accurate intervals
- ✅ Voltage ramping with precise timing
- ✅ Excel UI with VBA macros

### Future Enhancements
- ⏳ DMM integration (UNI-T UT8804E)
- ⏳ Support for other Korad models
- ⏳ Advanced analysis features
- ⏳ Automated testing framework
- ⏳ Cross-platform support (Linux, macOS)

---

## 📊 Project Stats

- **Lines of code**: ~1,500 (Python) + ~500 (VBA)
- **Documentation**: 20+ guides and references
- **Test coverage**: 89 test scripts in /testing/
- **Protocol docs**: 2 complete hardware protocols
- **Development time**: 3 days (with Factory.ai)

---

## 🌟 Why VoltAmpero?

### For Researchers
- Automated data collection for experiments
- Reproducible voltage sweeps
- Timestamped measurements for analysis
- Easy CSV export for papers/reports

### For Engineers
- Rapid PSU testing and characterization
- Automated stress testing
- Precise control for calibration
- Professional data logging

### For Hobbyists
- Simple Excel interface (no coding needed)
- Learn power electronics through experimentation
- Document your projects with data
- Reusable protocol documentation

### For Developers
- Complete protocol documentation (save hours of reverse engineering!)
- Thread-safe serial communication example
- Excel-Python integration reference
- Well-organized, documented codebase

---

## 🔗 Links

- **Documentation**: [/docs/](docs/)
- **Protocol Specs**: [docs/PROTOCOLS_INDEX.md](docs/PROTOCOLS_INDEX.md)
- **User Guide**: [USER_GUIDE.md](USER_GUIDE.md)
- **License**: [LICENSE](LICENSE)
- **Factory.ai**: https://factory.ai

---

## 📈 Repository Statistics

![Python](https://img.shields.io/badge/python-3.11+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/license-Non--Commercial-red.svg)
![Status](https://img.shields.io/badge/status-Production%20Ready-brightgreen.svg)

---

## 💡 Quick Example

**Typical workflow:**

1. Open `VoltAmpero.xlsm`
2. Enter COM port (e.g., COM3)
3. Click "Connect PSU"
4. Set voltage to 12V, current to 1A
5. Click "Apply Settings"
6. Click "Output ON"
7. Set log interval to 1000ms
8. Click "Start Logging"
9. Watch live data populate!

**That's it!** No coding, no complex setup.

---

## 🧪 Testing

The project includes comprehensive testing:
- 89 test scripts in `/testing/`
- Protocol verification tests
- Timing accuracy tests
- Thread safety tests
- Excel integration tests

Run tests:
```bash
python testing/test_apply_settings_debug.py
python testing/test_interval_verification.py
```

---

## 📜 Version History

### v1.0.0 (2026-01-08)
- Initial release
- Full Korad KWR102 V2.3 support
- Data logging with accurate intervals
- Voltage ramping with precise timing
- Thread-safe serial communication
- Comprehensive protocol documentation
- Excel UI with VBA macros

**See [CHANGELOG.md](CHANGELOG.md) for detailed history** *(coming soon)*

---

## 🌐 Protocol Documentation

One of the unique features of VoltAmpero is **complete protocol documentation** for the hardware:

### [Korad KWR102 Protocol](docs/KORAD_KWR102_PROTOCOL.md)
- Complete command set with examples
- Hardware connection details
- Thread-safe implementation patterns
- Timing and performance optimization
- Comparison with other Korad models
- **Save hours of Wireshark analysis!**

### [UNI-T UT8804E Protocol](docs/UNIT_UT8804E_PROTOCOL.md)
- USB HID communication details
- CP2110 bridge initialization
- Binary packet format
- Mode and range mappings
- **Complete reverse-engineered protocol!**

**These protocols can be reused in your own projects!**

---

## 🎓 Learning Resources

### For Understanding the Code:
1. Start with [REPOSITORY_STRUCTURE.md](REPOSITORY_STRUCTURE.md)
2. Read protocol docs to understand hardware communication
3. Review `psu_korad.py` for serial communication patterns
4. Study `voltampero.py` for threading and timing logic
5. Check `/docs/` for development history

### For Using the Software:
1. Read [QUICK_SETUP.md](QUICK_SETUP.md)
2. Follow [USER_GUIDE.md](USER_GUIDE.md)
3. Experiment with different settings
4. Export and analyze your data

---

## 🏭 About Elion-Circular

VoltAmpero is developed by Elion-Circular for research and educational purposes.

**Mission**: Provide accessible tools for power electronics experimentation and data collection.

---

## 🤖 About Factory.ai

This project was developed with significant assistance from [Factory.ai](https://factory.ai), an AI-powered software development platform.

**Factory.ai helped with:**
- Protocol reverse engineering
- Code generation and debugging
- Documentation creation
- Architecture design
- Performance optimization

Learn more: https://factory.ai

---

## ⚖️ Legal

### Copyright
© 2026 Elion-Circular. All rights reserved.

### License Summary
- ✅ **Allowed**: Personal, educational, research, non-profit use
- ❌ **Prohibited**: Commercial use (strictly enforced)
- 📝 **Required**: Factory.ai attribution in all copies

### Disclaimer
This software controls electrical equipment. Use at your own risk. The authors assume no liability for damages or injuries.

**Full license**: [LICENSE](LICENSE)  
**Plain English explanation**: [docs/LICENSE_EXPLAINED.md](docs/LICENSE_EXPLAINED.md)

---

## 📧 Contact

- **Issues**: Use GitHub Issues for bug reports and feature requests
- **Questions**: See documentation first, then open an issue
- **Commercial use**: Not available (strictly prohibited)

---

## ⭐ Star This Repository

If you find VoltAmpero useful for your research or projects, please star this repository! It helps others discover the project and the protocol documentation.

---

**Built with ❤️ by Elion-Circular using Factory.ai**

*Making power electronics experimentation accessible to everyone.*
