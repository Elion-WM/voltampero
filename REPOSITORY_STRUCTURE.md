# VoltAmpero Repository Structure

## Root Directory (Production Files)

### Core Application Files
- **voltampero.py** - Main Python application with logging, ramping, and Excel integration
- **psu_korad.py** - Korad KWR102 PSU driver with thread-safe serial communication
- **multimeter_unit.py** - Multimeter driver (for future DMM integration)
- **VoltAmpero.xlsm** - Excel workbook with UI and controls
- **VoltAmpero.bas** - VBA module with macros for PSU control
- **xlwings.conf** - xlwings configuration

### Documentation
- **README.md** - Project overview and main documentation
- **USER_GUIDE.md** - User manual for operating the system
- **QUICK_SETUP.md** - Quick start guide
- **SRS.md** - Software Requirements Specification
- **requirements.txt** - Python dependencies

## Subdirectories

### `/docs/` - Historical Documentation
Contains detailed documentation of fixes and solutions developed during the project:
- COMPLETE_SOLUTION.md
- FINAL_FIX.md
- FINAL_WORKING_SOLUTION.md
- FIX_APPLY_SETTINGS.md
- FIX_LIVE_READINGS_AND_RAMP.md
- FIX_VBA_BUTTON.md
- LIVE_READINGS_SIMPLE_FIX.md
- PROTOCOL_FIX_SUMMARY.md
- QUICK_FIX_COM_PORT.md
- REAL_PSU_FIX.md
- SOLUTION_SUMMARY.md
- EXCEL_SETUP.md

### `/testing/` - Test Scripts and Diagnostics
All test scripts, diagnostic tools, and temporary files used during development:
- Test scripts (test_*.py, test_*.vbs)
- Diagnostic scripts (diagnose_*.py, debug_*.py)
- Fix utilities (fix_*.py, add_*.py, check_*.py)
- Verification scripts (verify_*.py)
- Log files (*.txt, *.csv)
- Batch utilities (*.bat)
- Legacy installers (get-pip.py)

### `/python/` - Python Virtual Environment
Python virtual environment with all dependencies installed.

### `/.git/` - Git Repository
Version control metadata.

### `/__pycache__/` - Python Cache
Compiled Python bytecode (auto-generated).

---

## Key Features Implemented

✅ **PSU Control**
- Connect to Korad KWR102 via USB-to-Serial
- Set voltage and current
- Output ON/OFF control
- Thread-safe serial communication

✅ **Data Logging**
- Configurable intervals (500ms minimum, accurate timing)
- Real-time data collection to Excel Data sheet
- Live readings display in Control sheet
- CSV export functionality

✅ **Voltage Ramping**
- Accurate timing with overhead compensation
- Concurrent operation with data logging
- Configurable start/end voltage and duration
- Progress tracking

✅ **Excel Integration**
- VBA macros for all controls
- xlwings for Python-Excel communication
- Named ranges for all inputs/outputs
- Automatic button assignments

---

## System Requirements

- Windows 10/11
- Python 3.11+
- Excel (with macro support)
- USB-to-Serial driver for PSU
- xlwings, pyserial

---

## Usage

1. Open **VoltAmpero.xlsm** in Excel
2. Set COM port for PSU
3. Click **"Connect PSU"**
4. Configure voltage/current settings
5. Click **"Apply Settings"**
6. Click **"Output ON"** to enable output
7. Use **"Start Logging"** to begin data collection
8. Use **"Start Ramp"** for voltage ramping

---

## Maintenance Notes

- Core files in root are production-ready
- `/testing/` contains development history (can be archived)
- `/docs/` contains detailed fix documentation (reference material)
- Python environment is self-contained in `/python/`

---

*Last updated: 2026-01-08*
