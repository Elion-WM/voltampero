# Software Requirements Specification (SRS)
## VoltAmpero - Lab Instrument Control Software

**Version:** 1.0  
**Date:** December 2024  
**Author:** Elion-WM  

---

## 1. Introduction

### 1.1 Purpose
This document specifies the software requirements for VoltAmpero, a laboratory instrument control system that integrates a Korad KWR102 power supply and UNI-T UT8804E multimeter with Microsoft Excel for data acquisition and control.

### 1.2 Scope
VoltAmpero provides:
- Real-time control of laboratory power supply
- Automated data logging from PSU and multimeter
- Voltage ramping with configurable cycles
- Excel-based user interface with live charts
- CSV data export functionality

### 1.3 Definitions and Acronyms
| Term | Definition |
|------|------------|
| PSU | Power Supply Unit (Korad KWR102) |
| DMM | Digital Multimeter (UNI-T UT8804E) |
| OCP | Over Current Protection |
| OVP | Over Voltage Protection |
| HID | Human Interface Device (USB protocol) |
| VBA | Visual Basic for Applications |

### 1.4 References
- Korad KWR102 User Manual
- UNI-T UT8804E User Manual
- xlwings Documentation (https://docs.xlwings.org)

---

## 2. Overall Description

### 2.1 Product Perspective
VoltAmpero is a standalone desktop application that bridges laboratory hardware with Excel spreadsheets using Python and xlwings. It operates without requiring administrator privileges on Windows systems.

### 2.2 Product Functions
- **F1**: Power Supply Control
- **F2**: Multimeter Data Acquisition
- **F3**: Data Logging
- **F4**: Voltage Ramping
- **F5**: Excel Integration
- **F6**: CSV Export

### 2.3 User Classes and Characteristics
| User Class | Description |
|------------|-------------|
| Lab Technician | Primary user performing measurements and experiments |
| Engineer | Advanced user configuring voltage ramps and data analysis |
| Student | Basic user running pre-configured tests |

### 2.4 Operating Environment
- **OS**: Windows 10/11 (64-bit)
- **Software**: Python 3.8+, Microsoft Excel (with macros enabled)
- **Hardware**: USB ports for PSU and DMM connections

### 2.5 Design and Implementation Constraints
- No administrator rights required for installation
- Must use xlwings for Excel integration
- Serial communication for PSU (115200 baud)
- USB HID protocol for DMM

### 2.6 Assumptions and Dependencies
- User has Microsoft Excel installed
- USB cables provided with instruments
- Python available via Microsoft Store

---

## 3. Specific Requirements

### 3.1 Functional Requirements

#### 3.1.1 Power Supply Control (F1)

| ID | Requirement | Priority |
|----|-------------|----------|
| F1.1 | System shall connect to Korad KWR102 via specified COM port | High |
| F1.2 | System shall set output voltage (0-30V, 0.01V resolution) | High |
| F1.3 | System shall set current limit (0-5A, 0.001A resolution) | High |
| F1.4 | System shall turn output ON/OFF | High |
| F1.5 | System shall enable/disable OCP | Medium |
| F1.6 | System shall enable/disable OVP | Medium |
| F1.7 | System shall read actual voltage and current | High |
| F1.8 | System shall display connection status | High |
| F1.9 | System shall list available COM ports | Medium |

#### 3.1.2 Multimeter Data Acquisition (F2)

| ID | Requirement | Priority |
|----|-------------|----------|
| F2.1 | System shall auto-detect UNI-T UT8804E via USB HID | High |
| F2.2 | System shall read measurement value | High |
| F2.3 | System shall read measurement unit | High |
| F2.4 | System shall read measurement mode (V, A, Ohm, etc.) | High |
| F2.5 | System shall support ~3 readings per second | High |
| F2.6 | System shall display connection status | High |

#### 3.1.3 Data Logging (F3)

| ID | Requirement | Priority |
|----|-------------|----------|
| F3.1 | System shall log timestamp for each reading | High |
| F3.2 | System shall log elapsed time in seconds | High |
| F3.3 | System shall log PSU voltage and current | High |
| F3.4 | System shall log PSU setpoint values | Medium |
| F3.5 | System shall log DMM value, unit, and mode | High |
| F3.6 | System shall support configurable log interval (default 300ms) | Medium |
| F3.7 | System shall start/stop logging on command | High |
| F3.8 | System shall clear logged data on command | Medium |

#### 3.1.4 Voltage Ramping (F4)

| ID | Requirement | Priority |
|----|-------------|----------|
| F4.1 | System shall ramp voltage from start to end value | High |
| F4.2 | System shall complete ramp within specified duration | High |
| F4.3 | System shall support multiple cycles (0 = infinite) | Medium |
| F4.4 | System shall support delay between cycles | Medium |
| F4.5 | System shall support ping-pong mode (alternating direction) | Medium |
| F4.6 | System shall display current cycle number | Medium |
| F4.7 | System shall display ramp progress (0-100%) | Medium |
| F4.8 | System shall start/stop/pause ramp on command | High |

#### 3.1.5 Excel Integration (F5)

| ID | Requirement | Priority |
|----|-------------|----------|
| F5.1 | System shall integrate with Excel via xlwings | High |
| F5.2 | System shall update Control sheet with live readings | High |
| F5.3 | System shall write log data to Data sheet | High |
| F5.4 | System shall read settings from named ranges | High |
| F5.5 | System shall provide VBA macros for all functions | High |
| F5.6 | System shall support button-triggered actions | High |

#### 3.1.6 CSV Export (F6)

| ID | Requirement | Priority |
|----|-------------|----------|
| F6.1 | System shall export logged data to CSV file | High |
| F6.2 | System shall include timestamp in export filename | Medium |
| F6.3 | System shall save CSV in workbook directory | Medium |
| F6.4 | System shall display export status message | Low |

### 3.2 Non-Functional Requirements

#### 3.2.1 Performance

| ID | Requirement |
|----|-------------|
| NFR1 | Logging shall capture readings at minimum 3 Hz |
| NFR2 | Excel updates shall not freeze the UI |
| NFR3 | PSU commands shall execute within 100ms |

#### 3.2.2 Reliability

| ID | Requirement |
|----|-------------|
| NFR4 | System shall handle device disconnection gracefully |
| NFR5 | System shall not lose logged data on export failure |
| NFR6 | System shall recover from communication errors |

#### 3.2.3 Usability

| ID | Requirement |
|----|-------------|
| NFR7 | Installation shall not require admin rights |
| NFR8 | All functions accessible via Excel buttons |
| NFR9 | Status indicators shall show device states |

#### 3.2.4 Portability

| ID | Requirement |
|----|-------------|
| NFR10 | System shall run on Windows 10 and 11 |
| NFR11 | System shall support Python 3.8 through 3.12 |

#### 3.2.5 Security

| ID | Requirement |
|----|-------------|
| NFR12 | No sensitive data stored in configuration |
| NFR13 | No network access required |

---

## 4. System Interfaces

### 4.1 Hardware Interfaces

#### 4.1.1 Korad KWR102 Power Supply
| Parameter | Value |
|-----------|-------|
| Interface | USB-Serial (Virtual COM Port) |
| Baud Rate | 115200 |
| Data Bits | 8 |
| Parity | None |
| Stop Bits | 1 |
| Commands | SCPI-like text protocol |

**Command Set:**
| Command | Description |
|---------|-------------|
| `*IDN?` | Query device identification |
| `VSET1:xx.xx` | Set voltage (volts) |
| `ISET1:x.xxx` | Set current (amps) |
| `VOUT1?` | Query output voltage |
| `IOUT1?` | Query output current |
| `OUT1` | Turn output ON |
| `OUT0` | Turn output OFF |
| `OCP1` | Enable OCP |
| `OCP0` | Disable OCP |

#### 4.1.2 UNI-T UT8804E Multimeter
| Parameter | Value |
|-----------|-------|
| Interface | USB HID |
| Vendor ID | 0x10c4 (Silicon Labs) |
| Product ID | 0xea80 |
| Data Rate | ~3 readings/second |
| Protocol | Binary packet stream |

### 4.2 Software Interfaces

#### 4.2.1 Python Dependencies
| Package | Version | Purpose |
|---------|---------|---------|
| pyserial | >=3.5 | Serial communication with PSU |
| hidapi | >=0.14 | USB HID communication with DMM |
| xlwings | >=0.30 | Excel integration |

#### 4.2.2 Excel Named Ranges
| Name | Cell | Description |
|------|------|-------------|
| PSUPort | B3 | COM port for PSU |
| PSUStatus | D3 | PSU connection status |
| DMMStatus | D4 | DMM connection status |
| LoggingStatus | D5 | Logging active/stopped |
| RampStatus | D6 | Ramp running/stopped |
| LogInterval | B8 | Logging interval (ms) |
| LiveVoltage | B11 | Current PSU voltage |
| LiveCurrent | B12 | Current PSU current |
| LiveDMM | B13 | Current DMM reading |
| SetVoltage | B16 | Voltage setpoint |
| SetCurrent | B17 | Current limit setpoint |
| OCPEnabled | B18 | OCP on/off |
| RampStartV | B21 | Ramp start voltage |
| RampEndV | B22 | Ramp end voltage |
| RampDuration | B23 | Ramp duration (seconds) |
| RampCycles | B24 | Number of cycles |
| RampDelay | B25 | Delay between cycles |
| RampPingPong | B26 | Ping-pong mode |

---

## 5. Data Requirements

### 5.1 Log Entry Structure
| Field | Type | Format |
|-------|------|--------|
| Timestamp | DateTime | YYYY-MM-DD HH:MM:SS.mmm |
| Elapsed_s | Float | Seconds with 3 decimals |
| PSU_Voltage_V | Float | Volts with 4 decimals |
| PSU_Current_A | Float | Amps with 4 decimals |
| PSU_Setpoint_V | Float | Volts with 2 decimals |
| PSU_Setpoint_A | Float | Amps with 3 decimals |
| DMM_Value | Float | 6 decimal precision |
| DMM_Unit | String | V, mV, A, mA, Ohm, etc. |
| DMM_Mode | String | DC_V, AC_V, DC_A, etc. |

### 5.2 CSV Export Format
- Comma-separated values
- UTF-8 encoding
- Header row included
- Filename: `voltampero_log_YYYYMMDD_HHMMSS.csv`

---

## 6. Appendices

### Appendix A: File Structure
```
voltampero/
├── README.md              # Project documentation
├── SRS.md                 # This document
├── EXCEL_SETUP.md         # Excel configuration guide
├── QUICK_SETUP.md         # Quick start guide
├── requirements.txt       # Python dependencies
├── psu_korad.py          # Korad KWR102 driver
├── multimeter_unit.py    # UNI-T UT8804E driver
├── voltampero.py         # Main controller
├── VoltAmpero.bas        # VBA module for Excel
└── xlwings.conf          # xlwings configuration
```

### Appendix B: Revision History
| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | Dec 2024 | Elion-WM | Initial release |
