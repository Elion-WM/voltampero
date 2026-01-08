# Excel Workbook Setup Guide for VoltAmpero

## Step 1: Create New Excel Workbook

1. Open Excel
2. Save as **VoltAmpero.xlsm** (Excel Macro-Enabled Workbook)
3. Enable macros when prompted

## Step 2: Create Sheets

Create two sheets named exactly:
- **Control** (main interface)
- **Data** (logging data)

---

## Step 3: Set Up "Control" Sheet

### Named Ranges (Important!)
Go to **Formulas > Name Manager** and create these named ranges:

| Name | Cell | Description |
|------|------|-------------|
| PSUPort | B3 | COM port (e.g., COM3) |
| PSUStatus | D3 | PSU connection status |
| DMMStatus | D4 | DMM connection status |
| LoggingStatus | D5 | Logging status |
| RampStatus | D6 | Ramp status |
| LogInterval | B8 | Log interval in ms (default: 300) |
| LiveVoltage | B11 | Live PSU voltage display |
| LiveCurrent | B12 | Live PSU current display |
| LiveDMM | B13 | Live DMM reading display |
| SetVoltage | B16 | Voltage setpoint input |
| SetCurrent | B17 | Current limit input |
| OCPEnabled | B18 | OCP checkbox (TRUE/FALSE) |
| RampStartV | B21 | Ramp start voltage |
| RampEndV | B22 | Ramp end voltage |
| RampDuration | B23 | Ramp duration (seconds) |
| RampCycles | B24 | Number of cycles (0=infinite) |
| RampDelay | B25 | Delay between cycles (seconds) |
| RampPingPong | B26 | Ping-pong mode (TRUE/FALSE) |
| RampCycle | D21 | Current cycle display |
| RampVoltage | D22 | Current ramp voltage |
| RampProgress | D23 | Progress (0-1 for progress bar) |
| ExportStatus | B30 | Export status message |

### Suggested Layout for Control Sheet

```
Row 1: [Title] VoltAmpero - Lab Instrument Control
Row 2: ─────────────────────────────────────────────
Row 3: PSU Port:     [B3: COM3]      Status: [D3: Disconnected]
Row 4: DMM:          [auto-detect]   Status: [D4: Disconnected]
Row 5: Logging:                      Status: [D5: Stopped]
Row 6: Ramp:                         Status: [D6: Stopped]
Row 7: ─────────────────────────────────────────────
Row 8: Log Interval (ms): [B8: 300]
Row 9:
Row 10: === LIVE READINGS ===
Row 11: Voltage (V):  [B11: 0.000]
Row 12: Current (A):  [B12: 0.000]
Row 13: DMM:          [B13: --- ---]
Row 14:
Row 15: === PSU CONTROL ===
Row 16: Set Voltage (V): [B16: 5.00]
Row 17: Set Current (A): [B17: 1.000]
Row 18: OCP Enabled:     [B18: checkbox]
Row 19:
Row 20: === VOLTAGE RAMP ===
Row 21: Start (V):    [B21: 0]       Cycle: [D21: 0/0]
Row 22: End (V):      [B22: 12]      Voltage: [D22: 0]
Row 23: Duration (s): [B23: 60]      Progress: [D23: progress bar]
Row 24: Cycles:       [B24: 1]
Row 25: Delay (s):    [B25: 0]
Row 26: Ping-Pong:    [B26: checkbox]
Row 27:
Row 28: === DATA EXPORT ===
Row 29:
Row 30: [B30: Export status message]
```

### Add Buttons
Insert buttons (Developer > Insert > Button) with these assignments:

| Button Text | Macro Name |
|-------------|------------|
| Connect PSU | ConnectPSU |
| Connect DMM | ConnectDMM |
| Disconnect All | DisconnectAll |
| Output ON | OutputOn |
| Output OFF | OutputOff |
| Apply Settings | ApplySettings |
| Start Logging | StartLogging |
| Stop Logging | StopLogging |
| Start Ramp | StartRamp |
| Stop Ramp | StopRamp |
| Pause Ramp | PauseRamp |
| Export CSV | ExportCSV |
| Clear Data | ClearData |
| Test (Simulated) | InitSimulated |

---

## Step 4: Set Up "Data" Sheet

### Headers (Row 1)
```
A1: Timestamp
B1: Elapsed_s
C1: PSU_Voltage_V
D1: PSU_Current_A
E1: PSU_Setpoint_V
F1: PSU_Setpoint_A
G1: DMM_Value
H1: DMM_Unit
I1: DMM_Mode
```

### Optional: Create Chart
1. Select data range A1:D100 (will expand)
2. Insert > Line Chart
3. Set up dual axis for Voltage and Current
4. Link chart to Data table

---

## Step 5: Add VBA Code

Press **Alt+F11** to open VBA Editor, then:

### 1. Add xlwings Reference
- Tools > References
- Check "xlwings" (if available)
- Or it will auto-configure

### 2. Insert Module
- Right-click project > Insert > Module
- Paste this code:

```vba
Option Explicit

' VoltAmpero Excel VBA Interface
' Calls Python functions via xlwings

Sub ConnectPSU()
    Dim port As String
    port = Range("PSUPort").Value
    If port = "" Then
        MsgBox "Please enter COM port (e.g., COM3)", vbExclamation
        Exit Sub
    End If
    RunPython "from voltampero import get_controller; c=get_controller(); c.attach_excel(); c.connect_psu('" & port & "')"
End Sub

Sub ConnectDMM()
    RunPython "from voltampero import get_controller; c=get_controller(); c.attach_excel(); c.connect_dmm()"
End Sub

Sub DisconnectAll()
    RunPython "from voltampero import va_disconnect_all; va_disconnect_all()"
End Sub

Sub OutputOn()
    RunPython "from voltampero import va_output_on; va_output_on()"
End Sub

Sub OutputOff()
    RunPython "from voltampero import va_output_off; va_output_off()"
End Sub

Sub ApplySettings()
    Dim voltage As Double, current As Double, ocp As Boolean
    voltage = Range("SetVoltage").Value
    current = Range("SetCurrent").Value
    ocp = Range("OCPEnabled").Value
    
    RunPython "from voltampero import get_controller; c=get_controller(); c.set_voltage(" & voltage & "); c.set_current(" & current & "); c.set_ocp(" & IIf(ocp, "True", "False") & ")"
End Sub

Sub StartLogging()
    RunPython "from voltampero import va_start_logging; va_start_logging()"
End Sub

Sub StopLogging()
    RunPython "from voltampero import va_stop_logging; va_stop_logging()"
End Sub

Sub StartRamp()
    RunPython "from voltampero import va_start_ramp; va_start_ramp()"
End Sub

Sub StopRamp()
    RunPython "from voltampero import va_stop_ramp; va_stop_ramp()"
End Sub

Sub PauseRamp()
    RunPython "from voltampero import va_pause_ramp; va_pause_ramp()"
End Sub

Sub ExportCSV()
    RunPython "from voltampero import va_export_csv; va_export_csv()"
End Sub

Sub ClearData()
    RunPython "from voltampero import va_clear_data; va_clear_data()"
End Sub

Sub InitSimulated()
    ' Initialize with simulated devices for testing
    RunPython "from voltampero import va_init_simulated; va_init_simulated()"
    MsgBox "Simulated mode initialized. PSU and DMM connected.", vbInformation
End Sub

Sub RefreshReadings()
    ' Can be called by a timer or button
    RunPython "from voltampero import get_controller; c=get_controller(); c.attach_excel(); r=c._capture_reading(); c._write_entry_to_excel(r) if r else None"
End Sub

' Auto-refresh timer (optional)
Dim NextRefresh As Date

Sub StartAutoRefresh()
    NextRefresh = Now + TimeSerial(0, 0, 1)
    Application.OnTime NextRefresh, "AutoRefreshTick"
End Sub

Sub StopAutoRefresh()
    On Error Resume Next
    Application.OnTime NextRefresh, "AutoRefreshTick", , False
End Sub

Sub AutoRefreshTick()
    RefreshReadings
    StartAutoRefresh
End Sub
```

---

## Step 6: Configure xlwings

### xlwings.conf file
Create a file named `xlwings.conf` in the same folder as your Excel file:

```ini
[xlwings]
PYTHONPATH=C:\Users\User\OneDrive - ELION\Pulpit\voltampero
INTERPRETER=python
```

Adjust PYTHONPATH to your actual folder path.

### Alternative: Use xlwings addin
1. Run in command prompt: `xlwings addin install`
2. This adds an xlwings ribbon tab to Excel

---

## Step 7: First Run Test

1. Open VoltAmpero.xlsm
2. Click "Test (Simulated)" button
3. If successful, you'll see "Connected" status
4. Click "Start Logging" - data should appear in Data sheet
5. Click "Stop Logging" after a few seconds
6. Click "Export CSV" to save data

---

## Troubleshooting

### "Python not found"
- Install Python from Microsoft Store (no admin needed)
- Or set full path in xlwings.conf: `INTERPRETER=C:\Users\YourName\AppData\Local\Programs\Python\Python311\python.exe`

### "Module not found"
- Make sure PYTHONPATH in xlwings.conf points to the voltampero folder
- Install dependencies: `pip install --user pyserial hidapi xlwings`

### Macros disabled
- File > Options > Trust Center > Trust Center Settings
- Enable macros for this workbook

### HID device not found (multimeter)
- Install hidapi: `pip install hidapi`
- On Windows, may need to copy `hidapi.dll` to script folder
