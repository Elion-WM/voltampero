Attribute VB_Name = "VoltAmpero"
Option Explicit

' VoltAmpero Excel VBA Interface
' Calls Python functions via xlwings
' Import this module into your Excel workbook

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
    RunPython "import voltampero; voltampero.va_init_simulated()"
    MsgBox "Simulated mode initialized. PSU and DMM connected.", vbInformation
End Sub

Sub RefreshReadings()
    ' Drain up to 200 queued entries from logging thread to Excel (main thread safe)
    RunPython "from voltampero import va_drain_queue; va_drain_queue(200)"
End Sub

' ========== Auto-refresh timer (optional) ==========
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

' ========== Sheet Setup Helper ==========
Sub SetupWorkbook()
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    
    ' Create Control sheet if not exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Control")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Control"
    End If
    On Error GoTo 0
    
    ' Create Data sheet if not exists
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("Data")
    If dataWs Is Nothing Then
        Set dataWs = ThisWorkbook.Sheets.Add(After:=ws)
        dataWs.Name = "Data"
    End If
    On Error GoTo 0
    
    ' Setup Control sheet
    With ws
        ' Title
        .Range("A1").Value = "VoltAmpero - Lab Instrument Control"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        ' Connection section
        .Range("A3").Value = "PSU Port:"
        .Range("B3").Value = "COM3"
        .Range("C3").Value = "Status:"
        .Range("D3").Value = "Disconnected"
        
        .Range("A4").Value = "DMM:"
        .Range("B4").Value = "(auto-detect)"
        .Range("C4").Value = "Status:"
        .Range("D4").Value = "Disconnected"
        
        .Range("A5").Value = "Logging:"
        .Range("C5").Value = "Status:"
        .Range("D5").Value = "Stopped"
        
        .Range("A6").Value = "Ramp:"
        .Range("C6").Value = "Status:"
        .Range("D6").Value = "Stopped"
        
        .Range("A8").Value = "Log Interval (ms):"
        .Range("B8").Value = 300
        
        ' Live readings
        .Range("A10").Value = "=== LIVE READINGS ==="
        .Range("A10").Font.Bold = True
        .Range("A11").Value = "Voltage (V):"
        .Range("B11").Value = 0
        .Range("A12").Value = "Current (A):"
        .Range("B12").Value = 0
        .Range("A13").Value = "DMM:"
        .Range("B13").Value = "--- ---"
        
        ' PSU Control
        .Range("A15").Value = "=== PSU CONTROL ==="
        .Range("A15").Font.Bold = True
        .Range("A16").Value = "Set Voltage (V):"
        .Range("B16").Value = 5
        .Range("A17").Value = "Set Current (A):"
        .Range("B17").Value = 1
        .Range("A18").Value = "OCP Enabled:"
        .Range("B18").Value = False
        
        ' Voltage Ramp
        .Range("A20").Value = "=== VOLTAGE RAMP ==="
        .Range("A20").Font.Bold = True
        .Range("A21").Value = "Start (V):"
        .Range("B21").Value = 0
        .Range("C21").Value = "Cycle:"
        .Range("D21").Value = "0/0"
        .Range("A22").Value = "End (V):"
        .Range("B22").Value = 12
        .Range("C22").Value = "Voltage:"
        .Range("D22").Value = 0
        .Range("A23").Value = "Duration (s):"
        .Range("B23").Value = 60
        .Range("C23").Value = "Progress:"
        .Range("D23").Value = 0
        .Range("A24").Value = "Cycles:"
        .Range("B24").Value = 1
        .Range("A25").Value = "Delay (s):"
        .Range("B25").Value = 0
        .Range("A26").Value = "Ping-Pong:"
        .Range("B26").Value = False
        
        ' Export
        .Range("A28").Value = "=== DATA EXPORT ==="
        .Range("A28").Font.Bold = True
        .Range("A30").Value = "Status:"
        .Range("B30").Value = ""
        
        ' Column widths
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 15
    End With
    
    ' Setup Data sheet headers
    With dataWs
        .Range("A1").Value = "Timestamp"
        .Range("B1").Value = "Elapsed_s"
        .Range("C1").Value = "PSU_Voltage_V"
        .Range("D1").Value = "PSU_Current_A"
        .Range("E1").Value = "PSU_Setpoint_V"
        .Range("F1").Value = "PSU_Setpoint_A"
        .Range("G1").Value = "DMM_Value"
        .Range("H1").Value = "DMM_Unit"
        .Range("I1").Value = "DMM_Mode"
        .Range("A1:I1").Font.Bold = True
    End With
    
    ' Create Named Ranges
    Call CreateNamedRanges
    
    MsgBox "Workbook setup complete! Now add buttons for each macro.", vbInformation
End Sub

Sub CreateNamedRanges()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    On Error Resume Next
    
    ' Delete existing names first
    Dim nm As Name
    For Each nm In wb.Names
        If Left(nm.Name, 3) <> "_xl" Then nm.Delete
    Next nm
    
    ' Control sheet named ranges
    wb.Names.Add Name:="PSUPort", RefersTo:="=Control!$B$3"
    wb.Names.Add Name:="PSUStatus", RefersTo:="=Control!$D$3"
    wb.Names.Add Name:="DMMStatus", RefersTo:="=Control!$D$4"
    wb.Names.Add Name:="LoggingStatus", RefersTo:="=Control!$D$5"
    wb.Names.Add Name:="RampStatus", RefersTo:="=Control!$D$6"
    wb.Names.Add Name:="LogInterval", RefersTo:="=Control!$B$8"
    wb.Names.Add Name:="LiveVoltage", RefersTo:="=Control!$B$11"
    wb.Names.Add Name:="LiveCurrent", RefersTo:="=Control!$B$12"
    wb.Names.Add Name:="LiveDMM", RefersTo:="=Control!$B$13"
    wb.Names.Add Name:="SetVoltage", RefersTo:="=Control!$B$16"
    wb.Names.Add Name:="SetCurrent", RefersTo:="=Control!$B$17"
    wb.Names.Add Name:="OCPEnabled", RefersTo:="=Control!$B$18"
    wb.Names.Add Name:="RampStartV", RefersTo:="=Control!$B$21"
    wb.Names.Add Name:="RampEndV", RefersTo:="=Control!$B$22"
    wb.Names.Add Name:="RampDuration", RefersTo:="=Control!$B$23"
    wb.Names.Add Name:="RampCycles", RefersTo:="=Control!$B$24"
    wb.Names.Add Name:="RampDelay", RefersTo:="=Control!$B$25"
    wb.Names.Add Name:="RampPingPong", RefersTo:="=Control!$B$26"
    wb.Names.Add Name:="RampCycle", RefersTo:="=Control!$D$21"
    wb.Names.Add Name:="RampVoltage", RefersTo:="=Control!$D$22"
    wb.Names.Add Name:="RampProgress", RefersTo:="=Control!$D$23"
    wb.Names.Add Name:="ExportStatus", RefersTo:="=Control!$B$30"
    
    On Error GoTo 0
End Sub

Sub AddButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Control")
    
    Dim btn As Button
    Dim leftPos As Double
    Dim topPos As Double
    
    leftPos = 280
    topPos = 40
    
    ' Connection buttons
    Set btn = ws.Buttons.Add(leftPos, topPos, 100, 22)
    btn.OnAction = "ConnectPSU"
    btn.Caption = "Connect PSU"
    
    Set btn = ws.Buttons.Add(leftPos + 105, topPos, 100, 22)
    btn.OnAction = "ConnectDMM"
    btn.Caption = "Connect DMM"
    
    Set btn = ws.Buttons.Add(leftPos + 210, topPos, 100, 22)
    btn.OnAction = "DisconnectAll"
    btn.Caption = "Disconnect All"
    
    topPos = topPos + 30
    
    Set btn = ws.Buttons.Add(leftPos, topPos, 100, 22)
    btn.OnAction = "InitSimulated"
    btn.Caption = "Test (Simulated)"
    
    ' PSU Control buttons
    topPos = 220
    
    Set btn = ws.Buttons.Add(leftPos, topPos, 80, 22)
    btn.OnAction = "OutputOn"
    btn.Caption = "Output ON"
    
    Set btn = ws.Buttons.Add(leftPos + 85, topPos, 80, 22)
    btn.OnAction = "OutputOff"
    btn.Caption = "Output OFF"
    
    Set btn = ws.Buttons.Add(leftPos + 170, topPos, 90, 22)
    btn.OnAction = "ApplySettings"
    btn.Caption = "Apply Settings"
    
    ' Logging buttons
    topPos = topPos + 30
    
    Set btn = ws.Buttons.Add(leftPos, topPos, 90, 22)
    btn.OnAction = "StartLogging"
    btn.Caption = "Start Logging"
    
    Set btn = ws.Buttons.Add(leftPos + 95, topPos, 90, 22)
    btn.OnAction = "StopLogging"
    btn.Caption = "Stop Logging"
    
    ' Ramp buttons
    topPos = topPos + 30
    
    Set btn = ws.Buttons.Add(leftPos, topPos, 80, 22)
    btn.OnAction = "StartRamp"
    btn.Caption = "Start Ramp"
    
    Set btn = ws.Buttons.Add(leftPos + 85, topPos, 80, 22)
    btn.OnAction = "StopRamp"
    btn.Caption = "Stop Ramp"
    
    Set btn = ws.Buttons.Add(leftPos + 170, topPos, 80, 22)
    btn.OnAction = "PauseRamp"
    btn.Caption = "Pause Ramp"
    
    ' Export buttons
    topPos = topPos + 30
    
    Set btn = ws.Buttons.Add(leftPos, topPos, 80, 22)
    btn.OnAction = "ExportCSV"
    btn.Caption = "Export CSV"
    
    Set btn = ws.Buttons.Add(leftPos + 85, topPos, 80, 22)
    btn.OnAction = "ClearData"
    btn.Caption = "Clear Data"
    
    MsgBox "Buttons added successfully!", vbInformation
End Sub

