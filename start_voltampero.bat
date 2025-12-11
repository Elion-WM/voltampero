@echo off
setlocal ENABLEDELAYEDEXPANSION

REM VoltAmpero Launcher - single entrypoint
REM Usage: start_voltampero.bat [launch|monitor|reset|sim|help]

set REPO_DIR=%~dp0
set REPO_DIR=%REPO_DIR:~0,-1%
set WORKBOOK=VoltAmpero.xlsm
set XLWINGS_CONF=%REPO_DIR%\xlwings.conf
set PYTHON_EXE=%REPO_DIR%\python\python.exe

if "%1"=="" goto :menu
if /I "%1"=="launch" goto :launch
if /I "%1"=="monitor" goto :monitor
if /I "%1"=="reset" goto :reset
if /I "%1"=="sim" goto :sim
if /I "%1"=="help" goto :help

:help
  echo Usage: %~nx0 ^<launch^|monitor^|reset^|sim^|help^>
  exit /b 0

:menu
  echo ==============================
  echo VoltAmpero Launcher
  echo Repo: %REPO_DIR%
  echo Workbook: %WORKBOOK%
  echo ------------------------------
  echo  1^) Launch Excel (with setup)
  echo  2^) Monitor status
  echo  3^) Reset (stop log/ramp, disconnect)
  echo  4^) Initialize Simulated mode
  echo  5^) Help
  echo  0^) Exit
  echo ==============================
  set /p CHOICE=Select option: 
  if "%CHOICE%"=="1" goto :launch
  if "%CHOICE%"=="2" goto :monitor
  if "%CHOICE%"=="3" goto :reset
  if "%CHOICE%"=="4" goto :sim
  if "%CHOICE%"=="5" goto :help
  if "%CHOICE%"=="0" exit /b 0
  goto :menu

:ensure_conf
  if exist "%XLWINGS_CONF%" goto :eof
  echo Creating xlwings.conf at %XLWINGS_CONF%
  > "%XLWINGS_CONF%" echo [xlwings]
  >> "%XLWINGS_CONF%" echo PYTHONPATH=%REPO_DIR%
  >> "%XLWINGS_CONF%" echo INTERPRETER=%PYTHON_EXE%
  goto :eof

:launch
  call :ensure_conf
  if not exist "%PYTHON_EXE%" (
    echo ERROR: Embedded Python not found at %PYTHON_EXE%
    echo Please run setup per README or adjust xlwings.conf INTERPRETER.
    pause
    exit /b 1
  )
  if exist "%REPO_DIR%\%WORKBOOK%" (
    echo Opening %WORKBOOK%...
    start "" "%REPO_DIR%\%WORKBOOK%"
  ) else (
    echo Workbook %WORKBOOK% not found in repo root.
    echo Opening EXCEL_SETUP.md for instructions...
    start "" "%REPO_DIR%\EXCEL_SETUP.md"
    start "" excel.exe
  )
  exit /b 0

:monitor
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; ^
    try{ $excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $excel=$null } ^
    if(-not $excel){ Write-Host 'Excel not running.'; exit 1 } ^
    $wb=$null; foreach($b in $excel.Workbooks){ if($b.Name -like 'VoltAmpero*.xlsm'){ $wb=$b; break } } ^
    if(-not $wb){ Write-Host 'VoltAmpero workbook not found.'; exit 2 } ^
    function GetName($n){ try { return $wb.Names.Item($n).RefersToRange.Value } catch { return '' } } ^
    Write-Host 'Monitoring... Press Ctrl+C to stop.'; ^
    while($true){ ^
      $psu=GetName('PSUStatus'); $dmm=GetName('DMMStatus'); $log=GetName('LoggingStatus'); $ramp=GetName('RampStatus'); ^
      $v=GetName('LiveVoltage'); $a=GetName('LiveCurrent'); $d=GetName('LiveDMM'); ^
      $cycle=GetName('RampCycle'); $rv=GetName('RampVoltage'); $prog=GetName('RampProgress'); ^
      $ts=(Get-Date).ToString('HH:mm:ss'); ^
      Write-Host ("[$ts] PSU={0} DMM={1} LOG={2} RAMP={3}  V={4}V I={5}A  DMM={6}  Cycle={7} Vramp={8} Prog={9}" -f $psu,$dmm,$log,$ramp,$v,$a,$d,$cycle,$rv,$prog); ^
      Start-Sleep -Seconds 3 ^
    }"
  exit /b %ERRORLEVEL%

:reset
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; ^
    try{ $excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $excel=$null } ^
    if(-not $excel){ Write-Host 'Excel not running.'; exit 1 } ^
    $wb=$null; foreach($b in $excel.Workbooks){ if($b.Name -like 'VoltAmpero*.xlsm'){ $wb=$b; break } } ^
    if(-not $wb){ Write-Host 'VoltAmpero workbook not found.'; exit 2 } ^
    Write-Host ('Resetting workbook {0}...' -f $wb.Name); ^
    try { $excel.Run("'"+$wb.Name+"'!StopLogging") } catch {} ^
    try { $excel.Run("'"+$wb.Name+"'!StopRamp") } catch {} ^
    try { $excel.Run("'"+$wb.Name+"'!DisconnectAll") } catch {} ^
    Write-Host 'Done.'"
  exit /b %ERRORLEVEL%

:sim
  call :ensure_conf
  if not exist "%REPO_DIR%\%WORKBOOK%" (
    echo Workbook %WORKBOOK% not found. Aborting.
    exit /b 2
  )
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; ^
    $excel=$null; try{ $excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $excel=$null } ^
    if(-not $excel){ $excel=New-Object -ComObject Excel.Application; $excel.Visible=$true } ^
    $wb=$null; foreach($b in $excel.Workbooks){ if($b.FullName -ieq '%REPO_DIR%\%WORKBOOK%'){ $wb=$b; break } } ^
    if(-not $wb){ $wb=$excel.Workbooks.Open('%REPO_DIR%\%WORKBOOK%') } ^
    try { $excel.Run("'"+$wb.Name+"'!InitSimulated"); } catch { Write-Host 'InitSimulated macro not found.' } ^
    Write-Host 'Simulated mode initialized.'"
  exit /b %ERRORLEVEL%

endlocal
