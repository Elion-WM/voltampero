@echo off
echo ================================================
echo Unlocking COM Port (Killing Excel and Python)
echo ================================================
echo.
echo WARNING: This will close Excel and Python
echo Make sure you saved your work in Excel!
echo.
pause
echo.
echo Killing Python processes...
taskkill /F /IM python.exe >nul 2>&1
if %errorlevel%==0 (
    echo [OK] Python processes terminated
) else (
    echo [INFO] No Python processes found
)

echo.
echo Killing Excel processes...
taskkill /F /IM EXCEL.EXE >nul 2>&1
if %errorlevel%==0 (
    echo [OK] Excel processes terminated
) else (
    echo [INFO] No Excel processes found
)

echo.
echo Waiting 3 seconds for processes to clean up...
timeout /t 3 /nobreak >nul

echo.
echo [OK] COM port should be unlocked now
echo You can now run the test or open Excel
echo.
pause
