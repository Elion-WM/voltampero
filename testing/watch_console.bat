@echo off
echo ================================================
echo VoltAmpero Console Watcher
echo ================================================
echo.
echo This will show Python debug output when you
echo click buttons in Excel.
echo.
echo Instructions:
echo 1. Leave this window open
echo 2. Open Excel with VoltAmpero.xlsm
echo 3. Click buttons in Excel
echo 4. Watch for debug messages here
echo.
echo Press Ctrl+C to stop
echo.
echo ================================================
echo.

REM Start Python in a way that shows all output
python -u -c "import time; print('Watching for Excel activity...'); print('Click buttons in Excel now!'); print(); [time.sleep(1) for _ in iter(int, 1)]"
