@echo off
echo Adding Python to PATH...
SET PATH=C:\Users\User\GitHub\voltampero\python;%PATH%
echo Python added to PATH for this session
echo.
echo Starting Excel with VoltAmpero.xlsm...
start "" "C:\Users\User\GitHub\voltampero\VoltAmpero.xlsm"
echo.
echo Excel is now running with Python in PATH.
echo You can now use the buttons in Excel.
echo.
echo Keep this window open while using Excel!
pause
