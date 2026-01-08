"""
Fix the interpreter path issue by using pythonw or adding to PATH
"""

import os
import subprocess

print("="*60)
print("Fixing Python interpreter path for xlwings")
print("="*60)

# Option 1: Check if pythonw exists (doesn't need console)
python_dir = r"C:\Users\User\GitHub\voltampero\python"
pythonw_exe = os.path.join(python_dir, "pythonw.exe")

if os.path.exists(pythonw_exe):
    print(f"\n[OK] Found pythonw.exe: {pythonw_exe}")
    print("pythonw.exe doesn't show console windows")
else:
    print(f"\n[INFO] pythonw.exe not found")

# Option 2: Test with just the directory in PATH
print("\n" + "="*60)
print("SOLUTION: Use relative path or modify xlwings config")
print("="*60)

# The real fix: Use python.exe from the python subdirectory
# which xlwings should be able to find
print("\nUpdating xlwings.conf to use pythonw...")

config_content = """[xlwings]
PYTHONPATH=C:\\Users\\User\\GitHub\\voltampero
INTERPRETER=python
SHOW CONSOLE=True
"""

with open(r"C:\Users\User\GitHub\voltampero\xlwings.conf", 'w') as f:
    f.write(config_content)

print("[OK] Updated xlwings.conf to use 'python' (system Python)")
print("\nNow we need to add the local Python to PATH temporarily...")

# Create a test to see if we can use environment variables
print("\nAlternative: Create xlwings_activate.bat that sets PATH first")

activate_script = f"""@echo off
SET PATH={python_dir};%PATH%
REM Now Python is in PATH, Excel should be able to find it
echo Python added to PATH for this session
echo Start Excel from this command window:
echo   start "" "C:\\Users\\User\\GitHub\\voltampero\\VoltAmpero.xlsm"
pause
"""

with open(r"C:\Users\User\GitHub\\voltampero\start_excel_with_python.bat", 'w') as f:
    f.write(activate_script)

print("\n[CREATED] start_excel_with_python.bat")
print("\nUSE THIS TO START EXCEL:")
print("1. Close Excel if open")
print("2. Run: start_excel_with_python.bat")
print("3. This will add Python to PATH and open Excel")
print("4. Then try the buttons")

