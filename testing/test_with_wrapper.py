"""
Test with the wrapper batch file
"""

import subprocess
import os
import tempfile

# Test with wrapper
INTERPRETER = r"C:\Users\User\GitHub\voltampero\python_wrapper.bat"
PYTHONPATH = r"C:\Users\User\GitHub\voltampero"
PythonCommand = "import sys; print('Hello from Python via wrapper')"

log_file = os.path.join(tempfile.gettempdir(), "test_xlwings2.log")

cmd = (
    f'cmd.exe /C "{INTERPRETER}" -B -c '
    f'"import xlwings.utils;xlwings.utils.prepare_sys_path(\\\"{PYTHONPATH}\\\"); '
    f'{PythonCommand}" '
    f'2> "{log_file}"'
)

print("Testing with wrapper batch file...")
print(f"\nCommand: {cmd}\n")

try:
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=10)
    print(f"Exit code: {result.returncode}")
    
    if result.returncode == 0:
        print("[SUCCESS] Python executed successfully!")
    else:
        print(f"[ERROR] Exit code: {result.returncode}")
    
    if os.path.exists(log_file):
        size = os.path.getsize(log_file)
        if size > 0:
            with open(log_file, 'r') as f:
                print(f"Log content:\n{f.read()}")
        else:
            print("Log file is empty (good if exit code is 0)")
        os.remove(log_file)
        
except Exception as e:
    print(f"ERROR: {e}")
