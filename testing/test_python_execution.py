"""
Test if Python can be executed the way xlwings does it
"""

import subprocess
import os
import tempfile

# Simulate exactly what xlwings does
INTERPRETER = r"C:\Users\User\GitHub\voltampero\python\python.exe"
PYTHONPATH = r"C:\Users\User\GitHub\voltampero"
PythonCommand = "import sys; print('Hello from Python')"

# Create a temp log file
log_file = os.path.join(tempfile.gettempdir(), "test_xlwings.log")

# Build the command exactly like xlwings does
cmd = (
    f'cmd.exe /C "{INTERPRETER}" -B -c '
    f'"import xlwings.utils;xlwings.utils.prepare_sys_path(\\\"{PYTHONPATH}\\\"); '
    f'{PythonCommand}" '
    f'2> "{log_file}"'
)

print("Testing xlwings-style Python execution...")
print(f"\nCommand: {cmd}\n")

# Execute
try:
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=10)
    print(f"Exit code: {result.returncode}")
    print(f"STDOUT: {result.stdout}")
    print(f"STDERR: {result.stderr}")
    
    # Check log file
    if os.path.exists(log_file):
        size = os.path.getsize(log_file)
        print(f"\nLog file: {log_file}")
        print(f"Log file size: {size} bytes")
        
        if size > 0:
            with open(log_file, 'r') as f:
                content = f.read()
                print(f"Log content:\n{content}")
        else:
            print("Log file is EMPTY (this causes Error 62!)")
        
        os.remove(log_file)
    else:
        print(f"\nLog file was not created!")
        
except Exception as e:
    print(f"ERROR: {e}")

print("\n" + "="*60)
print("If the exit code is non-zero but log file is empty,")
print("that's the source of Error 62!")
print("="*60)
