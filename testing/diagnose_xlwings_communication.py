"""
Diagnose xlwings communication issue - Error 62
"""

import os
import tempfile
import sys

print("="*60)
print("Diagnosing xlwings Communication (Error 62)")
print("="*60)

# Check TEMP directory
print("\n1. Checking TEMP directory:")
temp_dir = tempfile.gettempdir()
print(f"   TEMP dir: {temp_dir}")
print(f"   Exists: {os.path.exists(temp_dir)}")
print(f"   Writable: {os.access(temp_dir, os.W_OK)}")

# Try creating a temp file
try:
    test_file = tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt')
    test_file.write("test")
    test_file.close()
    print(f"   [OK] Can create temp files: {test_file.name}")
    os.unlink(test_file.name)
except Exception as e:
    print(f"   [ERROR] Cannot create temp files: {e}")

# Check xlwings.bas for RunPython implementation
print("\n2. Checking xlwings.bas RunPython implementation:")
import xlwings as xw
xlwings_dir = os.path.dirname(xw.__file__)
xlwings_bas = os.path.join(xlwings_dir, "xlwings.bas")

if os.path.exists(xlwings_bas):
    print(f"   [OK] Found: {xlwings_bas}")
    
    # Read the file to see how RunPython works
    with open(xlwings_bas, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        
        # Look for RunPython implementation
        if 'Sub RunPython' in content:
            print("   [OK] RunPython sub found")
            
            # Check if it uses CreateObject for communication
            if 'CreateObject' in content:
                print("   Uses CreateObject for communication")
            
            # Check if it uses Shell or WScript
            if 'Shell' in content:
                print("   Uses Shell command")
            if 'WScript.Shell' in content:
                print("   Uses WScript.Shell")
                
            # Check for file I/O
            if 'Open ' in content and 'Input' in content:
                print("   [WARNING] Uses file I/O - this is where Error 62 occurs")
                print("   The VBA code is trying to read a response file from Python")
else:
    print(f"   [ERROR] xlwings.bas not found!")

# Check environment variables
print("\n3. Checking environment variables:")
env_vars = ['TEMP', 'TMP', 'USERPROFILE', 'APPDATA']
for var in env_vars:
    value = os.environ.get(var, 'NOT SET')
    print(f"   {var}: {value}")

# Check if xlwings uses UDF server
print("\n4. Checking xlwings configuration:")
print(f"   xlwings version: {xw.__version__}")
print(f"   xlwings location: {xw.__file__}")

# Look for xlwings config
config_file = os.path.join(os.path.dirname(xw.__file__), 'xlwings.conf')
if os.path.exists(config_file):
    print(f"   [INFO] Global xlwings.conf exists at: {config_file}")

print("\n5. Testing Python execution directly:")
# Try to simulate what xlwings does
test_script = '''
import sys
print("Python executed successfully")
sys.exit(0)
'''

temp_script = tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.py')
temp_script.write(test_script)
temp_script.close()

print(f"   Created test script: {temp_script.name}")

# Try to run it
import subprocess
try:
    result = subprocess.run(
        [r"C:\Users\User\GitHub\voltampero\python\python.exe", temp_script.name],
        capture_output=True,
        text=True,
        timeout=5
    )
    print(f"   [OK] Python execution result: {result.stdout.strip()}")
    print(f"   Return code: {result.returncode}")
except Exception as e:
    print(f"   [ERROR] Python execution failed: {e}")

os.unlink(temp_script.name)

print("\n" + "="*60)
print("LIKELY ISSUE:")
print("xlwings VBA code creates a temporary file to get Python's")
print("response, but either:")
print("1. Python isn't writing the response file")
print("2. VBA can't read the response file")
print("3. There's a timing issue")
print("\nLet me create a workaround...")
print("="*60)
