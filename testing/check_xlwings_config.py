"""
Check xlwings configuration in detail
"""

import xlwings as xw
import os
import sys

print("="*60)
print("xlwings Configuration Check")
print("="*60)

# Check xlwings.conf
print("\n1. xlwings.conf file:")
conf_path = r"C:\Users\User\GitHub\voltampero\xlwings.conf"
if os.path.exists(conf_path):
    print(f"   [OK] Found at: {conf_path}")
    with open(conf_path, 'r') as f:
        content = f.read()
        print("   Content:")
        for line in content.split('\n'):
            print(f"      {line}")
else:
    print("   [ERROR] Not found!")

# Check Python interpreter
print("\n2. Python interpreter:")
interpreter = r"C:\Users\User\GitHub\voltampero\python\python.exe"
if os.path.exists(interpreter):
    print(f"   [OK] Found at: {interpreter}")
else:
    print(f"   [ERROR] Not found at: {interpreter}")

# Check PYTHONPATH
print("\n3. Python path:")
pythonpath = r"C:\Users\User\GitHub\voltampero"
if os.path.exists(pythonpath):
    print(f"   [OK] Path exists: {pythonpath}")
    print(f"   Contents:")
    for item in os.listdir(pythonpath):
        if item.endswith('.py'):
            print(f"      - {item}")
else:
    print(f"   [ERROR] Path not found!")

# Check voltampero.py
print("\n4. voltampero.py module:")
voltampero_py = os.path.join(pythonpath, "voltampero.py")
if os.path.exists(voltampero_py):
    print(f"   [OK] Found")
    # Try importing it
    try:
        sys.path.insert(0, pythonpath)
        import voltampero
        print("   [OK] Can import voltampero module")
        print(f"   [OK] Has get_controller: {hasattr(voltampero, 'get_controller')}")
        print(f"   [OK] Has va_init_simulated: {hasattr(voltampero, 'va_init_simulated')}")
    except Exception as e:
        print(f"   [ERROR] Cannot import: {e}")
else:
    print(f"   [ERROR] Not found!")

# Check xlwings version
print(f"\n5. xlwings version: {xw.__version__}")

# Try to create a minimal test
print("\n6. Testing xlwings Book.caller() simulation:")
try:
    # This won't work outside Excel but let's see the error
    print("   (This test only works when called from Excel)")
except Exception as e:
    print(f"   Expected error: {e}")

print("\n" + "="*60)
