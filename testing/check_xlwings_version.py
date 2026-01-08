"""
Check xlwings version and see if there's a known issue
"""

import xlwings as xw
import sys

print("="*60)
print("xlwings Diagnostic Information")
print("="*60)

print(f"\nPython version: {sys.version}")
print(f"xlwings version: {xw.__version__}")
print(f"xlwings location: {xw.__file__}")

print("\nChecking for known issues with version 0.33.17...")

# Version 0.33.17 might have issues - let's check
if xw.__version__ == "0.33.17":
    print("\n[WARNING] You're using xlwings 0.33.17")
    print("There might be compatibility issues with this version.")
    print("\nRecommendation: Try upgrading xlwings:")
    print("   python -m pip install --upgrade xlwings")
else:
    print(f"\nxlwings version {xw.__version__} - checking compatibility...")

print("\n" + "="*60)
