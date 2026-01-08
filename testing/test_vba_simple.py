"""
Simple test to verify VBA can call Python
Add this to test if VBA RunPython is working at all
"""

def test_vba_call():
    """Simple function to test VBA -> Python communication"""
    print("SUCCESS: VBA can call Python!")
    
    # Try to write to a file to prove it ran
    with open("vba_test_success.txt", "w") as f:
        f.write("VBA successfully called Python at: ")
        from datetime import datetime
        f.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    return "VBA_OK"

if __name__ == "__main__":
    result = test_vba_call()
    print(result)
