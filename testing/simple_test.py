"""
Very simple test to be called from Excel
"""

def simple_test():
    """Simple function that returns a value to Excel"""
    import xlwings as xw
    
    wb = xw.Book.caller()
    sheet = wb.sheets["Control"]
    sheet.range("D3").value = "Python Works!"
    
    return "SUCCESS"

if __name__ == "__main__":
    print("This script should be called from Excel via RunPython")
    simple_test()
