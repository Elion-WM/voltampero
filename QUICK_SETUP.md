# VoltAmpero Quick Excel Setup (3 Steps)

## Step 1: Create Excel File
1. Open Excel
2. **Save As** → `VoltAmpero.xlsm` (Macro-Enabled Workbook)
3. Save in the same folder as voltampero.py

## Step 2: Import VBA Module
1. Press **Alt+F11** (opens VBA Editor)
2. **File → Import File...**
3. Select `VoltAmpero.bas` from this folder
4. Close VBA Editor (Alt+Q)

## Step 3: Run Setup Macros
1. Press **Alt+F8** (opens Macro dialog)
2. Run **SetupWorkbook** → Creates sheets, cells, and named ranges
3. Run **AddButtons** → Adds all control buttons

**Done!** Your workbook is ready. Click "Test (Simulated)" to verify it works.

---

## First Test
1. Click **Test (Simulated)** button
2. Click **Start Logging**
3. Watch data appear in the Data sheet
4. Click **Stop Logging**
5. Click **Export CSV** to save

## Troubleshooting
- **Macros disabled?** → File → Options → Trust Center → Enable macros
- **Python not found?** → Edit xlwings.conf with your Python path
- **Module error?** → Run: `pip install --user pyserial hidapi xlwings`
