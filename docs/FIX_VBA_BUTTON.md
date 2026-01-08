# VBA Button Not Calling Python Function

## Problem
The "Output ON" button in Excel is NOT calling the Python function at all.
- No debug_log.txt file is created
- This means VBA macro isn't executing

## Possible Causes

1. **Button not assigned to macro** - Most likely
2. **Macro name mismatch**
3. **xlwings not properly configured**
4. **Macro security blocking execution**

## How to Fix

### Option 1: Check Button Assignment (Most Likely Issue)

1. Open VoltAmpero.xlsm
2. **Right-click** on the "Output ON" button
3. Select **"Assign Macro..."**
4. You should see "OutputOn" in the list
5. If NOT, or if different macro is selected:
   - Select **"OutputOn"** from the list
   - Click **OK**
6. Repeat for "Output OFF" button → assign to "OutputOff"

### Option 2: Manually Test the Macro

1. Open VoltAmpero.xlsm
2. Press `Alt + F8` (Macro dialog)
3. Find "OutputOn" in the list
4. Click **Run**
5. Check if `debug_log.txt` appears in the folder

If log file appears → Button assignment is the problem
If log file doesn't appear → xlwings configuration issue

### Option 3: Check xlwings Configuration

1. In Excel VBA Editor (`Alt + F11`)
2. Check if xlwings VBA module exists
3. Go to Tools → References
4. Check if any references are marked as MISSING

### Option 4: Recreate the Button

If button assignment is lost:

1. Open VoltAmpero.xlsm
2. Delete the old "Output ON" button
3. Insert → Shapes → Button
4. Draw the button
5. Assign Macro dialog appears → Select "OutputOn"
6. Type "Output ON" as button text

## Quick Test

Run this VBScript to test if VBA can call Python:
```
cscript test_vba_directly.vbs
```

This will:
1. Open Excel
2. Call the OutputOn macro
3. Check if debug_log.txt was created
4. Show results

## Expected Behavior

**When working correctly:**
1. Click "Output ON" button
2. `debug_log.txt` file created instantly
3. PSU output turns on

**Current behavior:**
1. Click "Output ON" button
2. Nothing happens
3. No log file created
4. = VBA macro not running at all

## Action Items

1. Check button macro assignment
2. Test macro directly with Alt+F8
3. If still fails, run test_vba_directly.vbs
4. Report results

---

The Python code is correct. The VBA button just isn't calling it.
