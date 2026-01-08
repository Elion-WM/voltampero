# .gitignore Quick Reference for VoltAmpero

## What's Being Ignored

This reference explains what each section of `.gitignore` does.

### Python Files (Auto-Generated)
```gitignore
__pycache__/
*.py[cod]
*.so
```
**Why**: Python creates these cache files every time you run code. They're specific to your machine and Python version.

### Virtual Environments
```gitignore
python/
venv/
.venv/
```
**Why**: Your `python/` folder is 50MB+ of packages. Users should install their own with `pip install -r requirements.txt`.

### Debug/Log Files
```gitignore
*.log
debug_*.txt
timing_*.txt
capture_*.txt
```
**Why**: These are temporary debugging outputs. Not part of the project.

**EXCEPTION**: We keep `requirements.txt` and `LICENSE` (marked with `!`)

### Excel Temp Files
```gitignore
~$*.xlsm
~$*.xlsx
*.tmp
```
**Why**: Excel creates these when file is open. Causes conflicts if committed.

### IDE Files
```gitignore
.vscode/
.idea/
*.swp
```
**Why**: Your personal editor settings. Other users have their own preferences.

### OS Files
```gitignore
.DS_Store         # macOS
Thumbs.db         # Windows
```
**Why**: Operating system metadata. Meaningless to others.

### Project-Specific
```gitignore
*_backup_*.xlsm
voltampero_log_*.csv
get-pip.py
```
**Why**: Backup files, test outputs, and temporary installers.

### Security (CRITICAL!)
```gitignore
.env
secrets.txt
passwords.txt
api_keys.txt
```
**Why**: NEVER commit passwords or API keys! They become public on GitHub!

---

## Files That ARE Committed

These important files are NOT ignored:

✅ **Source Code**:
- `voltampero.py`
- `psu_korad.py`
- `multimeter_unit.py`

✅ **Application**:
- `VoltAmpero.xlsm`
- `VoltAmpero.bas`

✅ **Documentation**:
- `README.md`
- `docs/*.md`
- `LICENSE`

✅ **Configuration**:
- `requirements.txt`
- `xlwings.conf`

✅ **Repository Structure**:
- `REPOSITORY_STRUCTURE.md`
- `.gitignore`

---

## Testing the .gitignore

### See what Git will ignore:
```bash
git status --ignored
```

### See what Git will commit:
```bash
git status
```

### Check if a specific file is ignored:
```bash
git check-ignore -v filename.txt
```

---

## Before and After

### BEFORE .gitignore:
```bash
$ git status
Untracked files:
  __pycache__/voltampero.cpython-311.pyc
  __pycache__/psu_korad.cpython-311.pyc
  python/Lib/site-packages/...
  python/Scripts/...
  ... (500+ files)
  timing_debug.txt
  debug_log.txt
  ~$VoltAmpero.xlsm
```

### AFTER .gitignore:
```bash
$ git status
Untracked files:
  voltampero.py
  psu_korad.py
  VoltAmpero.xlsm
  docs/
  LICENSE
  README.md
```

**Much cleaner!** 🎉

---

## Updating .gitignore

### Add new pattern:
Edit `.gitignore` and add:
```gitignore
# New pattern
new_debug_*.log
```

### Apply to already-tracked files:
If you already committed files you want to ignore:
```bash
# Remove from Git tracking (but keep file locally)
git rm --cached filename.txt

# Or remove entire folder
git rm -r --cached __pycache__/

# Commit the change
git commit -m "Remove ignored files from tracking"
```

---

## Common Issues

### Issue: "File still showing in git status"
**Cause**: File was already committed before adding to .gitignore  
**Fix**: Remove from tracking:
```bash
git rm --cached filename.txt
git commit -m "Stop tracking filename.txt"
```

### Issue: "Can't find .gitignore"
**Cause**: File starts with dot (hidden on some systems)  
**Windows Fix**: Enable "Show hidden files" in File Explorer  
**Verify**: `dir /a` in command prompt shows hidden files

### Issue: ".gitignore not working"
**Cause**: Usually file permissions or encoding  
**Fix**: Ensure file is named exactly `.gitignore` (no extension!)

---

## Pro Tips

### 1. Ignore folder but keep structure:
```gitignore
logs/*      # Ignore all files in logs/
!logs/.gitkeep  # But keep the folder itself
```

### 2. Ignore everything except specific files:
```gitignore
folder/*         # Ignore everything in folder
!folder/important.txt  # Except this file
```

### 3. Pattern matching:
```gitignore
*.log           # All .log files anywhere
**/temp/        # "temp" folder at any level
file?.txt       # file1.txt, file2.txt, etc.
```

---

## Your Repository Size

### Without .gitignore:
- Total size: ~50-100 MB
- Files tracked: 500+
- Clone time: Minutes

### With .gitignore:
- Total size: ~200 KB
- Files tracked: ~20-30
- Clone time: Seconds

**Improvement**: 250x smaller, 20x fewer files! 🚀

---

## Verification Checklist

After creating .gitignore, verify:

- [ ] `git status` shows only important files
- [ ] `__pycache__/` is not listed
- [ ] `python/` folder is not listed
- [ ] `*.log` files are not listed
- [ ] `~$*.xlsm` Excel temp files are not listed
- [ ] `requirements.txt` IS listed (exception)
- [ ] `LICENSE` IS listed (exception)

---

## Resources

- Official Git docs: https://git-scm.com/docs/gitignore
- GitHub templates: https://github.com/github/gitignore
- Python template: https://github.com/github/gitignore/blob/main/Python.gitignore
- Interactive tester: https://www.toptal.com/developers/gitignore

---

*Your repository is now professionally organized! 🎯*
