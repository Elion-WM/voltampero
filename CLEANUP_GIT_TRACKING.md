# Clean Up Git Tracking

## Problem

The `.gitignore` file works ONLY for **new files**. Files that were already committed to git BEFORE creating `.gitignore` are still being tracked.

**Right now, git is tracking:**
- `__pycache__/` files (should be ignored)
- `python/` virtual environment (should be ignored)
- Old test files (should be ignored)

## Solution

We need to **remove these files from git tracking** (but keep them on your computer).

---

## Step-by-Step Cleanup

### Step 1: Remove Virtual Environment from Tracking

```bash
git rm -r --cached python/
```

**What this does:**
- Removes `python/` from git tracking
- Keeps the folder on your computer (you still need it!)
- Next commit will delete it from the repository

### Step 2: Remove Python Cache

```bash
git rm -r --cached __pycache__/
```

**What this does:**
- Removes all `__pycache__/` folders from tracking
- Cache files will regenerate when you run Python

### Step 3: Remove Old Files That Were Moved

```bash
git add -u
```

**What this does:**
- Stages all the deletions (files moved to /testing/ and /docs/)

### Step 4: Add New Files

```bash
git add .gitignore
git add LICENSE
git add REPOSITORY_STRUCTURE.md
git add docs/
git add testing/
```

**What this does:**
- Adds the new files we created

### Step 5: Commit the Cleanup

```bash
git commit -m "Repository cleanup: Add LICENSE, .gitignore, organize structure

- Add LICENSE (Elion-Circular, non-commercial)
- Add comprehensive .gitignore
- Remove python/ virtual env from tracking
- Remove __pycache__/ from tracking
- Move test scripts to /testing/ (89 files)
- Move documentation to /docs/ (15 files)
- Add protocol documentation (Korad KWR102, UNI-T UT8804E)
- Update VoltAmpero code with fixes

Co-authored-by: factory-droid[bot] <138933559+factory-droid[bot]@users.noreply.github.com>"
```

---

## Quick One-Line Cleanup

If you want to do it all at once:

```bash
git rm -r --cached python/ __pycache__/ && git add -A && git commit -m "Cleanup: Remove ignored files from tracking"
```

---

## Expected Result

### BEFORE cleanup:
```
Changes not staged for commit:
  modified: __pycache__/psu_korad.cpython-311.pyc
  modified: python/Lib/site-packages/...
  ... (50+ files)
```

### AFTER cleanup:
```
On branch main
nothing to commit, working tree clean
```

---

## Verification

After cleanup, check:

```bash
# Should show only source files
git ls-files

# Should NOT include:
# - python/
# - __pycache__/
# - *.log files
```

---

## Why This Matters

**Current repository size**: ~50MB (with python/ and cache)  
**After cleanup**: ~200KB (clean code only)  

**Benefit**: 250x smaller, faster clones, professional appearance

---

**Ready to run the cleanup?** This is a one-time operation to clean up the repository.
