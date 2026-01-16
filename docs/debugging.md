# Debugging Guide — Zotero PowerPoint Picker

This document collects common runtime issues, error messages, and their root causes,
with reproducible fixes.

It is intended for developers and advanced users.

---

## General debugging checklist

Before investigating specific errors, verify:

### 1) You are using the intended Python interpreter

```powershell
python -c "import sys; print(sys.executable)"
```

If this points to a system Python (instead of your virtual environment), either:
- activate the correct virtual environment, or
- install dependencies for that system interpreter.

### 2) Dependencies are installed

Recommended (virtual environment):

```powershell
# from the project directory
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -r requirements.txt
```

---

## win32com / pywin32 not found (Windows)

### Symptoms

- `NameError: name 'win32' is not defined`
- `ModuleNotFoundError: No module named 'win32com'`

### Root cause

PowerPoint automation relies on **pywin32**.  
`win32com` is not part of Python itself and must be installed for the Python interpreter
that runs the script.

This error usually means the script is executed with a Python environment where
dependencies are missing.

### Fix (recommended: virtual environment)

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -r requirements.txt
```

### Fix (system Python, no venv)

```powershell
python -m pip install -U pip
python -m pip install pywin32
python -m pip install -r requirements.txt
```

### Verification

```powershell
python -c "import win32com.client as win32; print('win32com OK')"
python -c "import pyzotero; print('pyzotero OK')"
```

---

## PowerPoint COM errors ("CoInitialize was not called")

### Symptoms

- COM errors mentioning `CoInitialize`
- Errors during background actions (cleanup, bibliography update)

### Root cause

COM objects are accessed from a thread that has not been initialized for COM usage.

### Fix

Every thread that accesses PowerPoint COM must call:

```python
import pythoncom

pythoncom.CoInitialize()
try:
    # COM work here
    ...
finally:
    pythoncom.CoUninitialize()
```

Important:
- This must happen **inside the worker thread** that uses COM, not just in the main thread.

---

## No citation inserted / text appears in the wrong place

### Symptoms

- Citation text is appended at the end of a text box
- Citation is inserted although no text cursor is visible

### Root cause

PowerPoint distinguishes between:
- selecting a text frame, and
- placing a real text cursor inside the text.

Insertion is only safe when a real text cursor exists.

### Expected behavior

If no text cursor is present, insertion must fail with a clear error.
Avoid silent fallbacks to shape-level text ranges.

---

## "Resolve anchors … gefunden=0" (no anchors found)

### Symptoms

Log output similar to:

```text
Resolve anchors … gefunden=0
```

### Root cause

No PowerPoint presentation is currently open, or no bibliography anchor has been set yet.

### Status

This is not an error. It is expected behavior when:
- PowerPoint is open without a presentation, or
- the bibliography anchor has not been created yet.

---

## Configuration dialog does not appear

### Symptoms

- Application starts, but no credential dialog is shown
- No error message appears

### Root cause

Valid Zotero credentials were already found via:
- local user config file, or
- environment variables.

### Fix

Delete the local configuration file to force the dialog:

**Windows**
```text
%APPDATA%\ZoteroPowerPoint\config.json
```

Then restart the application.

---

## Environment variables override local config

### Symptoms

- Changing credentials in the dialog has no effect
- Old credentials keep being used

### Root cause

Environment variables override values from the local config file.

### Check (PowerShell)

```powershell
echo $Env:ZOTERO_API_KEY
echo $Env:ZOTERO_LIBRARY_ID
echo $Env:ZOTERO_LIBRARY_TYPE
```

Unset them if you want to rely on the local configuration dialog.

---

## Logging and diagnostics

### Recommendations

- Always log which source provided credentials:
  - file
  - env
  - prompt
- Never log API keys or secrets.
- Mask sensitive values if needed.

### When a bug is likely not in the code

Check these before changing code:

- No PowerPoint presentation is open
- Wrong Python interpreter is used
- Missing dependencies
- Environment variables overriding expected behavior