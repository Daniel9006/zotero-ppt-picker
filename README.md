# zotero-ppt-picker

A Python-based GUI tool to insert Zotero citations and generate or update a bibliography
in Microsoft PowerPoint.

---

## Documentation

This repository contains additional documentation for developers and
platform-specific topics:

- Developer notes and internal details: `docs/development.md`
- Debugging and known runtime issues: `docs/debugging.md`
- macOS / Linux notes and limitations: `docs/mac_linux.md`

The main `README.md` focuses on installation, configuration, and usage.

---

## Requirements

- Windows 11
- Microsoft PowerPoint (desktop version)
- Python 3.13+
- Git

---

## Setup (Windows)

```powershell
# clone repository
git clone https://github.com/Daniel9006/zotero-ppt-picker.git
cd zotero-ppt-picker

# create virtual environment
py -m venv .venv
.\.venv\Scripts\Activate.ps1

# install dependencies
python -m pip install -U pip
pip install -r requirements.txt
```

---

## Zotero credentials configuration

This tool accesses Zotero via the Zotero Web API.

Credentials are stored **per user**, locally, and are **never committed to the repository**.

### Credential storage location

- **Windows**
  ```
  %APPDATA%\ZoteroPowerPoint\config.json
  ```
- **macOS / Linux**
  ```
  ~/.config/ZoteroPowerPoint/config.json
  ```

The configuration file contains:
- `api_key`
- `library_id`
- `library_type` (`user` or `group`)

---

### First run behavior

On first launch (or if no configuration file exists), the application opens a dialog
asking for your Zotero credentials.

You can choose to:
- save the credentials locally, or
- use them only for the current session

---

### Change or reset credentials

To force the configuration dialog to appear again:

1. Close the application
2. Delete the local configuration file
3. Start the application again

---

### Environment variables (optional)

Credentials can also be provided via environment variables:

```
ZOTERO_API_KEY
ZOTERO_LIBRARY_ID
ZOTERO_LIBRARY_TYPE
```

Environment variables override the local configuration file.

This method is intended for advanced users only.

---

## Usage

1. Open Microsoft PowerPoint
2. Open a presentation
3. Place the text cursor where the citation should be inserted
4. Run the script:
   ```powershell
   python zotero_picker_ppt.py
   ```
5. Select a reference and insert it

---

## Troubleshooting

### win32com / pywin32 not found (Windows)

If you see errors like:

- `NameError: name 'win32' is not defined`
- `ModuleNotFoundError: No module named 'win32com'`

the script is being executed with a Python interpreter where required
dependencies are not installed.

`win32com` is provided by **pywin32** and is **not part of standard Python**.

#### Fix

Ensure dependencies are installed for the interpreter used to start the script:

```powershell
python -m pip install -U pip
python -m pip install -r requirements.txt
```

#### Verify installation

```powershell
python -c "import win32com.client as win32; print('win32com OK')"
python -c "import pyzotero; print('pyzotero OK')"
```

#### Check Python interpreter

```powershell
python -c "import sys; print(sys.executable)"
```

If this points to a system Python instead of the virtual environment,
activate the correct environment or install dependencies for that interpreter.

---

### Configuration dialog does not appear

This usually means valid credentials were already found.

Delete the local configuration file to force the dialog to appear again.

---

## Notes for developers

- Configuration handling is implemented in `config/zotero_config.py`
- Versioning rules are defined in `VERSIONING.md`
- Coding rules are defined in `CODING_STANDARDS.md`

For architectural details and refactoring plans, see `docs/development.md`.