# zotero-ppt-picker

A Python-based GUI tool to insert Zotero citations and generate or update a bibliography
in Microsoft PowerPoint.

---

## Documentation

This repository contains additional documentation for developers and
platform-specific topics:

- Developer notes and internal details: `docs/development.md`
- macOS / Linux notes and limitations: `docs/mac_linux.md`

The main `README.md` focuses on installation, configuration, and usage.

---

## Requirements

- Windows 11
- Microsoft PowerPoint (desktop version)
- Python 3.13+ (recommended: `py` launcher)
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

Credentials are saved in a user-specific configuration file:

- **Windows**
  ```
  %APPDATA%\ZoteroPowerPoint\config.json
  ```
- **macOS / Linux**
  ```
  ~/.config/ZoteroPowerPoint/config.json
  ```

The file contains:
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
2. Delete the local configuration file:
   - Windows:
     ```
     %APPDATA%\ZoteroPowerPoint\config.json
     ```
3. Start the application again

    - The credential dialog will be shown again.

---

### Alternative: Environment variables (optional)

Instead of the local configuration file, credentials can also be provided via
environment variables:

```
ZOTERO_API_KEY
ZOTERO_LIBRARY_ID
ZOTERO_LIBRARY_TYPE   (user | group)
```

Environment variables **override** the local configuration file.

⚠️ This approach is **not recommended** for collaborative setups, as it is less
transparent and harder to manage across multiple users.

---

### Security notes

- No credentials are stored in the Git repository
- No secrets are hardcoded
- Each user uses their own local configuration
- Cross-platform compatible (Windows, macOS, Linux)

---

## Troubleshooting

### No configuration dialog appears
This usually means valid credentials were already found.

- Check whether the local config file exists:
  ```
  %APPDATA%\ZoteroPowerPoint\config.json
  ```
- Delete the file to force the dialog to appear again.

---

### Invalid API key
- Ensure the API key was generated on:
  ```
  https://www.zotero.org/settings/security/Applications
  ```
- Make sure the key has sufficient permissions (read access at minimum).

---

### Wrong Library ID
- **User library**  
  The Library ID is your **User ID**:
  - Zotero Web → Settings → Security → Applications

- **Group library**  
  The Library ID is the **Group ID**:
  - Zotero Web → Groups → select the group  
  - The Group ID is usually visible in the URL

---

### Dialog window does not appear or appears behind other windows
- This can happen on some Windows setups with multiple monitors or DPI scaling.
- The application forces the dialog to the foreground, but if issues persist:
  - Close all PowerPoint windows
  - Restart the script
  - Avoid running it minimized on first launch

---

### Credentials seem to be ignored
- Environment variables override the local config file.
- Check whether any of the following are set in your system:
  ```
  ZOTERO_API_KEY
  ZOTERO_LIBRARY_ID
  ZOTERO_LIBRARY_TYPE
  ```
- Remove them if you want to rely on the local configuration dialog.

---

## Notes for developers

- Configuration handling is implemented in `config/zotero_config.py`
- Credential loading priority:
  1. Local user config file
  2. Environment variables
  3. Interactive GUI prompt
- The design supports multi-user and collaborative workflows by default

For more detailed developer documentation (architecture, debugging,
and future refactors), see:

- `docs/development.md`