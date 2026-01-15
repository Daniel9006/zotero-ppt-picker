# zotero-ppt-picker – Developer Notes

This document describes the internal architecture and design decisions for developers.

---

## Project structure

```
zotero-ppt-picker/
├── zotero_picker_ppt.py      # main application entry point
├── config/
│   ├── __init__.py
│   └── zotero_config.py      # configuration & credential handling
├── test_zotero_config.py     # Standalone test for config & credential dialog
├── requirements.txt          # Python dependencies
├── README.md                 # User documentation (Windows, usage, credentials)
├── README_mac_linux.md       # macOS / Linux setup notes
├── README_dev.md             # Developer & architecture documentation
├── TEAM.md                   # Team / collaboration notes
├── .env                      # Optional local overrides (never commit secrets)
└── .gitignore
```

### Notes
- `test_zotero_config.py` is intended for **manual testing and debugging** of the
  configuration and credential dialog without running the full application.
- `.env` is optional and should only be used locally.
- Any file containing real API keys must **never** be committed.

---

## Configuration architecture

Credential loading priority:

1. Local user config file  
   - Platform-specific location (via `platformdirs`)
   - Stored as JSON
2. Environment variables  
   - `ZOTERO_API_KEY`
   - `ZOTERO_LIBRARY_ID`
   - `ZOTERO_LIBRARY_TYPE`
3. Interactive GUI prompt (Tkinter)

This ensures:
- no secrets in the repository
- per-user configuration
- deterministic behavior

⚠️ Files containing real Zotero API keys must never be committed.  
Use the interactive configuration dialog or local `.env` files instead.

---

## zotero_config.py

Key responsibilities:

- Loading credentials from multiple sources
- Validating user input
- Persisting configuration safely
- Providing a single entry point:
  ```python
  load_zotero_config(allow_prompt=True, parent=...)
  ```

Errors are raised as `ConfigError` and must be handled by the caller.

---

## GUI design notes

- Tkinter dialogs are implemented defensively for Windows focus handling
- Modal dialogs use `grab_set()` only after becoming visible
- A hidden but real root window is used to avoid Windows/Tk issues

---

## Testing

Recommended local tests:

```bash
# force credential dialog
rm ~/.config/ZoteroPowerPoint/config.json

python test_zotero_config.py
```

---

## Contribution guidelines

- Do not commit credentials or `.env` files
- Keep configuration logic inside `zotero_config.py`
- Avoid platform-specific logic outside dedicated modules
