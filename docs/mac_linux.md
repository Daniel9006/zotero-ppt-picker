# zotero-ppt-picker (macOS / Linux Setup)

This document complements the main README and describes setup steps for macOS and Linux.

---

## Requirements

- macOS or Linux
- Microsoft PowerPoint (desktop, where available)
- Python 3.10+
- Git

---

## Setup (macOS / Linux)

```bash
# clone repository
git clone https://github.com/Daniel9006/zotero-ppt-picker.git
cd zotero-ppt-picker

# create virtual environment
python3 -m venv .venv
source .venv/bin/activate

# install dependencies
python -m pip install -U pip
pip install -r requirements.txt
```

---

## Notes

- On macOS, ensure Python is allowed to control other applications
  (System Settings → Privacy & Security → Automation).
- On Linux, PowerPoint integration may be limited depending on the desktop
  environment and Office version.

---

## Zotero credentials

Credential handling is identical across platforms.  
See the main README section **“Zotero credentials configuration”** for details.
