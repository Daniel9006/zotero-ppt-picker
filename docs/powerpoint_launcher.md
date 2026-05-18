# PowerPoint picker launcher

This document describes the minimal Windows/PowerPoint launcher introduced for
`v0.1.0-alpha.20`.

The launcher starts the existing Python picker from PowerPoint. It does not add a
new citation engine, does not change bibliography handling, and does not replace
the existing `zotero_picker_ppt.py` workflow.

---

## Architecture

```text
PowerPoint VBA macro
  -> scripts/start_picker.cmd
     -> zotero_picker_ppt.py
```

Responsibilities:

- `powerpoint/LaunchZoteroPicker.bas` contains a small VBA macro template.
- `scripts/start_picker.cmd` resolves the repository root relative to itself.
- The command launcher changes into the repository root before starting Python.
- The existing `zotero_picker_ppt.py` application remains the only picker and
  citation/bibliography implementation.

This separation keeps PowerPoint integration independent from citation logic.

---

## Requirements

- Windows
- Microsoft PowerPoint desktop version
- Python and the project virtual environment set up as described in `README.md`
- Project dependencies installed via `requirements.txt`
- Zotero credentials configured through the existing application flow

The recommended runtime path is:

```text
.venv\Scripts\pythonw.exe
```

The launcher can also fall back to:

```text
.venv\Scripts\python.exe
pyw.exe
py.exe
```

The `.venv` path remains the preferred and most predictable option.

---

## Setup in PowerPoint

1. Open PowerPoint.
2. Open the VBA editor with `Alt+F11`.
3. Import `powerpoint/LaunchZoteroPicker.bas`, or copy the macro into a module.
4. Edit the local launcher path in the VBA module:

   ```vb
   Private Const PICKER_LAUNCHER_PATH As String = "C:\Path\To\zotero-ppt-picker\scripts\start_picker.cmd"
   ```

5. Save the macro-enabled presentation or add-in file according to your local
   PowerPoint macro workflow.
6. Run `LaunchZoteroPicker` from PowerPoint.

Optional: assign the macro to the PowerPoint Quick Access Toolbar or to a custom
Ribbon button through PowerPoint's built-in customization UI.

---

## Usage

1. Open PowerPoint and a presentation.
2. Place the text cursor where the citation should be inserted.
3. Run the `LaunchZoteroPicker` macro.
4. Use the existing picker as before.

The picker behavior after startup is unchanged. Citation insertion, notes
citation support, **Dokument aktualisieren**, and **Bibliographie neu schreiben**
continue to use the existing Python implementation.

---

## Troubleshooting

### Launcher file not found

If the macro reports that the launcher was not found, update
`PICKER_LAUNCHER_PATH` in `powerpoint/LaunchZoteroPicker.bas` so it points to
your local `scripts/start_picker.cmd` file.

### Missing `.venv`

Create the virtual environment from the repository root:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Python not found

Use the project virtual environment whenever possible. If `.venv` is missing,
the launcher tries `pyw.exe` and `py.exe`, but those fallback interpreters may
not have the required dependencies installed.

### Picker starts, but PowerPoint is not active

Open PowerPoint, activate the target presentation, and place the text cursor in
a text box before inserting a citation.

### Zotero configuration is missing

The existing picker opens the Zotero credentials dialog when no valid local
configuration is available. Credentials are still stored through the existing
configuration flow; the launcher does not modify configuration files.

---

## Out of scope

This launcher is intentionally minimal. It is not:

- a full Office Ribbon implementation
- a signed PPAM deployment
- an EXE package
- an installer
- a new citation engine
- a Zotero configuration mechanism
- a COM/threading refactor

Locator/page support, style-engine refactoring, and bibliography model changes
remain separate future work.
