# PowerPoint picker launcher and Ribbon actions

This document describes the Windows/PowerPoint launcher integration for
`zotero-ppt-picker`.

The launcher was introduced in `v0.1.0-alpha.20` to start the existing Python
picker from PowerPoint. It was extended later with Ribbon/CLI actions for common
PowerPoint workflows.

The PowerPoint integration does not add a separate citation engine. It calls the
existing `zotero_picker_ppt.py` implementation and keeps citation, bibliography,
state, and Zotero access logic in Python.

---

## Architecture

```text
PowerPoint Ribbon button
  -> VBA callback in powerpoint/LaunchZoteroPicker.bas
     -> scripts/start_picker.cmd [optional --action ...]
        -> zotero_picker_ppt.py
           -> shared workflow implementation
```

The launcher supports two modes:

```text
scripts/start_picker.cmd
```

Starts the full Picker UI.

```text
scripts/start_picker.cmd --action set-bibliography-target
scripts/start_picker.cmd --action rewrite-bibliography
scripts/start_picker.cmd --action update-document
```

Runs a specific PowerPoint action without opening the Picker UI.

---

## Responsibilities

- `powerpoint/customUI14.xml` defines the custom PowerPoint Ribbon tab.
- `powerpoint/LaunchZoteroPicker.bas` contains the VBA callbacks.
- `scripts/start_picker.cmd` resolves the repository root relative to itself and
  starts Python from a stable working directory.
- `zotero_picker_ppt.py` remains the only citation and bibliography
  implementation.
- CLI/Ribbon actions call the same shared Python workflow functions that are used
  by the Picker-App buttons.

This avoids a parallel implementation in VBA or in the command launcher.

---

## Ribbon buttons

The Ribbon tab is named `Zotero` and contains these buttons:

```text
Zitationen
- Zitation einfuegen

Dokument
- Dokument aktualisieren

Bibliographie
- Bibliographie neu schreiben
- Bibliographie-Ziel festlegen
```

In the XML file, German umlauts are encoded as XML character references, for
example:

```xml
label="Zitation einf&#x00FC;gen"
```

PowerPoint displays this as `Zitation einfügen`, while the XML file remains
ASCII-safe and avoids encoding issues.

---

## VBA callbacks

`powerpoint/LaunchZoteroPicker.bas` exposes these Ribbon callbacks:

```vb
LaunchZoteroPicker
ZoteroUpdateDocument
ZoteroRewriteBibliography
ZoteroSetBibliographyTarget
```

The callbacks delegate to a single internal helper that calls
`scripts/start_picker.cmd` with the appropriate optional action argument.

Expected mapping:

```text
LaunchZoteroPicker              -> scripts/start_picker.cmd
ZoteroUpdateDocument            -> scripts/start_picker.cmd --action update-document
ZoteroRewriteBibliography       -> scripts/start_picker.cmd --action rewrite-bibliography
ZoteroSetBibliographyTarget     -> scripts/start_picker.cmd --action set-bibliography-target
```

---

## Runtime behavior

The command launcher prefers:

```text
.venv\Scripts\pythonw.exe
```

It can fall back to:

```text
.venv\Scripts\python.exe
pyw.exe
py.exe
```

The `.venv` path remains the preferred and most predictable option because it
uses the project dependencies installed from `requirements.txt`.

For Ribbon/CLI actions, `zotero_picker_ppt.py` keeps a hidden Tk event loop alive
and runs the actual workflow in a worker thread. This matches the Picker-App
execution model more closely than a synchronous command-line call and avoids
PowerPoint COM instability observed in headless action runs.

### Launcher window behavior

Ribbon buttons are started through `scripts/start_picker.cmd`, but the VBA
launcher hides the transient command window. Users should not see a flashing
console window when clicking Ribbon buttons.

For the Picker UI button, the VBA launcher first checks whether a Picker window
is already open. If an existing `Zotero Picker` window is found, it is brought to
the foreground instead of starting a second Picker instance.

---

## Setup in PowerPoint

1. Open the macro-enabled PowerPoint file or add-in project.
2. Open the VBA editor with `Alt+F11`.
3. Import or update `powerpoint/LaunchZoteroPicker.bas`.
4. Verify the local project path in the module:

   ```vb
   Private Const PROJECT_ROOT As String = "C:\Users\daniel\OneDrive\Zotero_Add-In\Python\zotero-ppt-picker"
   ```

5. Add or update `powerpoint/customUI14.xml` with Office RibbonX Editor.
6. Save the `.pptm` or `.ppam`.
7. Close PowerPoint completely.
8. Reopen PowerPoint and verify that the `Zotero` Ribbon tab is visible.

For `.ppam` deployment, load the add-in through:

```text
File -> Options -> Add-ins -> Manage: PowerPoint Add-ins -> Go... -> Add New...
```

---

## Manual launcher checks

Run these checks from the repository root with an active PowerPoint presentation:

```powershell
.\scripts\start_picker.cmd
```

Expected result: the Picker UI opens.

For actions:

```powershell
Remove-Item .\zotero_ppt.log -ErrorAction SilentlyContinue
.\scripts\start_picker.cmd --action set-bibliography-target
```

```powershell
Remove-Item .\zotero_ppt.log -ErrorAction SilentlyContinue
.\scripts\start_picker.cmd --action rewrite-bibliography
```

```powershell
Remove-Item .\zotero_ppt.log -ErrorAction SilentlyContinue
.\scripts\start_picker.cmd --action update-document
```

Expected log marker for action runs:

```text
COM enter(no-lock): cli-action-worker:<action-name>
```

---

## Ribbon retest checklist

Use a test presentation with at least one citation and a bibliography text field.

### Bibliographie-Ziel festlegen

1. Select the bibliography text box in PowerPoint so the frame is visible.
2. Click `Bibliographie-Ziel festlegen`.

Expected result:

```text
Bibliographie-Ziel gesetzt. Gefundene Anker: 1 (Bibliographie aktualisiert).
```

Expected log markers:

```text
Anchor selection:
Anchor saved:
Resolve anchors: found=1
Bib entries generated:
Bib write: placed=
Bib update OK:
```

### Bibliographie neu schreiben

Click `Bibliographie neu schreiben`.

Expected result:

```text
Bibliographie aktualisiert (<style label>).
```

Expected log markers:

```text
cli-action-worker:rewrite-bibliography
Bib update: anchors=1
Bib entries generated:
Bib write: placed=
Bib update OK:
```

### Dokument aktualisieren

Click `Dokument aktualisieren`.

Expected result:

```text
Aktualisiert: <n> Zitat(e) im Dokument.
```

Expected log markers:

```text
cli-action-worker:update-document
DocumentUpdate: keys_after_prune=
Bib update OK:
```

### Zitation einfuegen

Click `Zitation einfuegen`.

Expected result:

- the Picker UI opens,
- a citation can be selected and inserted,
- the bibliography is automatically updated if a bibliography target exists.

---

## Confirmed test status

The current implementation has been manually tested with:

```text
CLI actions
- --action set-bibliography-target
- --action rewrite-bibliography
- --action update-document

Launcher actions
- scripts/start_picker.cmd --action set-bibliography-target
- scripts/start_picker.cmd --action rewrite-bibliography
- scripts/start_picker.cmd --action update-document

PowerPoint .pptm Ribbon buttons
- Zitation einfuegen
- Dokument aktualisieren
- Bibliographie neu schreiben
- Bibliographie-Ziel festlegen

PowerPoint .ppam Ribbon buttons
- Zitation einfuegen
- Dokument aktualisieren
- Bibliographie neu schreiben
- Bibliographie-Ziel festlegen
```

---

## Troubleshooting

### Zotero tab is not visible

Close and reopen PowerPoint. If the tab is still missing, verify that the `.pptm`
or `.ppam` contains the updated `customUI14.xml` and that the add-in is loaded.

### Macro callback not found

Verify that the `onAction` names in `customUI14.xml` exactly match the public VBA
Sub names in `LaunchZoteroPicker.bas`.

### Launcher file not found

Verify the `PROJECT_ROOT` constant in `powerpoint/LaunchZoteroPicker.bas` and
confirm that this file exists:

```text
scripts\start_picker.cmd
```

### Python or dependency errors

Use the project virtual environment whenever possible:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Action starts but no bibliography is written

Check `zotero_ppt.log`. A successful action should include:

```text
cli-action-worker:
Resolve anchors: found=1
Bib write: placed=
Bib update OK:
```

If no bibliography target is found, select the bibliography text box and run
`Bibliographie-Ziel festlegen`.

### Picker is already open but hidden behind PowerPoint

Click `Zitation einfuegen` again. The launcher should bring the existing Picker
window to the foreground instead of opening a second Picker window.

If this does not work, check whether the Picker window title still starts with
`Zotero Picker`, because the VBA launcher uses this title for window activation.

---

## Out of scope

This launcher and Ribbon integration is not:

- a new citation engine
- a separate bibliography implementation
- a Zotero configuration mechanism
- an installer
- a signed deployment package

Locator/page support, deeper CSL/style-engine refactoring, and bibliography model
changes remain separate future work.
