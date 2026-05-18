# PowerPoint Ribbon add-in workflow

This document describes the experimental PowerPoint Ribbon add-in workflow for
starting the existing Zotero PowerPoint picker from a dedicated Ribbon tab.

This is a follow-up to the minimal launcher introduced in `v0.1.0-alpha.20`.
The existing Python picker remains the only citation and bibliography
implementation.

---

## Goal

Provide a PowerPoint Ribbon entry point similar in spirit to the Zotero Word
workflow:

```text
PowerPoint Ribbon tab -> Picker starten button -> VBA callback -> scripts/start_picker.cmd -> zotero_picker_ppt.py
```

The first Ribbon button only starts the existing picker. Future buttons may call
other existing picker workflows after those callbacks are designed separately.

---

## Files

- `powerpoint/LaunchZoteroPicker.bas`
  - VBA module with the launcher macro and Ribbon callback.
- `powerpoint/customUI14.xml`
  - Office Ribbon XML for a custom `Zotero` tab and a `Picker starten` button.
- `scripts/start_picker.cmd`
  - Existing Windows launcher from `v0.1.0-alpha.20`.

---

## Scope

In scope:

- local PowerPoint Ribbon tab source files
- one Ribbon button: **Picker starten**
- callback into the existing VBA launcher macro
- documentation for manual `.ppam` creation and validation

Out of scope:

- signed PPAM deployment
- installer
- automatic path configuration
- EXE packaging
- additional picker workflow buttons
- changes to citation or bibliography logic
- changes to Zotero configuration
- COM/threading refactor

---

## Manual `.ppam` creation

PowerPoint does not provide a built-in editor for custom Ribbon XML. The usual
manual workflow is:

1. Create a new blank PowerPoint presentation.
2. Save it as a macro-enabled presentation, for example:

   ```text
   ZoteroPickerLauncher.pptm
   ```

3. Open the VBA editor with `Alt+F11`.
4. Import `powerpoint/LaunchZoteroPicker.bas`.
5. Adjust `PICKER_LAUNCHER_PATH` in the VBA module so it points to your local
   `scripts/start_picker.cmd` file.
6. Save and close the file.
7. Add `powerpoint/customUI14.xml` to the Office file package as the custom
   Ribbon XML using an Office Ribbon XML editor or equivalent package-editing
   workflow.
8. Save the file as a PowerPoint Add-In:

   ```text
   ZoteroPickerLauncher.ppam
   ```

9. Load the add-in in PowerPoint:

   ```text
   File -> Options -> Add-ins -> Manage: PowerPoint Add-ins -> Go... -> Add New...
   ```

10. Select and enable `ZoteroPickerLauncher.ppam`.

After loading the add-in, PowerPoint should show a `Zotero` tab with one button:

```text
Picker starten
```

---

## Expected behavior

When the user clicks **Picker starten**:

1. PowerPoint calls the Ribbon callback `LaunchZoteroPickerRibbon`.
2. The callback delegates to `LaunchZoteroPicker`.
3. The existing command launcher starts the existing Python picker.
4. Picker behavior after startup is unchanged.

---

## Retest checklist

| Done | Check | Expected result |
| --- | --- | --- |
| [ ] | Create local `.pptm` with `LaunchZoteroPicker.bas`. | VBA module imports without syntax errors. |
| [ ] | Add `customUI14.xml` as Ribbon XML. | PowerPoint accepts the custom UI. |
| [ ] | Save/load as `.ppam`. | Add-in loads without startup error. |
| [ ] | Open a normal `.pptx`. | The `Zotero` Ribbon tab is visible. |
| [ ] | Click **Picker starten**. | Existing picker starts. |
| [ ] | Insert a citation on a normal slide. | Existing picker behavior is unchanged. |
| [ ] | Insert a citation in notes. | Existing notes behavior is unchanged. |
| [ ] | Run **Dokument aktualisieren**. | Existing document update behavior is unchanged. |
| [ ] | Run **Bibliographie neu schreiben**. | Existing bibliography rewrite behavior is unchanged. |
| [ ] | Restart PowerPoint. | Add-in remains available if enabled. |
| [ ] | Inspect `zotero_ppt.log`. | No unexpected Ribbon/add-in-related error appears. |

---

## Manual retest result

Result for the current alpha scope:

- Static checks: PASS.
- `git diff --check`: PASS.
- No changes to `zotero_picker_ppt.py`: PASS.
- Local `.pptm` with Ribbon XML opens without startup callback error: PASS.
- `Zotero` Ribbon tab appears in the local `.pptm`: PASS.
- `Picker starten` button starts the existing picker: PASS.
- Command-window flash was removed by hiding the launcher process from VBA: PASS.
- `.ppam` add-in can be created and loaded manually: PASS.
- `Zotero` Ribbon tab appears in a normal `.pptx`: PASS.
- `Picker starten` starts the existing picker from a normal `.pptx`: PASS.
- APA slide citation after Ribbon/PPAM start: PASS.
- Bibliography target and update after Ribbon/PPAM start: PASS.
- APA notes citation after Ribbon/PPAM start: PASS.
- Automatic bibliography update after notes citation insert: PASS.
- PowerPoint restart persistence: PASS.
- Log inspection: PASS.

Known test notes:

- The `.ppam` file was created manually and is not committed to the repository.
- The add-in was tested as an unsigned local add-in.
- Only the first Ribbon button, **Picker starten**, is included in this scope.
- Direct Ribbon buttons for **Dokument aktualisieren**, **Bibliographie neu schreiben**, and **Bibliographie-Ziel festlegen** remain follow-up work.

---

## Known limitations

- The `.ppam` file is created manually.
- The add-in is not signed.
- Users may need to adjust PowerPoint Trust Center settings.
- The local launcher path is still edited manually in VBA.
- Only one Ribbon button is included in the initial scope.
