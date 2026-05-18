# Manual Testing

This document defines the reusable manual retest checklist for alpha releases of
`zotero-ppt-picker`.

The project currently relies on manual PowerPoint retests for citation,
bibliography, persistence, launcher behavior, and user-interface behavior. It
does not yet include a full automated test suite for PowerPoint COM integration.

Use this checklist before tagging a new alpha release and after small
stabilization changes that may affect citation state, bibliography generation,
PowerPoint object handling, launcher startup, or visible user feedback.

---

## Scope

This checklist covers:

- static Python checks
- citation insertion
- bibliography target handling
- the primary document update workflow
- the secondary bibliography rewrite workflow
- deletion and cleanup scenarios
- persistence after saving, closing, and reopening a presentation
- style-specific behavior for the supported base styles
- log inspection
- language-boundary checks between user-facing and maintainer-facing text
- citation insertion on normal slides and in PowerPoint notes
- bibliography target handling on normal slides
- the primary document update workflow with slide and notes citations
- the secondary bibliography rewrite workflow with slide and notes citations
- PowerPoint launcher startup checks when launcher code or documentation changes

This checklist does not cover:

- new citation styles
- separate notes bibliography mode
- locator or page support
- full PowerPoint Ribbon implementation
- signed PPAM deployment
- CSL/style-engine refactoring
- COM/threading refactors unless they were explicitly changed
- packaging or installer validation
- automated CI setup

---

## Static Checks

Run the static checks before starting manual PowerPoint testing.

From the repository root:

```powershell
python -c "import ast, pathlib; ast.parse(pathlib.Path('zotero_picker_ppt.py').read_text(encoding='utf-8')); print('AST parse OK')"
python -m py_compile zotero_picker_ppt.py
git diff --check
```

Expected result:

```text
AST parse OK
py_compile OK
git diff --check clean
```

If the AST command is not used in the local workflow, the minimum required static
check is:

```powershell
python -m py_compile zotero_picker_ppt.py
git diff --check
```

---

## Test Environment Record

Record the test environment for each alpha retest.

| Field | Value |
| --- | --- |
| Release / commit |  |
| Windows version |  |
| PowerPoint version |  |
| Python version |  |
| Zotero library type | user / group |
| Test presentation | new / existing / copied fixture |
| Launcher path tested | direct / cmd / VBA |
| Tester |  |
| Date |  |

Use a copied PowerPoint file for destructive tests such as citation deletion,
text box deletion, intentionally corrupted visible citation text, and launcher
failure-path tests.

---

## Base Style Matrix

Run the base workflow for each supported style.

| Style | Insert citation | Set bibliography target | Update document | Rewrite bibliography | Delete one citation | Delete all citations | Save / close / reopen | Log check |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| APA | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] |
| Harvard | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] |
| IEEE | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] |
| MLA | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] |
| Chicago Author-Date | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] | [ ] |

---

## Core Manual Retest Workflow

For each relevant style, verify the following sequence:

1. Start PowerPoint and open a test presentation.
2. Start the picker from the repository root:

   ```powershell
   python zotero_picker_ppt.py
   ```

3. Select the citation style.
4. Search for a Zotero item.
5. Insert a citation.
6. Set the bibliography target.
7. Run the primary document update workflow.
8. Run the secondary bibliography rewrite workflow.
9. Delete one visible citation and run the primary document update workflow again.
10. Delete all visible citations and run the primary document update workflow again.
11. Edit one visible citation into an intentionally damaged or corrupted form and
    run the primary document update workflow again.
12. Delete the bibliography text box and run the primary document update workflow
    again.
13. Test behavior when no bibliography target has been set.
14. Insert citations first, then set the bibliography target later.
15. Save, close, reopen, and run the primary document update workflow again.
16. Inspect `zotero_ppt.log`.

Expected base result:

- no crash
- no UI hang
- user-facing status, dialog, and UI text remain German
- maintainer-facing logs and diagnostics remain English
- bibliography behavior is deterministic
- citation metadata survives save / close / reopen where supported
- no unexpected traceback appears in the log

---

## PowerPoint Launcher Checks

Use these checks for releases that add or modify launcher behavior.

| Done | Check | Expected result |
| --- | --- | --- |
| [ ] | Start `scripts/start_picker.cmd` from PowerShell. | Picker starts without requiring a manual `cd` into another folder. |
| [ ] | Start `scripts/start_picker.cmd` from `cmd.exe`. | Picker starts through the command launcher. |
| [ ] | Start the launcher from another working directory. | The launcher resolves the repository root relative to itself. |
| [ ] | Import or copy `powerpoint/LaunchZoteroPicker.bas`. | The VBA module imports or copies without syntax changes. |
| [ ] | Set `PICKER_LAUNCHER_PATH` to the local `scripts/start_picker.cmd`. | The macro points to the correct local launcher file. |
| [ ] | Run `LaunchZoteroPicker` from PowerPoint with an active presentation. | The picker starts from PowerPoint. |
| [ ] | Temporarily set a wrong launcher path in the VBA macro. | A clear German message reports that the launcher was not found. |
| [ ] | Temporarily test missing `.venv` or missing Python path. | The command launcher shows a clear German error message. |
| [ ] | Insert a citation on a normal slide after launcher start. | Existing picker behavior is unchanged. |
| [ ] | Insert a citation in PowerPoint notes after launcher start. | Existing notes citation behavior is unchanged. |
| [ ] | Run **Dokument aktualisieren** after launcher start. | Existing document update behavior is unchanged. |
| [ ] | Run **Bibliographie neu schreiben** after launcher start. | Existing bibliography rewrite behavior is unchanged. |
| [ ] | Inspect `zotero_ppt.log`. | No unexpected launcher-related or picker-related errors appear. |

Classify failures as one of:

- launcher problem
- PowerPoint/VBA path problem
- Python/virtual-environment problem
- existing picker/Zotero/network problem

Expected result:

```text
PowerPoint launcher: passed / failed
Notes:
```

---

## Scenario Checklist

| Scenario | Expected result | Result |
| --- | --- | --- |
| Insert citation | The citation appears at the selected text location. | [ ] |
| Set bibliography target | The selected text box becomes the bibliography target. | [ ] |
| Update document | Citations and bibliography state are refreshed. | [ ] |
| Rewrite bibliography | The bibliography is rebuilt from the current citation state. | [ ] |
| Delete one citation | The deleted citation is removed from the bibliography after update. | [ ] |
| Delete all citations | The bibliography is cleared or marked empty after update. | [ ] |
| Corrupt visible citation text | The update workflow handles the damaged citation deterministically without crashing. | [ ] |
| Delete text box | A missing or deleted target is handled without crashing. | [ ] |
| Missing bibliography target | The user receives a clear German user-facing message. | [ ] |
| Late bibliography target setup | The bibliography can be built after citations already exist. | [ ] |
| Save / close / reopen persistence | Stored citation state remains usable after reopening the presentation. | [ ] |
| Log inspection | No unexpected error signatures are present. | [ ] |

---

## Notes Citation Checks

Use these checks for releases that include or may affect notes citation support.

| Done | Check |
| --- | --- |
| [ ] | Insert a citation on a normal slide. |
| [ ] | Insert a citation in PowerPoint notes. |
| [ ] | Verify that both slide and notes citations appear in the bibliography. |
| [ ] | Run **Dokument aktualisieren** with citations on slides and in notes. |
| [ ] | Run **Bibliographie neu schreiben** with citations on slides and in notes. |
| [ ] | Delete a notes citation and run **Dokument aktualisieren**. |
| [ ] | Verify that the deleted notes citation is removed from the bibliography. |
| [ ] | Verify that a bibliography can be built when all citations are only in notes. |
| [ ] | Delete all citations and run **Dokument aktualisieren**. |
| [ ] | Verify that the bibliography anchor remains and the bibliography is cleared. |
| [ ] | Save, close, and reopen the presentation. |
| [ ] | Run **Dokument aktualisieren** again after reopening. |
| [ ] | Verify that notes citation metadata persists after reopening. |

Expected result:

```text
Notes citation support: passed / failed
Notes:
```

---

## Style-Specific Checks

### APA

Verify:

- base author-date citation behavior
- bibliography update after insert/delete
- `a`/`b` disambiguation where applicable
- rollback or rebuild after deleting one item from a disambiguated group
- citation insertion on a normal slide and in notes
- bibliography includes notes citations
- deleting a notes citation removes it from the bibliography after document update
- author-year behavior remains correct across slide and notes citations

Result:

```text
APA: passed / failed
Notes:
```

### Harvard

Verify:

- base author-date citation behavior
- bibliography update after insert/delete
- `a`/`b` disambiguation where applicable
- rollback or rebuild after deleting one item from a disambiguated group
- citation insertion on a normal slide and in notes
- bibliography includes notes citations
- deleting a notes citation removes it from the bibliography after document update
- author-year behavior remains correct across slide and notes citations

Result:

```text
Harvard: passed / failed
Notes:
```

### IEEE

Verify:

- visible numeric citations use consecutive numbering
- deleting a citation renumbers remaining citations
- inserting a citation before an existing citation renumbers following citations
- bibliography numbering follows document order
- no duplicate bibliography labels appear, such as `[1] [1]`
- numbering follows `Slide 1 → Notes 1 → Slide 2`
- expected numbering is `[1]`, `[2]`, `[3]`
- deleting the notes citation renumbers remaining citations to `[1]`, `[2]`
- bibliography numbering remains consistent after document update

Result:

```text
IEEE: passed / failed
Notes:
```

### MLA

Verify:

- new in-text citations do not render as author-date citations
- no unexpected APA/Harvard repair path is applied
- bibliography update works in the current alpha scope
- no locator or page behavior is expected
- citation insertion on a normal slide and in notes
- notes citations do not render as author-date citations
- deleting a notes citation removes it from the bibliography after document update

Result:

```text
MLA: passed / failed
Notes:
```

### Chicago Author-Date

Verify:

- base author-date citation behavior
- bibliography update after insert/delete
- the update workflow does not regress APA, Harvard, IEEE, or MLA behavior
- citation insertion on a normal slide and in notes
- bibliography includes notes citations
- deleting a notes citation removes it from the bibliography after document update
- base resync works with notes citations

Result:

```text
Chicago Author-Date: passed / failed
Notes:
```

---

## Log Inspection

Inspect `zotero_ppt.log` after each retest pass.

Search for:

```text
Worker failed
Traceback
ERROR
RuntimeError
Bibliography not updated
unexpected style paths
stale German debug or maintainer logs
Insert fallback failed
NotesPage
notes fallback
launcher
```

Expected result:

- no unexpected `Worker failed`
- no unexpected `Traceback`
- no unexpected `ERROR`
- no unhandled `RuntimeError`
- no stale German maintainer/debug log messages
- user-facing UI, status, and dialog text remain German
- maintainer-facing logs, comments, and internal diagnostics remain English
- no unexpected `Insert fallback failed`
- no notes-specific traceback
- no notes-specific citation-state loss after save/close/reopen
- no unexpected launcher-related startup issue

Known or intentionally tested error paths should be recorded with context.

---

## Language Boundary Rule

Use this rule during manual review.

| Area | Required language |
| --- | --- |
| User-facing UI labels | German |
| User-facing status text | German |
| User-facing dialogs | German |
| User-facing launcher and VBA error messages | German |
| Maintainer-facing comments | English |
| Maintainer-facing docstrings | English |
| Maintainer-facing logs/debug text | English |
| Maintainer-facing documentation | English |

---

## Manual Retest Summary Template

Use the following template in release notes or release preparation notes.

```text
Manual test result:
- AST parse OK
- py_compile OK
- git diff --check OK
- PowerPoint launcher passed in alpha scope
- APA passed in alpha scope
- Harvard passed in alpha scope
- IEEE passed in alpha scope
- MLA passed in alpha scope
- Chicago Author-Date passed in alpha scope
- Log inspection passed

PowerPoint launcher:
- scripts/start_picker.cmd from PowerShell passed / failed
- scripts/start_picker.cmd from CMD passed / failed
- start from another working directory passed / failed
- PowerPoint VBA macro start passed / failed
- wrong launcher path error handling passed / failed
- missing .venv/Python error handling passed / failed

Notes citation support:
- Insert on slides and in notes passed / failed
- Document update with slide and notes citations passed / failed
- Bibliography rewrite with slide and notes citations passed / failed
- Notes citation deletion cleanup passed / failed
- Notes-only citation scenario passed / failed

Known limitations:
- Separate notes bibliography mode is not included
- Locator/page support is not included
- Full Ribbon/Add-in integration is not included
- Signed PPAM deployment is not included
- EXE packaging and installer support are not included
- No full CSL/style-engine validation
- No automated PowerPoint COM test suite
```

---

## Failure Reporting Template

Use this template when a manual retest fails.

```text
Failure:
Failure class: launcher / VBA path / Python environment / picker / Zotero / network / unknown
Style:
Scenario:
Release / commit:
PowerPoint version:
Python version:
Expected result:
Actual result:
Relevant log excerpt:
Reproducibility:
Notes:
```

Keep log excerpts short and remove private Zotero data before sharing them in
issues, pull requests, or release notes.
