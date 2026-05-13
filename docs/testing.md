# Manual Testing

This document defines the reusable manual retest checklist for alpha releases of
`zotero-ppt-picker`.

The project currently relies on manual PowerPoint retests for citation,
bibliography, persistence, and user-interface behavior. It does not yet include a
full automated test suite for PowerPoint COM integration.

Use this checklist before tagging a new alpha release and after small
stabilization changes that may affect citation state, bibliography generation,
PowerPoint object handling, or visible user feedback.

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

This checklist does not cover:

- new citation styles
- notes citation support
- locator or page support
- PowerPoint launcher or locator features
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
```

Expected result:

```text
AST parse OK
py_compile OK
```

If the AST command is not used in the local workflow, the minimum required static
check is:

```powershell
python -m py_compile zotero_picker_ppt.py
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
| Tester |  |
| Date |  |

Use a copied PowerPoint file for destructive tests such as citation deletion,
text box deletion, and intentionally corrupted visible citation text.

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

## Style-Specific Checks

### APA

Verify:

- base author-date citation behavior
- bibliography update after insert/delete
- `a`/`b` disambiguation where applicable
- rollback or rebuild after deleting one item from a disambiguated group

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
```

Expected result:

- no unexpected `Worker failed`
- no unexpected `Traceback`
- no unexpected `ERROR`
- no unhandled `RuntimeError`
- no stale German maintainer/debug log messages
- user-facing UI, status, and dialog text remain German
- maintainer-facing logs, comments, and internal diagnostics remain English

Known or intentionally tested error paths should be recorded with context.

---

## Language Boundary Rule

Use this rule during manual review.

| Area | Required language |
| --- | --- |
| User-facing UI labels | German |
| User-facing status text | German |
| User-facing dialogs | German |
| Maintainer-facing comments | English |
| Maintainer-facing docstrings | English |
| Maintainer-facing logs/debug text | English |

---

## Manual Retest Summary Template

Use the following template in release notes or release preparation notes.

```text
Manual test result:
- AST parse OK
- py_compile OK
- APA passed in alpha scope
- Harvard passed in alpha scope
- IEEE passed in alpha scope
- MLA passed in alpha scope
- Chicago Author-Date passed in alpha scope
- Log inspection passed

Known limitations:
- Notes citation support is not included
- Locator/page support is not included
- PowerPoint launcher/locator support is not included
- No full CSL/style-engine validation
- No automated PowerPoint COM test suite
```

---

## Failure Reporting Template

Use this template when a manual retest fails.

```text
Failure:
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
