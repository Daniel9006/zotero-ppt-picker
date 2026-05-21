# Versioning & Release Ladder (Zotero ↔ PowerPoint)

This repository uses a pragmatic, small-team-friendly versioning approach:

- clear milestones
- safe rollbacks
- minimal process overhead
- no CI assumptions

**Tags are the source of truth.**

Current public baseline: `v0.1.0-alpha.23`

Current development focus:
- technical stabilization of citation and bibliography mechanics
- persistent citation state and document resync reliability
- PowerPoint Ribbon and CLI actions for existing picker workflows
- improving PowerPoint launcher usability without adding duplicate workflow logic

---

## Tag format

We use SemVer-style tags:

- Stable: `vMAJOR.MINOR.PATCH` (e.g. `v0.2.0`)
- Pre-release: `vMAJOR.MINOR.PATCH-alpha.N`, `...-beta.N`, `...-rc.N`
  - Example: `v0.1.0-alpha.2`

Rules of thumb:
- **MINOR** = big step (new citation style, new platform support, major refactor)
- **PATCH** = fixes / small improvements, no new major features
- Pre-releases are used until the milestone is stable.

---

## Definitions (Definition of Done per maturity level)

### Alpha
Feature is implemented and demonstrable.  
Known issues may exist (including occasional crashes), but progress is trackable.

### Beta
Usable for real-world testing.  
Known issues are documented, error handling is deterministic, and the codebase follows
the coding standards.

### Release Candidate (RC)
No new features. Only bug fixes and stabilization.  
No known showstoppers in the supported scope.

### Stable
Happy path is robust.  
No known showstoppers. Errors are handled gracefully (no UI hangs, clear messages).

---

## Official Release Ladder

### Phase 0 — Windows foundation (Config + APA + COM)

#### `v0.1.0-alpha.N` (current stage)
Scope:
- Config / credentials flow implemented and stable enough for daily use
- APA citation style implemented
- IEEE technical alpha support introduced for numeric citation processing
- Picker / PowerPoint COM may still show intermittent issues
- Minimal Windows/PowerPoint launcher support may be added as small alpha features

Notes:
- Migration of maintainer-facing comments, docstrings, and debug/log messages to English has started.
- User-facing UI labels, status messages, and dialogs remain German.

#### Hotfix patch releases (run blockers)
If the project becomes non-runnable due to a packaging, environment, or import issue,
ship a **PATCH** release immediately (e.g. `v0.1.0-alpha.3`).

Typical examples:
- Missing dependency (`pywin32`)
- Wrong Python interpreter / virtual environment not activated
- Import-time failures caused by missing packages

No feature work is bundled into such hotfixes.

#### `v0.1.0-beta.N` (Quality Gate 1)
This is the first quality-enforced milestone.

Required gates:
1. English maintainer-facing comments and docstrings
2. Specialized exceptions (no broad `except Exception` in core logic)
3. Type hints and docstrings instead of commenting every variable

Goal:
- COM issues are reduced, reproducible, and diagnosable

#### `v0.1.0-rc.N`
No new features. Stabilization only.

Required:
- Deterministic COM behavior
- No UI hangs or dead windows

#### `v0.1.0` (Windows + APA stable)
Definition of done:
- Windows happy path robust (Picker → Insert)
- COM issues fixed or gracefully handled
- APA citation style stable

---

### Phase 1 — Architecture refactor (citation style deduplication)

Goal:
Prevent copy/paste growth across citation styles.

Scope (no new citation styles):
- Shared citation formatting / rendering layer (“style engine”)
- Refactor `zotero_config.py` into a constructor/factory with validation

Gate:
- The style engine should exist before promoting additional citation styles beyond alpha-level support.

---

#### Refactor milestones (within versioning phases)

Refactors that improve architecture, reduce duplication, and prepare
for future features are gated by versioning phases. They do *not*
introduce new citation styles or platform support, but they raise quality
and maintainability.

- **Before Phase 1 complete:** style duplication and config structure
  may be inconsistent; planned refactors are allowed in feature branches.
- **Phase 1 gate (architecture refactor):**
  - Shared citation style engine introduced
  - `zotero_config.py` refactored into constructor/factory with validation
  - Deduplicated logic across citation styles
- Once Phase 1 goals are met, additional styles and style hardening
  (Chicago/Harvard/MLA and further IEEE validation) may continue
  following alpha→beta→rc→stable cycles.

These architecture refactors may span multiple pre-releases
(e.g. `v0.1.1-alpha`, `v0.1.2-beta`) and do not require delaying
smaller bugfix patches in Phase 0.

### Phase 2 — Citation style stabilization (each major style is a MINOR bump)

Each style is treated as a major milestone.

- `v0.2.0` — IEEE stabilization after alpha-level technical support
- `v0.3.0` — Chicago
- `v0.4.0` — Harvard

Each follows: alpha → beta → rc → stable.

Patch releases (`v0.2.x`, `v0.3.x`, `v0.4.x`) are for:
- bug fixes
- documentation
- tests
- small UX improvements

---

### Phase 3 — Platform expansion

#### `v0.5.0` — macOS preview

Prerequisites:
- Config handling is cleanly encapsulated
- Citation style engine is stable

macOS support is treated as a major milestone with its own stabilization cycle.

---

## Release documentation

This section documents verified alpha changes, fixes, and runtime-relevant behavior.

### 2026-01-16 — win32com / pywin32 not found (Windows)

**Symptoms**
- `NameError: name 'win32' is not defined`
- `ModuleNotFoundError: No module named 'win32com'`

**Root cause**
- The script was executed with a Python interpreter where required dependencies
  were not installed.
- `win32com` is provided by **pywin32**, which is not part of standard Python.

**Fix type**
- Environment / dependency fix
- No business-logic code bug

**Resolution**
- Ensure `pywin32` is installed for the Python interpreter used to start the script.
- Verify imports before running the application.

**Result**
- Script starts correctly
- PowerPoint COM automation is available
- Expected behavior when no presentation is open

**Related tags**
- Documentation added in `v0.1.0-alpha.1`
- Versioning and coding standards baseline in `v0.1.0-alpha.2`

### v0.1.0-alpha.15 — IEEE citation state and bibliography renumbering

**Scope**
- Technical IEEE blocker fix
- No general style-engine refactor
- No broad citation-style matrix changes

**Symptoms fixed**
- First IEEE citation inserted correctly as `[1]`, but follow-up citations could be inserted again as `[1]`
- Manual bibliography update could lose visible IEEE citations
- Cleanup could report no citations although visible `[n]` citations remained
- Late bibliography anchor setup could fail to rebuild the IEEE bibliography
- IEEE bibliography entries could show duplicate labels such as `[1] [1] ...`
- Inserting a new IEEE citation before existing citations did not renumber following citations correctly

**Root cause**
- IEEE initially used visible placeholders such as `⟦zp:KEY⟧`.
- These placeholders were replaced by visible numeric citations (`[1]`, `[2]`, ...).
- After replacement, the permanently reconstructable citation state was lost.
- IEEE numbering was based on placeholder scan order instead of persisted citation metadata and visible document order.
- Zotero returned IEEE bibliography entries with their own local numeric label, which conflicted with document-level numbering.

**Resolution**
- Persist IEEE citations in PowerPoint shape tags (`ZP_CITES`).
- Build IEEE numbering from stored cite records.
- Sort IEEE cite records by visible text position within each shape.
- Renumber visible citations and stored metadata together.
- Normalize Zotero-provided IEEE bibliography labels before applying document-level numbering.
- Support manual bibliography update, cleanup, and late bibliography anchor setup.
- Disable bibliography bullet formatting when writing entries.

**Manual test result**
- IEEE citations are numbered consecutively.
- Inserting a citation before existing citations renumbers following citations.
- Bibliography updates automatically and manually without losing entries.
- Zotero-provided IEEE labels are normalized to avoid duplicate numbering.
- Cleanup after partial and full deletion works.
- Late bibliography anchor setup rebuilds the IEEE bibliography.

### v0.1.0-alpha.16 – Base style matrix validation and UI label polish

This alpha release documents the completed base citation style validation for
the current supported styles and improves user-facing style status messages.

Changes:
- documented the base citation style matrix results in `docs/development.md`
- confirmed APA regression status
- confirmed IEEE as technically alpha-stable and broadly plausible
- confirmed Chicago Author-Date and Harvard as passed for the base alpha scope
- documented MLA as technically passed but requiring style-specific in-text rendering follow-up
- documented locator/detail references as future work
- removed internal citation style IDs from user-facing status messages

Notes:
- internal style IDs remain available in debug logs
- locator support for pages, chapters, clauses, figures, and tables is not part of this release
- MLA in-text rendering remains a follow-up topic 

### v0.1.0-alpha.17 – MLA in-text rendering fix

**Scope**
- Minimal MLA-specific in-text citation rendering fix
- Small user-facing bibliography UI label polish
- No COM, threading, anchor, or citation-state refactor

**Symptoms fixed**
- MLA in-text citations were technically stable but rendered in an author-date-like pattern, e.g. `(Author, 2024)` or `(Author, n.d.)`.
- Bibliography update status messages could expose internal style IDs instead of display labels.

**Resolution**
- Added a minimal MLA-specific in-text formatter for new inserts.
- MLA citations now render as minimal parenthetical labels, e.g. `(Author)`, `(Author and Author)`, or `(Corporate Author)`.
- Bibliography update status now uses display style names.
- Bibliography and cleanup button labels were clarified.

**Manual test result**
- APA, IEEE, Chicago Author-Date, Harvard, and MLA were retested.
- MLA insert, bibliography update, cleanup, and repeated manual bibliography refresh passed in the alpha scope.
- MLA no longer renders new in-text citations as author-date citations.

**Known limitations**
- No locator/page support.
- No full CSL/style-engine refactor.
- Existing MLA citations inserted with older versions are not migrated automatically.
- MLA disambiguation for identical visible labels is not implemented yet.

### v0.1.0-alpha.18 – Document update workflow and language cleanup

**Scope**
- User-facing workflow clarification for document-level maintenance.
- Primary workflow button renamed to **Dokument aktualisieren**.
- Secondary bibliography-only workflow exposed as **Bibliographie neu schreiben**.
- Maintainer-facing comments, docstrings, and debug/log messages cleaned up toward English.
- User-facing UI labels, status messages, and dialogs remain German.

**Workflow changes**
- **Dokument aktualisieren** is the primary workflow for resynchronizing visible citations with stored citation state.
- It removes deleted citations from the bibliography.
- It clears the bibliography when no citations remain.
- It tolerates missing bibliography targets.
- It repairs APA/Harvard suffix disambiguation.
- It runs IEEE renumbering.
- It performs only the base resync for MLA and Chicago Author-Date.

**Bibliography-only workflow**
- **Bibliographie neu schreiben** rewrites the bibliography when a bibliography target exists.
- It is not the primary repair or cleanup workflow.
- It does not primarily change visible citations.
- APA/Harvard suffix repair is handled by the document update workflow.
- IEEE renumbering is handled by the document update workflow.

**Not included**
- No COM or threading changes.
- No PowerPoint anchor changes.
- No Zotero Web API changes.
- No bibliography-state refactor.
- No CSL/style-engine refactor.
- No locator/page support.

**Static checks**
- AST parse OK.
- `py_compile` OK.

**Manual retest result**
- APA: PASS.
- Harvard: PASS.
- IEEE: PASS.
- MLA: PASS.
- Chicago Author-Date: PASS.

**Overall result**
- `v0.1.0-alpha.18` retest: PASS.

**Verified risks**
- Button workflow changed, but remains functionally stable.
- Deleted citations are cleaned up correctly.
- Bibliography is updated or cleared correctly.
- APA/Harvard disambiguation remains correct.
- IEEE renumbering remains correct.
- MLA/Chicago base behavior remains correct.
- Missing bibliography targets do not crash.
- Persistence after save, close, and reopen works.
- Maintainer-facing comments, docstrings, and debug/log messages are English.
- German user-facing UI text is preserved.
- Logs are clean.

### v0.1.0-alpha.19 – Notes citation support

**Scope**
- Minimal notes citation support for PowerPoint speaker notes.
- Document-wide citation scans now include normal slide shapes and NotesPage shapes.
- Document order is defined as `Slide 1 → Notes 1 → Slide 2 → Notes 2 → …`.
- Notes citations contribute to bibliography rebuilds and document updates.
- The bibliography anchor remains limited to normal slide shapes.

**Technical changes**
- Added a central document-wide citation shape iteration path.
- Document-wide citation operations now include slide shapes first, then notes shapes for the same slide.
- NotesPage-aware citation scanning is used by:
  - `collect_all_cites_by_key()`
  - `collect_all_cite_texts()`
  - `normalize_sig_group(...)`
  - `renormalize_all_sig_groups()`
  - `build_ieee_numbering_from_document()`
  - `resync_bibliography_keys_from_document(...)`
  - `renumber_ieee_and_update(...)`
- Insert into notes uses a fallback path when PowerPoint does not expose a reliable shape via the normal selection path:
  - insert a temporary marker
  - scan the slide and NotesPage shapes
  - find the shape containing the marker
  - remove the marker
  - store `ZP_CITES` metadata on the detected notes shape

**Behavior**
- Notes citations are included in **Dokument aktualisieren**.
- Notes citations are included in **Bibliographie neu schreiben**.
- Deleting a notes citation removes it from the bibliography after document update.
- IEEE numbering follows the document order `Slide 1 → Notes 1 → Slide 2 → Notes 2 → …`.
- If all citations exist only in notes, the bibliography can still be built from those citations.
- If all citations are deleted, the existing slide-based bibliography anchor remains and the bibliography is cleared.

**Not included**
- No locator/page support.
- No PowerPoint launcher or Ribbon integration.
- No CSL/style-engine refactor.
- No COM/threading refactor.
- No Zotero Web API changes.
- No separate notes bibliography mode.
- No change to the German user-facing UI language.

**Manual retest result**
- APA: PASS.
- Harvard: PASS.
- IEEE: PASS.
- MLA: PASS.
- Chicago Author-Date: PASS.
- Edge cases with notes-only citations and full deletion: PASS.

**Log assessment**
- Earlier errors were from pre-final insert fallback attempts.
- Final successful runs did not show new relevant `Worker failed`, `Traceback`, or `Insert fallback failed` entries.
- One IEEE error was caused by missing internet/DNS connectivity and was not related to notes or renumbering logic.

**Overall result**
- `v0.1.0-alpha.19 – Notes citation support`: release-ready.

### v0.1.0-alpha.20 – PowerPoint picker launcher

**Scope**
- Minimal Windows/PowerPoint launcher for starting the existing picker from PowerPoint.
- Added `scripts/start_picker.cmd` as the repository-relative Windows command launcher.
- Added `powerpoint/LaunchZoteroPicker.bas` as a PowerPoint VBA macro template.
- Added `docs/powerpoint_launcher.md` with setup, usage, and troubleshooting notes.
- Documented the launcher architecture in README, development notes, testing notes, and release documentation.

**Architecture**

```text
PowerPoint VBA macro → scripts/start_picker.cmd → zotero_picker_ppt.py
```

The launcher changes into the repository root before starting the existing picker.
It prefers `.venv\Scripts\pythonw.exe`, with fallback options for `.venv\Scripts\python.exe`, `pyw.exe`, and `py.exe`.

**Behavior**
- The existing Python picker remains the only citation and bibliography implementation.
- The launcher does not modify Zotero configuration files.
- The launcher does not change citation state handling, notes support, bibliography generation, or document update behavior.
- User-facing launcher and VBA error messages are German.

**Not included**
- No full Office Ribbon implementation.
- No signed PPAM deployment.
- No EXE packaging.
- No installer.
- No citation or bibliography logic changes.
- No Zotero configuration changes.
- No COM/threading refactor.
- No locator/page support.

**Manual retest result**
- Static checks: PASS.
- `git diff --check`: PASS.
- No changes to `zotero_picker_ppt.py`: PASS.
- Start via `scripts/start_picker.cmd` from PowerShell: PASS.
- Start via `scripts/start_picker.cmd` from CMD: PASS.
- Start from a different working directory: PASS.
- Start through the PowerPoint VBA macro: PASS.
- Wrong VBA launcher path shows an understandable German error message: PASS.
- APA slide citation after launcher start: PASS.
- APA notes citation after launcher start: PASS.
- Automatic bibliography update after notes citation insert: PASS.
- **Dokument aktualisieren** after launcher start: PASS.
- **Bibliographie neu schreiben** after launcher start: PASS.
- Log inspection: PASS.

**Not destructively tested**
- Missing `.venv` / missing Python fallback behavior.
- Reason: this would require temporary environment manipulation and may be masked by `pyw.exe` / `py.exe` fallbacks on the test system.

**Overall result**
- `v0.1.0-alpha.20 – PowerPoint picker launcher`: release-ready.

### v0.1.0-alpha.21 – PowerPoint Ribbon picker button

**Scope**
- Minimal PowerPoint Ribbon entry point for starting the existing picker from a dedicated `Zotero` tab.
- Added `powerpoint/customUI14.xml` as the Office Ribbon XML source for the `Zotero` tab and **Picker starten** button.
- Extended `powerpoint/LaunchZoteroPicker.bas` with a Ribbon callback for the new button.
- Added `docs/powerpoint_ribbon_addin.md` with the manual `.pptm`/`.ppam` creation and validation workflow.
- Kept the existing `scripts/start_picker.cmd` launcher as the process-start boundary.

**Architecture**

```text
PowerPoint Ribbon tab → Picker starten → VBA callback → scripts/start_picker.cmd → zotero_picker_ppt.py
```

The Ribbon button delegates to the existing VBA launcher macro, which starts the existing command launcher. The existing Python picker remains the only citation and bibliography implementation.

**Behavior**
- The `Zotero` tab is available after loading the manually created `.ppam` add-in.
- The **Picker starten** button starts the existing picker from normal `.pptx` presentations.
- The picker starts without a visible command-window flash from the Ribbon button.
- Citation insertion, notes support, bibliography target handling, **Dokument aktualisieren**, and **Bibliographie neu schreiben** remain implemented by the existing picker.
- User-facing Ribbon labels and launcher error messages are German.

**Not included**
- No direct Ribbon buttons for **Dokument aktualisieren**, **Bibliographie neu schreiben**, or **Bibliographie-Ziel festlegen**.
- No signed PPAM deployment.
- No installer.
- No automatic local path configuration.
- No EXE packaging.
- No citation or bibliography logic changes.
- No Zotero configuration changes.
- No COM/threading refactor.
- No locator/page support.

**Manual retest result**
- Static checks: PASS.
- `git diff --check`: PASS.
- No changes to `zotero_picker_ppt.py`: PASS.
- Local `.pptm` with Ribbon XML opens without startup callback error: PASS.
- `Zotero` Ribbon tab appears in the local `.pptm`: PASS.
- **Picker starten** button starts the existing picker: PASS.
- Command-window flash was removed by hiding the launcher process from VBA: PASS.
- `.ppam` add-in can be created and loaded manually: PASS.
- `Zotero` Ribbon tab appears in a normal `.pptx`: PASS.
- **Picker starten** starts the existing picker from a normal `.pptx`: PASS.
- APA slide citation after Ribbon/PPAM start: PASS.
- Bibliography target and update after Ribbon/PPAM start: PASS.
- APA notes citation after Ribbon/PPAM start: PASS.
- Automatic bibliography update after notes citation insert: PASS.
- PowerPoint restart persistence: PASS.
- Log inspection: PASS.

**Known limitations**
- The `.ppam` file is created manually and is not committed to the repository.
- The add-in was tested as an unsigned local add-in.
- Users may need to adjust PowerPoint Trust Center settings.
- The local launcher path is still edited manually in VBA.
- Only one Ribbon button is included in the initial scope.
- Direct Ribbon actions for document update, bibliography rewrite, and bibliography target setup remain follow-up work.

**Overall result**
- `v0.1.0-alpha.21 – PowerPoint Ribbon picker button`: release-ready.

### v0.1.0-alpha.22 – PowerPoint Ribbon actions and workflow unification

**Scope**
- Added direct PowerPoint Ribbon/CLI actions for:
  - `set-bibliography-target`
  - `rewrite-bibliography`
  - `update-document`
- Extended `scripts/start_picker.cmd` so action arguments are forwarded to Python.
- Extended `powerpoint/LaunchZoteroPicker.bas` with callbacks for all Ribbon actions.
- Updated `powerpoint/customUI14.xml` to expose four Ribbon buttons:
  - **Zitation einfügen**
  - **Dokument aktualisieren**
  - **Bibliographie neu schreiben**
  - **Bibliographie-Ziel festlegen**
- Grouped bibliography-related Ribbon actions under **Bibliographie**.
- Kept the Picker-App workflow and Ribbon/CLI workflow on shared Python workflow functions.
- Removed the experimental bibliography anchor reference helper path.
- Added `.xml` line-ending handling to `.gitattributes`.

**Architecture**

```text
PowerPoint Ribbon button
  -> VBA callback
     -> scripts/start_picker.cmd --action ...
        -> zotero_picker_ppt.py
           -> shared workflow implementation
```

The Ribbon and CLI actions do not implement separate citation or bibliography
logic. They call the same shared Python workflow functions used by the Picker-App
buttons.

For CLI/Ribbon action runs, Python keeps a hidden Tk event loop alive and runs
the selected workflow in a worker thread. This matches the Picker-App execution
model more closely than a synchronous command-line call and avoids PowerPoint COM
instability observed during bibliography target setup and writing.

**Behavior**
- Zitation einfügen opens the Picker-App and supports citation insertion and citation style selection.
- Bibliographie-Ziel festlegen stores the selected text box as bibliography target and updates the bibliography when citation keys exist.
- Bibliographie neu schreiben rebuilds the bibliography from the current citation state.
- Dokument aktualisieren resyncs citation state and updates the bibliography when a target exists.
- German user-facing labels and success/error dialogs are preserved.
- Maintainer-facing logs remain English.

**Fixes**
- Fixed unreliable CLI/Ribbon bibliography target setup where the target was detected and saved, but later rejected as unusable.
- Removed the redundant `_make_shape_ref()`, `_resolve_shape_refs()`, and `_get_shape_by_id()` helper path.
- Simplified bibliography writing to rely on the existing anchor resolver.
- Changed CLI action execution to use a hidden Tk root plus worker thread, aligning it with Picker-App button execution.
- Cleaned accidental PowerShell here-string remnants from generated launcher/Ribbon files.
- Made Ribbon XML robust against encoding problems by using XML character references for German umlauts.

**Manual retest result**
- Static checks: PASS.
- `python -m py_compile zotero_picker_ppt.py`: PASS.
- `python zotero_picker_ppt.py --help`: PASS.
- `git diff --check`: PASS.
- CLI `--action set-bibliography-target`: PASS.
- CLI `--action rewrite-bibliography`: PASS.
- CLI `--action update-document`: PASS.
- `scripts/start_picker.cmd --action set-bibliography-target`: PASS.
- `scripts/start_picker.cmd --action rewrite-bibliography`: PASS.
- `scripts/start_picker.cmd --action update-document`: PASS.
- `.pptm` Ribbon **Zitation einfügen**: PASS.
- `.pptm` Ribbon **Bibliographie-Ziel festlegen**: PASS.
- `.pptm` Ribbon **Bibliographie neu schreiben**: PASS.
- `.pptm` Ribbon **Dokument aktualisieren**: PASS.
- `.ppam` Ribbon smoke test: PASS.
- Log inspection: PASS.

**Known limitations**
- The `.ppam` file is still created manually and is not committed to the repository.
- The add-in was tested as an unsigned local add-in.
- Users may need to adjust PowerPoint Trust Center settings.
- The local `PROJECT_ROOT` path is still edited manually in VBA.
- No installer or EXE package is included.
- No locator/page support.
- No full CSL/style-engine refactor.

**Overall result**
- `v0.1.0-alpha.22 – PowerPoint Ribbon actions and workflow unification`: completed, tested, tagged, and released.

### v0.1.0-alpha.23 – PowerPoint launcher UX polish

**Scope**
- Improved the PowerPoint Ribbon launcher user experience.
- Hid the transient command window when starting actions from PowerPoint.
- Reused an already open Picker window instead of starting a second Picker instance.
- Rebuilt and retested the `.ppam` add-in with the updated VBA launcher.

**Behavior**
- Clicking **Zitation einfügen** opens the Picker if it is not already running.
- If the Picker is already open but hidden behind PowerPoint, clicking **Zitation einfügen** brings the existing Picker window to the foreground.
- Ribbon action buttons no longer show a briefly flashing command window.
- Existing action behavior remains unchanged:
  - **Dokument aktualisieren**
  - **Bibliographie neu schreiben**
  - **Bibliographie-Ziel festlegen**

**Fixes**
- Changed the VBA launcher call from visible command execution to hidden command execution.
- Added Picker-window reuse via `WScript.Shell.AppActivate("Zotero Picker")`.
- Avoided duplicate Picker windows when users click **Zitation einfügen** repeatedly.

**Manual retest result**
- `.pptm` launcher button without visible console window: PASS.
- `.ppam` rebuilt from the updated `.pptm`: PASS.
- `.ppam` **Zitation einfügen** with existing hidden Picker window: PASS.
- No second Picker window opened during repeated **Zitation einfügen** clicks: PASS.
- Citation insertion with the reused Picker window: PASS.
- Bibliography auto-update after inserted citations: PASS.
- Log inspection: PASS.

**Known limitations**
- Picker reuse depends on the Picker window title matching `Zotero Picker`.
- The `.ppam` file is still created manually and is not committed to the repository.
- The add-in remains an unsigned local add-in.
- Users may still need to adjust PowerPoint Trust Center settings.
- The local `PROJECT_ROOT` path is still edited manually in VBA.

**Overall result**
- `v0.1.0-alpha.23 – PowerPoint launcher UX polish`: completed, tested, and release-ready.