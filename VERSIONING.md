# Versioning & Release Ladder (Zotero ↔ PowerPoint)

This repository uses a pragmatic, small-team-friendly versioning approach:

- clear milestones
- safe rollbacks
- minimal process overhead
- no CI assumptions

**Tags are the source of truth.**

Current public baseline: `v0.1.0-alpha.15`

Current development focus:
- technical stabilization of citation and bibliography mechanics
- IEEE alpha hardening completed in `v0.1.0-alpha.15`
- persistent citation state and document resync reliability

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

Notes:
- Migration to English-only comments and documentation has started

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
1. English-only comments and docstrings
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

## Bugfix documentation

This section documents verified fixes that affected runtime behavior.

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