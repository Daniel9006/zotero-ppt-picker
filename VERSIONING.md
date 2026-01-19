# Versioning & Release Ladder (Zotero ↔ PowerPoint)

This repository uses a pragmatic, small-team-friendly versioning approach:

- clear milestones
- safe rollbacks
- minimal process overhead
- no CI assumptions

**Tags are the source of truth.**

Current public baseline: `v0.1.0-alpha.2`

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
- The style engine must exist before adding new citation styles.

---

### Phase 2 — New citation styles (each is a MINOR bump)

Each style is treated as a major milestone.

- `v0.2.0` — IEEE
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