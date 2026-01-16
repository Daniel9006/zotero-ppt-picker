# Coding Standards (Zotero ↔ PowerPoint)

These standards keep the project maintainable for a small team.
They are enforced starting at `v0.1.0-beta.1`.

---

## 1) Language & style

- **All comments and docstrings are in English**
- Keep comments short and valuable:
  - Prefer explaining *why* something is done
  - Avoid restating obvious code (“what”)

---

## 2) “Comment every variable” — interpretation

We do **not** add a comment for every variable by default.
Instead we require:

- Type hints for function signatures and important variables
- Docstrings for:
  - modules (where helpful)
  - classes
  - public functions / key internal functions
- Inline comments only when the intent is non-obvious

If a variable needs explanation, prefer:
- a better name, or
- a short inline comment where it’s defined

---

## 3) Exceptions & error handling

### Goals
- Exceptions are **specific**
- Errors are handled **deterministically**
- Users see actionable messages, not stack traces
- Logs contain enough context to debug

### Rules
- Avoid broad `except Exception:` in core logic.
  - Allowed at top-level boundaries (CLI/UI entry) to convert errors into user messages.
- Prefer domain-specific exceptions:
  - `ConfigError`
  - `ZoteroApiError`
  - `PowerPointComError`
  - `CitationStyleError`
- Always attach context:
  - which operation failed
  - relevant IDs (library id, item key) where safe
  - underlying exception as cause (`raise X(...) from e`)

### Example pattern
- Low-level code raises specific exceptions
- UI/controller catches those specific exceptions and shows a friendly message
- Unknown exceptions are logged and shown as a generic failure

---

## 4) Configuration architecture (`zotero_config.py`)

Configuration should be resolved via a clear constructor/factory approach:

Recommended shape:
- `ZoteroConfig.from_env()`
- `ZoteroConfig.from_user_config(path=...)`
- `ZoteroConfig.resolve(...)` (priority rules: env overrides optional, local config primary)
- Validation happens during construction/resolution

Rules:
- No scattered config parsing across the UI
- Invalid config fails early with `ConfigError` (and clear message)

---

## 5) Citation styles: avoid duplicated logic

### Goal
APA / IEEE / Chicago / Harvard share common mechanics.
We avoid copy/paste by introducing a shared layer.

Rules:
- Common formatting primitives live in shared helpers / a style engine:
  - author name formatting
  - date formatting
  - title casing rules
  - punctuation joining rules
  - DOI/URL rendering
  - in-text citation assembly (common steps)
- Each style should ideally be:
  - a set of rules/config
  - plus small style-specific overrides

Minimum requirement before adding the next style:
- style engine exists and is used by APA

---

## 6) Tests

- Add/adjust tests for any change that affects:
  - config resolution/validation
  - citation formatting rules
  - COM/PowerPoint boundary behavior (as far as testable)
- Prefer small unit tests for formatting rules and config edge cases

---

## 7) Commit hygiene (lightweight)

- Keep commits logically scoped
- Suggested prefixes:
  - `feat:`, `fix:`, `refactor:`, `docs:`, `test:`, `chore:`

---

## 8) Practical enforcement

Starting at `v0.1.0-beta.1`:
- PR/review (even informal) checks for:
  - English comments/docstrings
  - exception specificity
  - no new citation-style duplication
  - config changes follow constructor/factory pattern
