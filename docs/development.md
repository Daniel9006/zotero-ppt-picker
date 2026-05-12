# zotero-ppt-picker – Developer Notes

This document describes the internal architecture and design decisions for developers.

---

## Project structure

```
zotero-ppt-picker/
├── zotero_picker_ppt.py      # main application entry point
├── config/
│   ├── __init__.py
│   └── zotero_config.py      # configuration and credential handling
├── docs/
│   ├── development.md        # developer and architecture documentation
│   ├── debugging.md          # debugging notes and known runtime issues
│   └── mac_linux.md          # macOS / Linux notes and limitations
├── test_zotero_config.py     # standalone test for config and credential dialog
├── requirements.txt          # Python dependencies
├── README.md                 # user documentation (installation, configuration, usage)
├── TEAM.md                   # team and collaboration notes
├── VERSIONING.md             # versioning rules and release ladder
├── CODING_STANDARDS.md       # coding rules and quality expectations
├── .env                      # optional local overrides (never commit secrets)
└── .gitignore
```

### Notes
- `test_zotero_config.py` is intended for **manual testing and debugging** of the
  configuration and credential dialog without running the full application.
- `.env` is optional and should only be used locally.
- Any file containing real API keys must **never** be committed.

---

## Configuration architecture

Credential resolution:

1. Load local user config file, if present  
   - Platform-specific user config location
   - Stored as JSON
2. Apply environment variable overrides  
   - `ZOTERO_API_KEY`
   - `ZOTERO_LIBRARY_ID`
   - `ZOTERO_LIBRARY_TYPE`
3. Open interactive GUI prompt (Tkinter) if required values are still missing or invalid

This ensures:
- no secrets in the repository
- per-user configuration
- optional environment-based overrides
- deterministic behavior

⚠️ Files containing real Zotero API keys must never be committed.  
Use the interactive configuration dialog or local `.env` files instead.

---

## zotero_config.py

Key responsibilities:

- Loading credentials from multiple sources
- Validating user input
- Persisting configuration safely
- Providing a single entry point:
  ```python
  load_zotero_config(allow_prompt=True, parent=...)
  ```

Errors are raised as `ConfigError` and must be handled by the caller.

---

## GUI design notes

- Tkinter dialogs are implemented defensively for Windows focus handling
- Modal dialogs use `grab_set()` only after becoming visible
- A hidden but real root window is used to avoid Windows/Tk issues

---

## Citation state model

Citations are persisted in PowerPoint shape tags using `ZP_CITES`.

This internal citation state is required because visible citation text alone is not sufficient for deterministic cleanup, bibliography rebuilds, and style-specific renumbering.

Stored citation records contain at least:
- `key`: Zotero item key
- `cite`: currently visible citation text
- optional style-specific metadata such as `sig` or `style`

Important rules:
- Visible citation text and stored cite metadata must be updated together.
- Cleanup must derive bibliography keys from stored citation metadata.
- Numeric styles such as IEEE must not rely only on visible placeholder scans.
- IEEE numbering is built from persisted cite records sorted by visible document order.
- Bibliography labels returned by external formatters may need normalization before applying document-level numbering.

---

## Citation style validation status

As of `v0.1.0-alpha.17`, the base citation style matrix has been retested manually.

The scope of this validation was limited to the current alpha base functionality:
- citation insertion
- automatic bibliography update
- manual bibliography update
- cleanup after partial citation deletion
- full citation deletion and bibliography clearing
- late bibliography anchor setup
- missing bibliography anchor handling
- rough stylistic plausibility

Locator and detail references such as pages, chapters, clauses, figures, and tables were not part of this validation.

| Style | Status | Notes |
| --- | --- | --- |
| APA | Passed | Regression test passed. No regression found after IEEE and bibliography changes. |
| IEEE | Passed | Technically alpha-stable and broadly IEEE-plausible. Locator-like manual additions such as `[1, Clause 1]` are not cleanup-stable yet. |
| Chicago Author-Date | Passed | Technically stable and broadly plausible for the base alpha scope. |
| Harvard | Passed | Technically stable and broadly plausible for the base alpha scope. |
| MLA | Passed in alpha scope | Technically stable and no longer rendered as author-date for new inserts. Uses minimal MLA-plausible parenthetical labels. Locator/page support and full CSL-style validation remain future work. |

Open follow-up topics:
1. MLA disambiguation for identical visible labels remains future work.
2. Locator/detail reference support must be designed as a separate feature block.
3. Deeper CSL/style-engine validation remains future work.

---

## Testing

Recommended local tests:

```powershell
# Windows: force credential dialog
Remove-Item "$env:APPDATA\ZoteroPowerPoint\config.json" -ErrorAction SilentlyContinue

python test_zotero_config.py
```

```bash
# macOS / Linux: force credential dialog
rm -f ~/.config/ZoteroPowerPoint/config.json

python test_zotero_config.py
```

---

## Contribution guidelines

- Do not commit credentials or `.env` files
- Keep configuration logic inside `zotero_config.py`
- Avoid platform-specific logic outside dedicated modules

## Refactor roadmap (phase-aligned)

This section lists the planned architecture and cleanup steps, aligned
with the versioning phases defined in `VERSIONING.md`.

### Phase 1 – Architecture Refactor
Goal:
- Reduce duplicated logic in citation styles
- Centralize shared formatting and rendering
- Improve configuration resolution via constructor/factory

Planned tasks:
- Create a shared citation style engine
- Refactor `zotero_config.py` into:
  - `from_env()`
  - `from_user_config()`
  - `resolve()` with clear priority and validation
- Add targeted tests for config and style engine

Completion criteria:
- Style engine in place and used by APA
- No duplicated logic per citation style
- Config resolution follows a deterministic pipeline

### Phase 2 – Citation styles

IEEE has been introduced as an alpha-level technical implementation before the full style-engine refactor.

Current status:
- IEEE uses persistent citation records in PowerPoint shape tags.
- IEEE renumbering is based on visible document order.
- Bibliography rebuild, cleanup, and late bibliography anchor setup are supported.

Remaining architecture goal:
- Move APA / IEEE / future styles into a shared style engine.
- Reduce style-specific branching in `zotero_picker_ppt.py`.

See `VERSIONING.md` for details on versioning cycles.
