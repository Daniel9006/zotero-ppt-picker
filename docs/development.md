# zotero-ppt-picker – Developer Notes

This document describes the internal architecture and design decisions for developers.

---

## Project structure

```
zotero-ppt-picker/
├── zotero_picker_ppt.py        # main application entry point
├── config/
│   ├── __init__.py
│   └── zotero_config.py        # configuration and credential handling
├── docs/
│   ├── development.md          # developer and architecture documentation
│   ├── debugging.md            # debugging notes and known runtime issues
│   ├── powerpoint_launcher.md  # PowerPoint launcher and Ribbon documentation
│   ├── testing.md              # manual alpha retest checklist
│   └── mac_linux.md            # macOS / Linux notes and limitations
├── powerpoint/
│   ├── LaunchZoteroPicker.bas  # PowerPoint VBA callbacks
│   └── customUI14.xml          # PowerPoint Ribbon XML
├── scripts/
│   └── start_picker.cmd        # Windows command launcher for PowerPoint/VBA/Ribbon actions
├── test_zotero_config.py       # standalone test for config and credential dialog
├── requirements.txt            # Python dependencies
├── README.md                   # user documentation (installation, configuration, usage)
├── TEAM.md                     # team and collaboration notes
├── VERSIONING.md               # versioning rules and release ladder
├── CODING_STANDARDS.md         # coding rules and quality expectations
├── .env                        # optional local overrides (never commit secrets)
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

## PowerPoint launcher and Ribbon architecture

The project includes a Windows/PowerPoint launcher path for starting the existing
Picker UI and for running selected PowerPoint actions from Ribbon buttons.

```text
PowerPoint Ribbon button
  -> VBA callback in powerpoint/LaunchZoteroPicker.bas
     -> scripts/start_picker.cmd [optional --action ...]
        -> zotero_picker_ppt.py
           -> shared workflow implementation
```

The launcher is intentionally separated from citation and bibliography logic.
It does not introduce a new citation engine, bibliography model, Zotero
configuration mechanism, or style engine.

The current launcher supports two modes:

```text
scripts/start_picker.cmd
```

Starts the full Picker UI.

```text
scripts/start_picker.cmd --action set-bibliography-target
scripts/start_picker.cmd --action rewrite-bibliography
scripts/start_picker.cmd --action update-document
```

Runs a specific PowerPoint workflow without opening the Picker UI.

Responsibilities:

- `powerpoint/customUI14.xml` defines the custom PowerPoint Ribbon tab.
- `powerpoint/LaunchZoteroPicker.bas` contains the VBA callbacks for the Ribbon
  buttons.
- `scripts/start_picker.cmd` resolves the repository root relative to its own
  location and forwards optional action arguments to Python.
- `zotero_picker_ppt.py` remains the only citation and bibliography
  implementation.
- CLI/Ribbon actions call the same shared Python workflow functions that are used
  by the Picker-App buttons.

The Ribbon currently exposes these user-facing actions:

```text
Zitationen
- Zitation einfuegen

Dokument
- Dokument aktualisieren

Bibliographie
- Bibliographie neu schreiben
- Bibliographie-Ziel festlegen
```

For CLI/Ribbon action runs, `zotero_picker_ppt.py` keeps a hidden Tk event loop
alive and runs the workflow in a worker thread. This mirrors the Picker-App
execution model and avoids PowerPoint COM instability observed with synchronous
headless action execution.

Out of scope:

- new citation engine
- separate bibliography implementation
- Zotero credential-flow changes
- EXE packaging
- installer
- signed or centrally deployed PPAM rollout
- locator/page support

Detailed setup and troubleshooting are documented in
`docs/powerpoint_launcher.md`.

---

## User-facing maintenance workflows

As of `v0.1.0-alpha.19`, the main user-facing maintenance workflow remains **Dokument aktualisieren**.

This workflow is the primary path after users edit or delete citations in PowerPoint. Internally, it calls the same central cleanup/resync logic that maintains the relationship between visible citation text, stored citation metadata, and bibliography contents.

Starting with `v0.1.0-alpha.19`, this document-wide resync includes citations stored in normal slide shapes and in PowerPoint NotesPage shapes.

### Dokument aktualisieren

The **Dokument aktualisieren** workflow:

- resynchronizes visible citations with stored citation state
- scans normal slide shapes and NotesPage shapes
- removes deleted slide or notes citations from the bibliography
- clears the bibliography when no citations remain
- tolerates missing bibliography targets
- repairs APA/Harvard suffix disambiguation
- runs IEEE renumbering across slide and notes citations
- normalizes MLA duplicate visible labels
- performs only the base citation-state resync for Chicago Author-Date

This is the preferred user-facing workflow for document maintenance.

### Bibliographie neu schreiben

The **Bibliographie neu schreiben** workflow is secondary and bibliography-only.

It rewrites the bibliography from the current stored citation state when a bibliography target exists. It is not the primary repair workflow and should not be presented as the main way to fix citation-state inconsistencies.

It does not primarily:

- change visible citations
- repair APA/Harvard suffixes
- renumber IEEE citations

It uses the current stored citation state from normal slide shapes and NotesPage shapes, but the bibliography target itself remains a normal slide shape.

---

## Citation state model

Citations are persisted in PowerPoint shape tags using `ZP_CITES`.

This internal citation state is required because visible citation text alone is not sufficient for deterministic cleanup, bibliography rebuilds, and style-specific renumbering.

### Citation scan scope

As of `v0.1.0-alpha.19`, document-wide citation scans include:

1. normal slide shapes
2. NotesPage shapes for the same slide

The intentional document order is:

```text
Slide 1 → Notes 1 → Slide 2 → Notes 2 → …
```

This order is relevant for numeric styles such as IEEE because numbering is derived from the document-wide citation order.

The shared citation-shape iteration is used by document-wide paths such as:

- `collect_all_cites_by_key()`
- `collect_all_cite_texts()`
- `normalize_sig_group(...)`
- `renormalize_all_sig_groups()`
- `build_ieee_numbering_from_document()`
- `resync_bibliography_keys_from_document(...)`
- `renumber_ieee_and_update(...)`

Bibliography anchor resolution intentionally remains limited to normal slide shapes. The anchor is not searched for or set in NotesPage shapes.

Stored citation records contain at least:
- `key`: Zotero item key
- `cite`: currently visible citation text
- `style`: citation style for newly written records
- optional style-specific metadata such as `sig`, `mla_label`, or `mla_qualifier`

As of `v0.1.0-alpha.25`, one presentation is treated as a single-style document.
The document style is inferred from active `ZP_CITES` records on normal slides and
PowerPoint notes. Explicit `style` metadata is preferred. Legacy records without
`style` are interpreted conservatively using IEEE citation syntax, legacy MLA
heuristics, and the previous document `state["style"]` for plausible author-year
records. Unknown records are not silently ignored.

When the user changes the selected style in a presentation with existing
citations, the application does not simply save a new `state["style"]`. Instead,
it asks for a controlled full-document conversion. A successful conversion updates
visible citation text and stored `ZP_CITES` records together, runs the relevant
style-specific normalization, rebuilds `bib_keys`, and only then stores the new
document style. If a bibliography target exists, the bibliography is rewritten in
the target style.

For MLA records created with `v0.1.0-alpha.24` or later, additional metadata may be stored:

- `style`: `mla`
- `mla_label`: the base visible MLA label, for example an author or corporate author
- `mla_qualifier`: a short-title/title qualifier used when multiple different Zotero items would otherwise share the same visible MLA label

MLA duplicate-label normalization updates the visible citation text and the stored `ZP_CITES` record together. It is intentionally metadata-based and does not call the Zotero Web API during document update.

Important rules:
- Visible citation text and stored cite metadata must be updated together.
- New mixed citation styles must be blocked before visible text or tags are written.
- `state["style"]` must not be changed before an explicit style conversion has succeeded.
- Cleanup must derive bibliography keys from stored citation metadata.
- Numeric styles such as IEEE must not rely only on visible placeholder scans.
- IEEE numbering is built from persisted cite records sorted by visible document order.
- Bibliography labels returned by external formatters may need normalization before applying document-level numbering.

---

### Notes insert fallback

PowerPoint does not always expose a reliable shape object when the text cursor is inside the notes pane.

For notes insertion, the application therefore uses a fallback path when the normal ShapeRange/parent lookup does not identify the target shape:

1. insert a temporary marker at the current cursor position
2. scan the current slide and its NotesPage shapes
3. find the text shape containing the marker
4. remove the marker
5. store the citation metadata (`ZP_CITES`) on the detected notes shape

This keeps the existing citation-state model unchanged while allowing notes citations to participate in document-wide update and bibliography workflows.

---

## Citation style validation status

As of `v0.1.0-alpha.24`, the base citation style matrix has been retested manually with slide and notes citations.

The scope of this validation was limited to the current alpha base functionality:

- citation insertion on normal slides
- citation insertion in PowerPoint notes
- bibliography target setup on normal slide shapes
- bibliography-only rebuild via **Bibliographie neu schreiben**
- document-level resync via **Dokument aktualisieren**
- cleanup after deleting slide or notes citations
- full citation deletion and bibliography clearing
- notes-only citation scenarios
- missing bibliography anchor handling
- persistence after save, close, and reopen
- rough stylistic plausibility

Locator and detail references such as pages, chapters, clauses, figures, and tables were not part of this validation.

| Style | Status | Notes |
| --- | --- | --- |
| APA | Passed | Insert on slide and in notes, notes citation in bibliography, notes deletion cleanup, persistence, document update, and bibliography rewrite passed. |
| IEEE | Passed | Numbering across `Slide 1 → Notes 1 → Slide 2` produced `[1]`, `[2]`, `[3]`; deleting a notes citation renumbered remaining citations and cleaned the bibliography correctly. |
| Chicago Author-Date | Passed in alpha scope | Insert on slide and in notes, notes citation deletion, bibliography cleanup, and unchanged base behavior passed. Duplicate author/year disambiguation remains a follow-up. |
| Harvard | Passed | Insert on slide and in notes, notes citation deletion, and bibliography cleanup passed. |
| MLA | Passed in alpha scope | Insert on slide and in notes, notes citation deletion, bibliography cleanup, duplicate visible-label handling, deletion rollback, and no-collision behavior passed. MLA notes citations did not regress to author-date rendering. Locator/page support and full CSL-style validation remain future work. |

Open follow-up topics:

1. Chicago Author-Date duplicate author/year disambiguation remains future work.
2. Locator/detail reference support must be designed as a separate feature block.
3. Deeper CSL/style-engine validation remains future work.
4. Signed/corporate PPAM rollout, installer packaging, and broader launcher deployment hardening remain future work.
5. A separate notes bibliography mode is not planned for this alpha scope.

---

## Testing

Manual PowerPoint alpha retests are documented in `docs/testing.md`.
Use that checklist before tagging alpha releases and after citation,
bibliography workflow, or launcher changes.

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

Launcher-specific retests are documented in the PowerPoint launcher section of
`docs/testing.md`.

---

## Contribution guidelines

- Do not commit credentials or `.env` files
- Keep configuration logic inside `zotero_config.py`
- Avoid platform-specific logic outside dedicated modules
- Keep PowerPoint launcher code separate from citation and bibliography logic
- Do not implement duplicate citation or bibliography workflows in VBA or command
  launcher code
- Route CLI and Ribbon actions through the shared Python workflow functions

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
