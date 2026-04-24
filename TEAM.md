# Team Workflow – Zotero ↔ PowerPoint

This project uses a simple and safe Git workflow for small teams of one or two people.

Goals:
- keep the main branch stable
- maintain a traceable history
- allow returning to known working states at any time

---

## Ground rules

1. **`main` should always be runnable**
   - no broken intermediate states
   - no `WIP` commits on `main`

2. **Small changes → commit directly or use a short-lived branch**

3. **Larger changes → use a dedicated feature branch**
   - examples: `feature/config-arch`, `fix/com-threading`

4. **No mandatory rebase workflow**
   - use `git merge`
   - focus on clarity, not on a perfect history

---

## Refactor governance

Major refactors such as architecture changes, core logic changes, or shared engines follow these rules:

- Refactors should be proposed and discussed before major effort.
- Use a feature or refactor branch, for example `feature/style-engine` or `refactor/config-factory`.
- Refactors should not be mixed with urgent bug fixes on the same branch.
- Do not merge an incomplete refactor into `main`; ensure:
  - tests pass
  - exceptions and error flows are deterministic
  - stable flows do not regress

**Stabilization vs. refactor**
- Bugfix stabilization → `fix/*` branches, merged into `main`
- Architectural refactor → `feature/*` or `refactor/*` branches
- Refactors belong to phase gates in `VERSIONING.md`

---

## Commit rules

- Commit small, logically related changes.
- Use meaningful commit messages, for example:
  - `feat: add config loader`
  - `fix: COM init in worker thread`
  - `refactor: split zotero access layer`

Recommended prefixes:
- `feat:` new feature
- `fix:` bug fix
- `refactor:` restructuring without intended behavior change
- `docs:` documentation or comments

---

## Branching

- `main` → stable main branch
- `feature/*` → new features or larger changes
- `fix/*` → bug fixes
- `experiment/*` → tests or playground branches that may be deleted

Before merging into `main`:

1. Test the branch locally.
2. Merge the latest `main` into the branch so the branch is not outdated.
3. For bug fixes affecting citation or bibliography behavior, include a short manual test note in the commit message or release notes.

---

## Stable versions and tags

When a state is verified to work:

- Zotero access works
- PowerPoint insertion works
- no known crash exists in the tested scope

→ create a tag, for example:

- `v0.1.0`
- `v0.2.0`
- `v0.2.1`

Tags are fixed markers for known working versions.

---

## Reusing old versions

To test an old tagged version:

```bash
git switch --detach v0.2.0
```
To return to the main branch:
```bash
git switch main
```