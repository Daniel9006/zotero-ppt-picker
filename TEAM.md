# Team Workflow – Zotero ↔ PowerPoint

Dieses Projekt nutzt einen einfachen, sicheren Git-Workflow
für kleine Teams (1–2 Personen).

Ziel:
- stabile Hauptversion
- nachvollziehbare Historie
- jederzeitige Rückkehr zu funktionierenden Ständen

---

## Grundregeln

1. **`main` soll immer lauffähig sein**
   - keine kaputten Zwischenstände
   - kein „WIP“ auf `main`

2. **Kleine Änderungen → direkt oder kurzer Branch**
3. **Größere Umbauten → eigener Feature-Branch**
   - z.B. `feature/config-arch`, `fix/com-threading`

4. **Kein Rebase-Zwang**
   - wir nutzen `git merge`
   - Fokus auf Verständlichkeit, nicht auf perfekte Historie

---

## Commit-Regeln

- Klein & logisch zusammenhängend committen
- Aussagekräftige Nachrichten, z.B.:
  - `feat: add config loader`
  - `fix: COM init in worker thread`
  - `refactor: split zotero access layer`

Empfohlene Präfixe:
- `feat:` neues Feature
- `fix:` Bugfix
- `refactor:` Umstrukturierung
- `docs:` Doku / Kommentare

---

## Branching

- `main` → stabiler Hauptbranch
- `feature/*` → neue Features / Umbauten
- `fix/*` → Bugfixes
- `experiment/*` → Tests / Spielwiese (dürfen gelöscht werden)

Vor Merge in `main`:
1. Branch lokal getestet
2. `main` wurde gemerged (kein veralteter Stand)

---

## Stabile Versionen (Tags)

Wenn ein Stand **nachweislich funktioniert**:
- Zotero-Zugriff OK
- PowerPoint-Insert OK
- kein bekannter Crash

→ **Tag setzen**, z.B.:

- `v0.1.0`
- `v0.2.0`
- `v0.2.1`

Tags sind **fix** und markieren funktionierende Versionen.

---

## Alte Versionen wiederverwenden

- Zum Testen:
  ```bash
  git switch --detach v0.2.0
