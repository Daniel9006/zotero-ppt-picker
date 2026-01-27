import os
import re
import sys
import json
import html
import uuid
import threading
import tkinter as tk
from tkinter import ttk, messagebox

import win32com.client as win32
from pyzotero import zotero
import requests

from typing import Optional
from config.zotero_config import load_zotero_config, ConfigError, ZoteroConfig

import contextlib
import logging
import traceback

import pythoncom  # pywin32

# ===================== Konfiguration =====================
MAX_RESULTS   = 50
HTTP_TIMEOUT  = 15
DOCPROP_NAME  = "ZPCiteState"  # CustomDocumentProperty (JSON)
DEFAULT_STYLE = "apa"
STYLE_CHOICES = [
    ("APA", "apa"),
    ("IEEE", "ieee"),
    ("Chicago Author-Date", "chicago-author-date"),
    ("Harvard", "harvard1"),
    ("MLA", "mla"),
]
PREF_FONT_SIZE = 14
MIN_FONT_SIZE  = 10

# Bibliographie-Anker (robust)
ALT_BIB_PREFIX = "ZP_BIB_GUID="

# Tags für stabile Anker-Erkennung
TAG_BIB_GUID_KEY = "ZP_BIB_GUID"

# JSON-Liste pro Shape: [{"key": "...", "cite": "(...)"}, ...]
CITE_TAG = "ZP_CITES"
# =========================================================

# ===================== COM Stabilisierung (Option A+) =====================
COM_LOCK = threading.RLock()

LOG = logging.getLogger("zotero_ppt")

@contextlib.contextmanager
def com_context(action: str = ""):
    """
    Option A+:
    - CoInitialize/CoUninitialize pro Worker-Thread
    - serialisiert alle COM-Zugriffe via COM_LOCK (RLock erlaubt nested calls)
    """
    pythoncom.CoInitialize()
    try:
        with COM_LOCK:
            if action:
                LOG.debug("COM enter: %s", action)
            yield
    finally:
        try:
            if action:
                LOG.debug("COM exit: %s", action)
        finally:
            pythoncom.CoUninitialize()
# ========================================================================

# ======== Zero-width Marker für Nicht-IEEE-Zitate ========
ZWM_START = "\u2062"
ZWM_END   = "\u2063"
ZWM_ALPH = ["\u200b", "\u200c", "\u200d", "\u2060"]
ZWM_MAP = {ch: i for i, ch in enumerate(ZWM_ALPH)}

def _zwm_encode_key(key: str) -> str:
    data = key.encode("utf-8")
    digits = []
    for b in data:
        digits.extend([(b >> 6) & 3, (b >> 4) & 3, (b >> 2) & 3, b & 3])
    return ZWM_START + "".join(ZWM_ALPH[d] for d in digits) + ZWM_END

def _zwm_decode_keys_from_text(text: str):
    keys, pos = [], 0
    while True:
        i = text.find(ZWM_START, pos)
        if i == -1: break
        j = text.find(ZWM_END, i + 1)
        if j == -1: break
        body = text[i+1:j]
        if body and len(body) % 4 == 0:
            out = bytearray()
            ok = True
            for k in range(0, len(body), 4):
                d = [ZWM_MAP.get(body[k+x], None) for x in range(4)]
                if None in d:
                    ok = False
                    break
                out.append((d[0]<<6) | (d[1]<<4) | (d[2]<<2) | d[3])
            if ok:
                try:
                    keys.append(out.decode("utf-8", errors="ignore"))
                except Exception:
                    pass
        pos = j + 1
    return keys
# =========================================================
def _debug(msg):
    try:
        print(f"[ZP] {msg}")
    except Exception:
        pass

_CFG: Optional[ZoteroConfig] = None

def reset_cfg_cache():
    """Erlaubt später ein 'Config neu laden' ohne Prozess-Neustart."""
    global _CFG
    _CFG = None

def get_cfg(*, allow_prompt: bool, parent: Optional[tk.Misc] = None) -> ZoteroConfig:
    """
    Single source of truth for Zotero credentials.
    - Loads persisted config.json first
    - ENV overrides
    - Prompts only if allow_prompt=True and missing/invalid
    """
    global _CFG
    if _CFG is not None:
        return _CFG

    try:
        _CFG = load_zotero_config(allow_prompt=allow_prompt, parent=parent)
        _debug(f"Zotero config loaded: type={_CFG.library_type}, id={_CFG.library_id}")
        return _CFG
    except ConfigError as e:
        # raise RuntimeError so existing messagebox handling stays consistent
        raise RuntimeError(str(e)) from e


# ===================== PowerPoint Helpers =================

def _get_presentation():
    with com_context("_activate_powerpoint"):
        app = win32.gencache.EnsureDispatch("PowerPoint.Application")
        pres = app.ActivePresentation
        if not pres:
            raise RuntimeError("Keine aktive Präsentation.")
        return pres
    
def _activate_powerpoint():
    with com_context("_activate_powerpoint"):
        try:
            app = win32.gencache.EnsureDispatch("PowerPoint.Application")
            app.Activate()
            if app.ActiveWindow is not None:
                app.ActiveWindow.Activate()
        except Exception:
            pass

def _get_current_slide_and_shape():
    with com_context("_activate_powerpoint"):
        app = win32.gencache.EnsureDispatch("PowerPoint.Application")
        win = app.ActiveWindow
        if not win:
            return None, None

        sel = win.Selection
        slide = None
        shape = None

        try:
            slide = sel.SlideRange(1)
        except Exception:
            try:
                slide = win.View.Slide
            except Exception:
                slide = None

        ppSelectionSlides = 1
        ppSelectionShapes = 2
        ppSelectionText   = 3
        try:
            sel_type = sel.Type
        except Exception:
            sel_type = None

        if sel_type == ppSelectionText:
            try:
                tr = sel.TextRange
                if tr is not None:
                    shape = tr.Parent
                    if shape is not None and getattr(shape, "HasTextFrame", False):
                        return slide, shape
            except Exception:
                pass
            try:
                sr = sel.ShapeRange
                if sr is not None and sr.Count >= 1:
                    shape = sr.Item(1)
                    if shape is not None and getattr(shape, "HasTextFrame", False):
                        return slide, shape
            except Exception:
                pass

        if sel_type == ppSelectionShapes:
            try:
                sr = sel.ShapeRange
                if sr is not None and sr.Count >= 1:
                    shape = sr.Item(1)
                    if shape is not None and getattr(shape, "HasTextFrame", False):
                        return slide, shape
            except Exception:
                pass

        if sel_type == ppSelectionSlides or shape is None:
            try:
                if slide is not None:
                    best, best_area = None, -1
                    for shp in slide.Shapes:
                        try:
                            if getattr(shp, "HasTextFrame", False):
                                _ = shp.TextFrame  # nur um sicherzugehen, dass der Zugriff nicht crasht
                                area = float(shp.Width) * float(shp.Height)
                                if area > best_area:
                                    best, best_area = shp, area
                        except Exception:
                            continue
                    if best is not None:
                        return slide, best
            except Exception:
                pass

        return slide, None

def ppt_insert_text_at_cursor(s):
    """
    Fügt Text ausschließlich an der echten Cursorposition ein.
    Kein Fallback auf Shape-Auswahl → sonst ungewolltes Anhängen.
    """
    with com_context("_activate_powerpoint"):
        app = win32.gencache.EnsureDispatch("PowerPoint.Application")
        win = app.ActiveWindow
        if not win:
            raise RuntimeError("Kein PowerPoint-Fenster aktiv.")

        sel = win.Selection

        # EINZIG erlaubter Fall: echter Textcursor
        try:
            tr = sel.TextRange
            if tr is not None:
                tr.InsertAfter(s)
                return
        except Exception:
            pass

        # alles andere ist NOK
        raise RuntimeError(
            "Kein Textcursor gefunden.\n"
            "Bitte klicke in das Textfeld (Cursor sichtbar) und versuche es erneut."
        )
    
def _copy_font(src_font, dst_font):
    """Kopiert möglichst viele Font-Eigenschaften robust."""
    props = ["Name", "Size", "Bold", "Italic", "Underline", "Color", "BaselineOffset"]
    for p in props:
        try:
            setattr(dst_font, p, getattr(src_font, p))
        except Exception:
            pass

def ppt_insert_hidden_marker(marker_text: str, trailing_text: str = " "):
    """
    Fügt Marker an Cursorposition ein, macht ihn unsichtbar (Hidden),
    und sorgt dafür, dass danach weitergeschrieben wird wie vorher
    (Font/Größe etc.). Optional wird danach ein normales Leerzeichen eingefügt.
    """
    app = win32.gencache.EnsureDispatch("PowerPoint.Application")
    win = app.ActiveWindow
    if not win:
        raise RuntimeError("Kein PowerPoint-Fenster aktiv.")
    sel = win.Selection

    # Basis-Range holen (Textcursor oder Shape-Auswahl)
    base_range = None
    try:
        tr = sel.TextRange
        if tr is not None:
            base_range = tr.Parent.TextRange
    except Exception:
        pass

    if base_range is None:
        try:
            sr = sel.ShapeRange
            if sr is not None and sr.Count >= 1:
                shp = sr.Item(1)
                if shp is not None and getattr(shp, "HasTextFrame", False):
                    base_range = shp.TextFrame.TextRange
        except Exception:
            pass

    if base_range is None:
        raise RuntimeError("Keine Text-Einfügeposition gefunden. Wähle ein Textfeld und setze den Cursor.")

    # Aktuelle Schreib-Formatierung merken (vom Selection-TextRange, fallback base_range)
    try:
        fmt_font = sel.TextRange.Font
    except Exception:
        fmt_font = base_range.Font

    insert_start_abs = base_range.Start + base_range.Length
    to_insert = marker_text + (trailing_text or "")
    base_range.InsertAfter(to_insert)

    # Relativer Start (1-basiert) innerhalb base_range
    rel_start_1b = (insert_start_abs - base_range.Start) + 1

    # Marker-Range
    mr = base_range.Characters(rel_start_1b, len(marker_text))

    # Marker: gleiche Schrift wie Umgebung, dann Hidden
    try:
        _copy_font(fmt_font, mr.Font)
    except Exception:
        pass
    try:
        mr.Font.Hidden = True
    except Exception:
        pass

    # Trailing-Text (z.B. Leerzeichen): sichtbar und Format wie vorher
    if trailing_text:
        trr = base_range.Characters(rel_start_1b + len(marker_text), len(trailing_text))
        try:
            _copy_font(fmt_font, trr.Font)
        except Exception:
            pass
        try:
            trr.Font.Hidden = False
        except Exception:
            pass

def _author_year_parts(item):
    data = item.get("data", {})
    creators = data.get("creators", []) or []
    names = []
    for c in creators:
        last = c.get("lastName") or c.get("name") or ""
        if last:
            names.append(last)
    if not names:
        names = [data.get("title", "o. A.")]

    author = names[0] if len(names) == 1 else (f"{names[0]} & {names[1]}" if len(names) == 2 else f"{names[0]} et al.")
    m = re.search(r"(\d{4})", data.get("date") or "")
    year = m.group(1) if m else ""
    return author, year

def _make_sig(item):
    author, year = _author_year_parts(item)
    year = year or "n.d."   # wichtig: sonst wäre sig nur "Autor"
    return f"{author}|{year}"
    
def collect_all_cites_by_key():
    pres = _get_presentation()
    by_key = {}
    for slide in pres.Slides:
        for shp in slide.Shapes:
            try:
                if getattr(shp, "HasTextFrame", False):
                    for c in prune_cites_in_shape(shp):
                        k = c.get("key")
                        if k and k not in by_key:
                            by_key[k] = c
            except Exception:
                continue
    return by_key
    
def _replace_all(text, old, new):
    if not old or old == new:
        return text, 0
    n = text.count(old)
    return text.replace(old, new), n

def normalize_sig_group(sig):
    """
    APA/Harvard:
    Vergibt a/b/... für alle *verschiedenen Zotero-Keys* mit identischem Autor+Jahr (sig).
    Baut Suffixe auch wieder ab, wenn nach Löschen nur noch 1 Key übrig ist.
    Aktualisiert:
    - sichtbaren Text in allen Shapes
    - gespeicherte Cite-Tags
    """
    pres = _get_presentation()

    # 1) alle Vorkommen dieser sig einsammeln (über Tags)
    occ = []  # (slide, shp, idx, key, old_cite)
    for slide in pres.Slides:
        for shp in slide.Shapes:
            try:
                if not getattr(shp, "HasTextFrame", False):
                    continue
                arr = prune_cites_in_shape(shp)
                for i, c in enumerate(arr):
                    if c.get("sig") == sig:
                        k = c.get("key")
                        oc = c.get("cite") or ""
                        if k and oc:
                            occ.append((slide, shp, i, k, oc))
            except Exception:
                continue

    if not occ:
        return

    # stabile Reihenfolge: erster Fund im Dokument
    keys_in_order = []
    for _, _, _, k, _ in occ:
        if k not in keys_in_order:
            keys_in_order.append(k)

    def strip_suffix(cite: str) -> str:
        # (Autor, 2020a) -> (Autor, 2020)
        # (Autor, n.d.a) -> (Autor, n.d.)
        return re.sub(r"((?:\d{4})|n\.d\.)[a-z]\)$", r"\1)", cite)

    # pro key eine "Basis" (ohne Suffix) merken
    base_by_key = {}
    for _, _, _, k, oc in occ:
        if k not in base_by_key:
            base_by_key[k] = strip_suffix(oc)

    # 2) Zieltexte pro Key bestimmen
    letters = "abcdefghijklmnopqrstuvwxyz"
    new_by_key = {}

    if len(keys_in_order) == 1:
        # ROLLBACK-FALL: a/b entfernen
        k = keys_in_order[0]
        new_by_key[k] = base_by_key.get(k) or strip_suffix(occ[0][4])
    else:
        # a/b/... vergeben
        for idx, k in enumerate(keys_in_order):
            base = base_by_key.get(k) or "(o. A.)"
            new_by_key[k] = base[:-1] + letters[idx] + ")"

    # 3) pro Shape sequenziell ersetzen (ohne Shape als Dict-Key!)
    by_shape = {}  # (slide_id, shape_id) -> {"shape": shp, "items":[(idx,key,old_cite),...]}
    for slide, shp, i, k, old_cite in occ:
        try:
            slide_id = int(slide.SlideID)
            shape_id = int(getattr(shp, "Id", getattr(shp, "ID", -1)))
        except Exception:
            continue
        key = (slide_id, shape_id)
        by_shape.setdefault(key, {"shape": shp, "items": []})["items"].append((i, k, old_cite))

    for _, pack in by_shape.items():
        shp = pack["shape"]
        items = pack["items"]
        try:
            tr = shp.TextFrame.TextRange
            txt = tr.Text or ""
            arr = _load_shape_cites(shp)

            changed = False
            for i, k, old_cite in sorted(items, key=lambda x: x[0]):
                new_cite = new_by_key.get(k, old_cite)
                if new_cite == old_cite:
                    continue

                txt, replaced = _replace_first(txt, old_cite, new_cite)
                if replaced:
                    changed = True

                # Tag updaten (Index passt zur Tag-Liste, solange prune vorher lief)
                if i < len(arr) and arr[i].get("key") == k:
                    arr[i]["cite"] = new_cite

            if changed:
                tr.Text = txt
                _save_shape_cites(shp, arr)

        except Exception:
            continue

def renormalize_all_sig_groups():
    pres = _get_presentation()
    sigs = []
    for slide in pres.Slides:
        for shp in slide.Shapes:
            try:
                if not getattr(shp, "HasTextFrame", False):
                    continue
                arr = prune_cites_in_shape(shp)
                for c in arr:
                    s = c.get("sig")
                    if s and s not in sigs:
                        sigs.append(s)
            except Exception:
                continue

    for s in sigs:
        normalize_sig_group(s)

def _safe_get(url, *, headers=None, params=None, timeout=HTTP_TIMEOUT, retries=2):
    """
    Requests GET mit Retry bei 5xx (Zotero sporadisch).
    Wirft erst nach Retries eine Exception.
    """
    last_exc = None
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=headers, params=params, timeout=timeout)          
            _debug(f"HTTP GET {r.status_code}: {r.url}")
            
            # bei 5xx -> Retry
            if r.status_code in (500, 502, 503, 504):
                last_exc = requests.HTTPError(f"{r.status_code} Server Error for url: {r.url}")
                _debug(f"HTTP RETRY wegen {r.status_code} (attempt {attempt+1}/{retries+1})")
                continue
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            _debug(f"HTTP EXC (attempt {attempt+1}/{retries+1}): {e}")
            last_exc = e
            continue
    _debug(f"HTTP FAIL nach Retries: {last_exc}")
    raise last_exc
# =========================================================


# ===================== Dokument-State =====================
def _get_docprop_by_name(props, name):
    try:
        return props.Item(name)
    except Exception:
        pass
    try:
        for p in props:
            try:
                if getattr(p, "Name", None) == name:
                    return p
            except Exception:
                continue
    except Exception:
        pass
    return None

def load_doc_state():
    pres = _get_presentation()
    props = pres.CustomDocumentProperties
    p = _get_docprop_by_name(props, DOCPROP_NAME)
    if p is None:
        return {}
    try:
        return json.loads(p.Value)
    except Exception:
        return {}

def save_doc_state(state):
    pres = _get_presentation()
    props = pres.CustomDocumentProperties
    payload = json.dumps(state, ensure_ascii=False)
    p = _get_docprop_by_name(props, DOCPROP_NAME)
    if p is not None:
        p.Value = payload
        return
    props.Add(DOCPROP_NAME, False, 4, payload)
# =========================================================


# ============ Bibliographie: stabiler Anker über Tags =====
def _get_shape_tag(shape, key):
    # Tags sind je nach Office-Version manchmal zickig → robust abfragen
    try:
        # VBA-Style: shape.Tags("key")
        return shape.Tags(key)
    except Exception:
        try:
            # Alternative: shape.Tags.Item("key")
            return shape.Tags.Item(key)
        except Exception:
            return ""

def _set_shape_tag(shape, key, value):
    try:
        shape.Tags.Add(key, value)
        return
    except Exception:
        try:
            shape.Tags.Delete(key)
        except Exception:
            pass
        try:
            shape.Tags.Add(key, value)
        except Exception:
            pass

def _load_shape_cites(shp):
    raw = _get_shape_tag(shp, CITE_TAG) or "[]"
    try:
        arr = json.loads(raw)
        return arr if isinstance(arr, list) else []
    except Exception:
        return []

def _save_shape_cites(shp, arr):
    try:
        _set_shape_tag(shp, CITE_TAG, json.dumps(arr, ensure_ascii=False))
    except Exception:
        pass

# LEGACY/OPTIONAL: Zero-width Marker (aktuell nicht verwendet in APA/Harvard; evtl. später wieder nützlich)
def collect_all_cite_texts():
    pres = _get_presentation()
    out = []
    for slide in pres.Slides:
        for shp in slide.Shapes:
            try:
                if getattr(shp, "HasTextFrame", False):
                    for c in prune_cites_in_shape(shp):
                        t = (c.get("cite") or "").strip()
                        if t:
                            out.append(t)
            except Exception:
                continue
    return out

def prune_cites_in_shape(shp):
    """Entfernt gespeicherte Cites, deren cite-Text nicht mehr im Shape-Text vorkommt."""
    arr = _load_shape_cites(shp)
    if not arr:
        return []

    try:
        txt = shp.TextFrame.TextRange.Text or ""
    except Exception:
        return arr

    kept = [c for c in arr if (c.get("cite") or "") in txt]
    if kept != arr:
        _save_shape_cites(shp, kept)
    return kept

def _format_authoryear_base_from_item(item):
    """Basis ohne a/b: (Autor, Jahr) – aus pyzotero Item-Dict."""
    data = item.get("data", {})
    creators = data.get("creators", []) or []
    names = []
    for c in creators:
        last = c.get("lastName") or c.get("name") or ""
        if last:
            names.append(last)
    if not names:
        names = [data.get("title", "o. A.")]

    author = names[0] if len(names) == 1 else (f"{names[0]} & {names[1]}" if len(names) == 2 else f"{names[0]} et al.")
    m = re.search(r"(\d{4})", data.get("date") or "")
    year = m.group(1) if m else ""
    return f"({author}, {year})" if year else f"({author}, n.d.)"

# LEGACY/OPTIONAL: Zero-width Marker (aktuell nicht verwendet in APA/Harvard; evtl. später wieder nützlich)
def _disambiguate_authoryear(base, existing_cites):
    """macht aus (Müller, 2020) -> (Müller, 2020a/b/...) falls nötig"""
    if base not in existing_cites:
        return base
    if not base.endswith(")"):
        return base
    stem = base[:-1]  # ohne )
    suffix = "a"
    while f"{stem}{suffix})" in existing_cites:
        suffix = chr(ord(suffix) + 1)
    return f"{stem}{suffix})"

def _replace_first(text, old, new):
    if not old:
        return text, False
    i = text.find(old)
    if i < 0:
        return text, False
    return text[:i] + new + text[i+len(old):], True

def set_bibliography_anchor_from_selection():
    slide, shp = _get_current_slide_and_shape()
    _debug(f"Anchor-Set: slide={getattr(slide,'SlideID',None)}, shape_id={getattr(shp,'Id',getattr(shp,'ID',None))}, hasTextFrame={getattr(shp,'HasTextFrame',False)}")

    if slide is None:
        raise RuntimeError("Keine aktive Folie gefunden.")

    # ROBUSTE Prüfung: echtes Textfeld?
    if shp is None or not getattr(shp, "HasTextFrame", False):
        raise RuntimeError(
            "Bitte wähle ein Textfeld aus.\n"
            "Hinweis: Textfeld anklicken (Rahmen sichtbar), "
            "nicht nur den Cursor setzen."
        )

    # absichtlich KEIN `shp.TextFrame.HasText` → Bibliographie-Felder dürfen leer sein

    st = load_doc_state()
    bib_guid = st.get("bib_guid") or str(uuid.uuid4())
    st["bib_guid"] = bib_guid

    # harter Anker über IDs (Id vs. ID robust abfangen)
    try:
        shape_id = int(getattr(shp, "Id", getattr(shp, "ID", -1)))
        slide_id = int(slide.SlideID)

        if shape_id > 0:
            st["bib_anchor"] = {
                "slide_id": slide_id,
                "shape_id": shape_id,
            }
        else:
            st.pop("bib_anchor", None)

    except Exception:
        st.pop("bib_anchor", None)

    save_doc_state(st)
    _debug(f"Anchor gespeichert: bib_guid={st.get('bib_guid')}, bib_anchor={st.get('bib_anchor')}")

    # 1) Tags versuchen (nice-to-have)
    _set_shape_tag(shp, TAG_BIB_GUID_KEY, bib_guid)

    # 2) ROBUST: AlternativeText setzen (persistiert zuverlässig)
    try:
        shp.AlternativeText = ALT_BIB_PREFIX + bib_guid
    except Exception:
        # manche Shapes blocken das – dann wenigstens Tag
        pass
    _debug("Anchor getaggt: Tag + AlternativeText gesetzt (falls möglich)")

    # 3) Sofort prüfen, ob wir das Ziel wiederfinden
    if not has_bibliography_anchor():
        _debug("Anchor-Check FAIL: _resolve_anchor_list() liefert 0")
        raise RuntimeError(
            "Bibliographie-Ziel konnte nicht gespeichert werden (PowerPoint hat den Anker nicht übernommen).\n"
            "Bitte: Textfeld einmal anklicken (Rahmen sichtbar), dann erneut „Bibliographie-Ziel…“."
        )

def _resolve_anchor_list():
    pres = _get_presentation()
    st = load_doc_state()
    bib_guid = st.get("bib_guid")
    anch = st.get("bib_anchor") or {}
    
    if bib_guid or anch:
        _debug(f"Resolve anchors: bib_guid={bib_guid}, bib_anchor={anch}")

    if not bib_guid and not anch:
        return []

    resolved = []
    seen = set()  # (slide_id, shape_id)

    # 1) Direkt über SlideID/ShapeID (am zuverlässigsten)
    slide_id = anch.get("slide_id")
    shape_id = anch.get("shape_id")
    if slide_id and shape_id:
        try:
            slide_id = int(slide_id)
            shape_id = int(shape_id)
            for slide in pres.Slides:
                if int(slide.SlideID) != slide_id:
                    continue
                for shp in slide.Shapes:
                    sid = int(getattr(shp, "Id", getattr(shp, "ID", -1)))
                    if sid == shape_id and getattr(shp, "HasTextFrame", False):
                        key = (int(slide.SlideID), sid)
                        if key not in seen:
                            resolved.append((slide, shp))
                            seen.add(key)
                        break
                break
        except Exception:
            pass

    # 2) Zusätzlich: alle Shapes mit GUID finden (für Fortsetzungsfolien / Duplikate)
    if bib_guid:
        for slide in pres.Slides:
            for shp in slide.Shapes:
                try:
                    if not getattr(shp, "HasTextFrame", False):
                        continue

                    sid = int(getattr(shp, "Id", getattr(shp, "ID", -1)))
                    key = (int(slide.SlideID), sid)
                    if key in seen:
                        continue

                    try:
                        alt = shp.AlternativeText or ""
                        if alt.strip() == (ALT_BIB_PREFIX + bib_guid):
                            resolved.append((slide, shp))
                            seen.add(key)
                            continue
                    except Exception:
                        pass

                    if _get_shape_tag(shp, TAG_BIB_GUID_KEY) == bib_guid:
                        resolved.append((slide, shp))
                        seen.add(key)

                except Exception:
                    continue
    
    _debug(f"Resolve anchors: gefunden={len(resolved)}")
    return resolved

def has_bibliography_anchor():
    return len(_resolve_anchor_list()) > 0

def get_status_summary():
    state = load_doc_state()
    style = state.get("style", DEFAULT_STYLE)
    keys = state.get("bib_keys", [])

    anchors = _resolve_anchor_list()
    if anchors:
        anchor_txt = f"gesetzt ({len(anchors)} Feld(er))"
    else:
        anchor_txt = "NICHT gesetzt"

    return f"Stil: {style.upper()} | Zitate: {len(keys)} | Bibliographie-Ziel: {anchor_txt}"

def _is_title_placeholder(shp) -> bool:
    """True, wenn Shape sehr wahrscheinlich ein Titel-Placeholder ist."""
    try:
        t = int(getattr(shp.PlaceholderFormat, "Type", -1))
        if t in (1, 3):  # Title + Center Title
            return True
    except Exception:
        pass

    # zusätzliche Heuristik (Fallback)
    try:
        name = (getattr(shp, "Name", "") or "").lower()
        if "title" in name:
            return True
    except Exception:
        pass

    return False

def _placeholder_type(shp):
    try:
        return int(shp.PlaceholderFormat.Type)
    except Exception:
        return None

def _find_best_text_placeholder(slide, src_shape=None):
    """
    Sucht auf einer Folie ein geeignetes Text-Placeholder-Feld (Layout-Shape),
    bevorzugt gleichen Placeholder-Typ wie src_shape (falls src_shape Placeholder ist).
    """
    want_type = _placeholder_type(src_shape) if src_shape is not None else None
    allowed = {2, 7}  # Body, Content
    if want_type not in allowed:
        want_type = None

    best = None
    best_score = -1

    for shp in slide.Shapes:
        try:
            if not getattr(shp, "HasTextFrame", False):
                continue
            # Placeholder-Objekt?
            _ = shp.PlaceholderFormat  # wirft Exception, wenn kein Placeholder
        except Exception:
            continue

        try:
            # Titel-Placeholder überspringen
            if _is_title_placeholder(shp):
                continue
        except Exception:
            pass

        # bevorzugt: gleicher Placeholder-Typ wie Quelle
        score = 0
        try:
            ptype = _placeholder_type(shp)
            if ptype not in allowed:
                continue

            if want_type is not None and ptype == want_type:
                score += 1000
        except Exception:
            continue

        # zweitbeste Heuristik: größtes Textfeld
        try:
            score += int(float(shp.Width) * float(shp.Height))
        except Exception:
            score += 0

        if score > best_score:
            best = shp
            best_score = score
        
    if best is None:
        _debug("Find placeholder: keiner gefunden (allowed types: {2,7})")

    return best

def _get_slide_title_text(slide):
    """Liest den Text des Titel-Placeholders einer Folie."""
    try:
        for shp in slide.Shapes:
            try:
                if int(shp.PlaceholderFormat.Type) in (1, 3):  # Title oder Center Title
                    if shp.TextFrame.HasText:
                        return shp.TextFrame.TextRange.Text
            except Exception:
                continue
    except Exception:
        pass
    return ""

def _set_slide_title_text(slide, text):
    """Setzt den Text des Titel-Placeholders einer Folie."""
    if not text:
        return
    try:
        for shp in slide.Shapes:
            try:
                if int(shp.PlaceholderFormat.Type) in (1, 3):  # Title oder Center Title
                    shp.TextFrame.TextRange.Text = text
                    return
            except Exception:
                continue
    except Exception:
        pass


def _duplicate_anchor_to_new_slide_like(src_slide, src_shape):
    pres = _get_presentation()

    # Titel der Quellfolie merken
    title_text = _get_slide_title_text(src_slide)

    # neue Folie mit gleichem Layout
    new_slide = pres.Slides.AddSlide(pres.Slides.Count + 1, src_slide.CustomLayout)
    _debug(f"Neue Bibliographie-Folie erzeugt (Layout: {src_slide.CustomLayout.Name})")

    # Titel übernehmen
    if title_text:
        _set_slide_title_text(new_slide, title_text)
        _debug(f"Titel übernommen: '{title_text}'")

    # 1) Versuche: passendes Layout-Placeholder-Textfeld finden
    new_shape = _find_best_text_placeholder(new_slide, src_shape=src_shape)

    if new_shape is not None:
        try:
            ptype = int(new_shape.PlaceholderFormat.Type)
        except Exception:
            ptype = "?"

        _debug(f"Layout-Placeholder gefunden (Type={ptype}, Name='{getattr(new_shape, 'Name', '')}')")

        try:
            new_shape.TextFrame.TextRange.Text = ""
        except Exception:
            pass
    else:
        _debug("WARNUNG: Kein geeignetes Text-Placeholder im Layout gefunden → Copy/Paste-Fallback")

        # Fallback
        src_shape.Copy()
        pasted = new_slide.Shapes.Paste()
        new_shape = pasted.Item(1)
        try:
            new_shape.TextFrame.TextRange.Text = ""
        except Exception:
            pass

    # gleiche GUID taggen (Fortsetzung gehört zum selben Bib-Set)
    st = load_doc_state()
    bib_guid = st.get("bib_guid")
    if bib_guid and new_shape is not None:
        _set_shape_tag(new_shape, TAG_BIB_GUID_KEY, bib_guid)
        try:
            new_shape.AlternativeText = ALT_BIB_PREFIX + bib_guid
        except Exception:
            pass

    return new_slide, new_shape
# =========================================================


# ============ Zotero: Bibliographie-Einträge ==============
def html_to_text(html_str):
    text = re.sub(r"<br\s*/?>", "\n", html_str, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    return re.sub(r"\s+\n", "\n", text).strip()

def get_bibliography_entry_webapi(api_key, library_id, library_type, item_key, style):
    base = f"https://api.zotero.org/{library_type}s/{library_id}"
    headers = {
        "Zotero-API-Key": api_key,
        "Zotero-API-Version": "3",
        "Accept": f"text/x-bibliography; style={style}",
    }
    candidates = [
        (f"{base}/items", {"itemKey": item_key, "format": "bib", "style": style}),
        (f"{base}/items/{item_key}", {"format": "bib", "style": style}),
    ]
    for url, params in candidates:
        try:
            r = _safe_get(url, headers=headers, params=params, timeout=HTTP_TIMEOUT, retries=2)
        except Exception:
            # probiere nächsten Candidate / später JSON-Fallback
            continue

        ct = (r.headers.get("Content-Type", "") or "").lower()
        txt = r.text

        looks_like_json = txt.lstrip().startswith("{") or txt.lstrip().startswith("[")
        if (("text/html" in ct) or ("text/x-bibliography" in ct)) and not looks_like_json:
            return html_to_text(txt)

    try:
        raw = _safe_get(
            f"{base}/items/{item_key}",
            headers={
                "Zotero-API-Key": api_key,
                "Zotero-API-Version": "3",
                "Accept": "application/json",
            },
            timeout=HTTP_TIMEOUT,
            retries=2
        )
    except Exception as e:
        # harter Fallback (minimal)
        return f"[Bibliographie nicht verfügbar: {item_key}]"

    data = raw.json()
    if isinstance(data, list) and data:
        data = data[0]
    d = data.get("data", {})
    title = d.get("title") or "[o. T.]"
    creators = d.get("creators") or []
    first = creators[0].get("lastName") or creators[0].get("name") if creators else ""
    m = re.search(r"(\d{4})", d.get("date") or "")
    year = m.group(1) if m else ""
    url = d.get("url") or ""
    pieces = [p for p in [first, f"({year})" if year else "", title, url] if p]
    return " ".join(pieces)
# =========================================================


# ======== Bibliographie: Schreiben & Paginierung =========
def _try_fit_entries_into_shape(shape, entries, preferred_size=PREF_FONT_SIZE, min_size=MIN_FONT_SIZE, line_sep="\r"):
    _debug(f"Fit: entries={len(entries)} pref={preferred_size} min={min_size} shapeH={getattr(shape,'Height',None)}")

    tr = shape.TextFrame.TextRange
    tr.Text = ""
    
    def write_with_size(sz, upto):
        tr.Text = ""
        for i in range(upto):
            tr.InsertAfter(entries[i] + line_sep)
        try:
            tr.Font.Size = sz
        except Exception:
            pass

    def overflows():
        try:
            return tr.BoundHeight > (shape.Height - 2)
        except Exception:
            return False

    for count in range(len(entries), 0, -1):
        size = preferred_size
        while size >= min_size:
            write_with_size(size, count)
            if not overflows():                
                _debug(f"Fit OK: count={count} size={size}")
                return count, size
            size -= 1

    if entries:
        write_with_size(min_size, 1)
        _debug("Fit WARN: nur 1 Eintrag passt (min size)")
        return 1, min_size
        
    return 0, preferred_size

def update_bibliography(keys, style, api_key, library_id, library_type, numbering=None):    
    _debug(f"Bib update: keys={len(keys)} style={style}")

    if not keys:
        return

    anchors = _resolve_anchor_list()
    _debug(f"Bib update: anchors={len(anchors)}")

    if not anchors:
        return

    # 1) kaputte Anchors rausfiltern (einmal!)
    safe_anchors = []
    for slide, shape in anchors:
        try:
            if shape is not None and getattr(shape, "HasTextFrame", False):
                _ = shape.TextFrame  # Zugriffstest
                safe_anchors.append((slide, shape))
        except Exception:
            continue

    anchors = safe_anchors
    _debug(f"Bib update: safe_anchors={len(anchors)}")

    if not anchors:
        return

    # 2) Bibliographie-Einträge erzeugen
    entries = []
    for k in keys:
        try:
            entry = get_bibliography_entry_webapi(api_key, library_id, library_type, k, style)
        except Exception:
            entry = f"[Bibliographie nicht verfügbar: {k}]"

        if numbering and k in numbering:
            entries.append(f"[{numbering[k]}] {entry}")
        else:
            entries.append(entry)


    remaining = entries[:] 
    _debug(f"Bib entries erzeugt: {len(entries)}")

    # 3) ZUERST bestehende Bibliographie-Textfelder füllen
    for slide, shape in anchors:
        if not remaining:
            break
        used, _ = _try_fit_entries_into_shape(shape, remaining)
        remaining = remaining[used:]

    # 4) NUR WENN NOCH REST → neue Folien erzeugen
    _debug(f"Bib pagination: remaining_start={len(remaining)}")
    while remaining:
        src_slide, src_shape = anchors[-1]
        new_slide, new_shape = _duplicate_anchor_to_new_slide_like(src_slide, src_shape)        
        _debug(f"Bib pagination: neue Folie erstellt, remaining_before_fit={len(remaining)}")

        # Sicherheitscheck
        if new_shape is None or not getattr(new_shape, "HasTextFrame", False):
            _debug("Bib pagination ABORT: new_shape ungültig oder ohne TextFrame")
            break

        anchors.append((new_slide, new_shape))
        used, _ = _try_fit_entries_into_shape(new_shape, remaining)        
        _debug(f"Bib pagination: used={used}, remaining_after_fit={len(remaining)}")
        
        remaining = remaining[used:]

# =========================================================


# =========== IEEE: Platzhalter & Renummerierung ==========
PH_RE = re.compile(r"⟦zp:([A-Za-z0-9]+)⟧")

def scan_all_placeholders():
    pres = _get_presentation()
    hits = []
    for si, slide in enumerate(pres.Slides, start=1):
        for hi, shp in enumerate(slide.Shapes, start=1):
            try:
                if shp.HasTextFrame and shp.TextFrame.HasText:
                    txt = shp.TextFrame.TextRange.Text
                    for m in PH_RE.finditer(txt):
                        hits.append((si, hi, m.start(), m.group(1)))
            except Exception:
                continue
    hits.sort(key=lambda x: (x[0], x[1], x[2]))
    return hits

def resync_bibliography_keys_from_document(state=None):
    pres = _get_presentation()
    keys = []

    for slide in pres.Slides:
        for shp in slide.Shapes:
            try:
                if shp.HasTextFrame:
                    kept = prune_cites_in_shape(shp)
                    for c in kept:
                        k = c.get("key")
                        if k and k not in keys:
                            keys.append(k)
            except Exception:
                continue

    cur = load_doc_state()
    cur["bib_keys"] = keys
    save_doc_state(cur)

    if state is not None:
        state.clear()
        state.update(cur)

    return keys

def insert_ieee_placeholder(key, *, parent=None) -> bool:
    ppt_insert_text_at_cursor(f" ⟦zp:{key}⟧")
    return renumber_ieee_and_update(parent=parent)

def renumber_ieee_and_update(*, parent=None) -> bool:
    try:
        cfg = get_cfg(allow_prompt=False, parent=parent)
    except RuntimeError:
        if parent is not None:
            show_missing_zotero_config(parent)
        return False
    
    state = load_doc_state()
    style = state.get("style") or DEFAULT_STYLE

    order_keys = [h[3] for h in scan_all_placeholders()]
    numbering, n = {}, 1
    for k in order_keys:
        if k not in numbering:
            numbering[k] = n
            n += 1

    pres = _get_presentation()
    for slide in pres.Slides:
        for shp in slide.Shapes:
            try:
                if shp.HasTextFrame and shp.TextFrame.HasText:
                    old = shp.TextFrame.TextRange.Text
                    new = PH_RE.sub(lambda m: f"[{numbering.get(m.group(1), '?')}]", old)
                    if new != old:
                        shp.TextFrame.TextRange.Text = new
            except Exception:
                continue

    state["bib_keys"] = list(numbering.keys())
    save_doc_state(state)

    if has_bibliography_anchor():
        update_bibliography(
            state["bib_keys"], style,
            cfg.api_key, cfg.library_id, cfg.library_type,
            numbering=numbering
        )

    return True
# =========================================================


def code_to_label(code):
    return next((name for name, c in STYLE_CHOICES if c == code), code)

def label_to_code(label):
    return next((c for name, c in STYLE_CHOICES if name == label), DEFAULT_STYLE)

def run_in_thread(fn, *, on_error=None):
    def _wrap():
        try:
            fn()
        except Exception as e:
            if on_error:
                on_error(e)
            else:
                _debug(f"Thread error: {e}")
    threading.Thread(target=_wrap, daemon=True).start()

def show_missing_zotero_config(parent):
    messagebox.showerror(
        "Zotero not configured",
        "Zotero credentials are not configured yet.\n\n"
        "Please open the Zotero Picker, enter your Zotero API key and "
        "Library ID, and save them.\n\n"
        "Then try again.",
        parent=parent
    )

class PickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zotero Picker → PowerPoint")

        # Optional: Mindestgröße, damit UI nicht "gequetscht" wirkt
        self.root.minsize(420, 300)

        self.cfg = None
        self.z = None

        try:
            self.state = load_doc_state()
            self.state.setdefault("style", DEFAULT_STYLE)
            self.state.setdefault("bib_keys", [])
            save_doc_state(self.state)
        except Exception as e:
            self.state = {"style": DEFAULT_STYLE, "bib_keys": []}
            self._ppt_startup_error = str(e)   # merken für später

        self.results = []

        # Suche: Debounce + "latest-only" (gegen Flackern / Out-of-order Threads)
        self._search_after_id = None
        self._search_token = 0

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="both", expand=True)

        row0 = ttk.Frame(frm); row0.pack(fill="x", pady=(0,6))
        ttk.Label(row0, text="Zitierstil:").pack(side="left")
        self.style_var = tk.StringVar(value=code_to_label(self.state["style"]))
        self.style_combo = ttk.Combobox(
            row0, textvariable=self.style_var, state="readonly",
            values=[name for name,_ in STYLE_CHOICES], width=28
        )
        self.style_combo.pack(side="left", padx=6)
        self.style_combo.bind("<<ComboboxSelected>>", self.on_style_change)

        row1 = ttk.Frame(frm); row1.pack(fill="x")
        ttk.Label(row1, text="Suche:").pack(side="left")
        self.query_var = tk.StringVar()
        self.entry = ttk.Entry(row1, textvariable=self.query_var)
        self.entry.pack(side="left", fill="x", expand=True)
        self.entry.bind("<KeyRelease>", self.on_key)
        self.entry.bind("<Return>", self.on_insert_click)

        # Ergebnisliste + Scrollbars (vertikal + horizontal)
        listfrm = ttk.Frame(frm)
        listfrm.pack(fill="both", expand=True, pady=8)

        self.listbox = tk.Listbox(listfrm, height=16)
        self.listbox.bind("<Button-1>", lambda e: self.listbox.focus_set())

        vsb = ttk.Scrollbar(listfrm, orient="vertical", command=self.listbox.yview)
        hsb = ttk.Scrollbar(listfrm, orient="horizontal", command=self.listbox.xview)

        # Listbox <-> Scrollbars koppeln
        self.listbox.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid-Layout: Listbox groß, V-Scrollbar rechts, H-Scrollbar unten über volle Breite
        self.listbox.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, columnspan=2, sticky="ew")

        # Frame-Resize-Regeln
        listfrm.rowconfigure(0, weight=1)
        listfrm.columnconfigure(0, weight=1)

        self.listbox.bind("<Double-Button-1>", self.on_insert_click)
        
        # --- Optional UX: horizontal scrollen mit Shift + Mausrad ---
        def _on_mousewheel(ev):
            # Windows/macOS: ev.delta; Shift gedrückt -> horizontal, sonst vertikal
            shift = ((ev.state & 0x0001) != 0) or ((ev.state & 0x0004) != 0)  # je nach Tk/Platform
            if shift:
                # X-Scroll: je nach Wheelrichtung
                step = -1 if ev.delta > 0 else 1
                self.listbox.xview_scroll(step, "units")
            else:
                step = -1 if ev.delta > 0 else 1
                self.listbox.yview_scroll(step, "units")
            return "break"

        def _on_shift_mousewheel(ev):
            # explizit Shift-Variante (falls Tk das separat liefert)
            step = -1 if ev.delta > 0 else 1
            self.listbox.xview_scroll(step, "units")
            return "break"

        # Windows / macOS
        self.listbox.bind("<MouseWheel>", _on_mousewheel)
        self.listbox.bind("<Shift-MouseWheel>", _on_shift_mousewheel)

        # Linux (X11): Mausrad sind Button-Events
        def _on_linux_wheel_up(ev):
            if (ev.state & 0x0001) or (ev.state & 0x0004):
                self.listbox.xview_scroll(-1, "units")
            else:
                self.listbox.yview_scroll(-1, "units")
            return "break"

        def _on_linux_wheel_down(ev):
            if (ev.state & 0x0001) or (ev.state & 0x0004):
                self.listbox.xview_scroll(1, "units")
            else:
                self.listbox.yview_scroll(1, "units")
            return "break"

        self.listbox.bind("<Button-4>", _on_linux_wheel_up)
        self.listbox.bind("<Button-5>", _on_linux_wheel_down)

        # Buttons: 2x2 Grid statt eine Zeile (kein Verschwinden bei kleiner Höhe)
        row2 = ttk.Frame(frm)
        row2.pack(fill="x", pady=4)

        buttons = [
            ("Zitation einfügen", self.on_insert_click),
            ("Bibliographie einfügen/aktualisieren", self.on_bib_update),
            ("Bibliographie-Ziel festlegen", self.on_set_anchor),
            ("Bereinigen", self.on_cleanup),
        ]

        for i, (label, cmd) in enumerate(buttons):
            btn = ttk.Button(row2, text=label, command=cmd)

            # Optional: ruhigeres Layout (ähnliche Breite) + sauberes Tab-Fokus-Verhalten
            btn.configure(takefocus=True)

            btn.grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=2)

        row2.columnconfigure(0, weight=1)
        row2.columnconfigure(1, weight=1)

        self.status = ttk.Label(frm, text="", anchor="w")
        self.status.pack(fill="x", pady=(6,0))

        self.entry.focus_set()

        self.set_status("Bereit.")

        if getattr(self, "_ppt_startup_error", None):
            self.set_status(f"PowerPoint nicht bereit: {self._ppt_startup_error}")

        # Config erst nach GUI-Start prüfen (nie blockierend im __init__)
        self.root.after(0, self._ensure_cfg_ready)

        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(300, lambda: self.root.attributes("-topmost", False))

    def set_status(self, msg=""):
        try:
            base = get_status_summary()
        except Exception as e:
            base = f"(Status nicht verfügbar: {e})"

        if msg:
            self.status.config(text=f"{msg}   |   {base}")
        else:
            self.status.config(text=base)

    def _ensure_cfg_ready(self):
        # 1) zuerst ohne Prompt probieren (nicht blockierend)
        try:
            self.cfg = get_cfg(allow_prompt=False, parent=self.root)
            self.z = zotero.Zotero(
                self.cfg.library_id,
                self.cfg.library_type,
                self.cfg.api_key,
            )
            self.set_status("Bereit.")
            return
        except RuntimeError:
            pass

        # 2) jetzt ist die GUI lebendig → Prompt darf modal sein
        try:
            self.cfg = get_cfg(allow_prompt=True, parent=self.root)
            self.z = zotero.Zotero(
                self.cfg.library_id,
                self.cfg.library_type,
                self.cfg.api_key,
            )
            self.set_status("Zotero konfiguriert.")
        except RuntimeError as e:
            messagebox.showerror(
                "Zotero Config",
                str(e),
                parent=self.root,
            )
            self.root.destroy()

    def on_style_change(self, event=None):
        self.state["style"] = label_to_code(self.style_var.get())
        save_doc_state(self.state)
        self.set_status(f"Stil gesetzt: {self.style_var.get()} ({self.state['style']})")

    def on_key(self, event=None):
        if self.z is None:
            self.set_status("Bitte Zotero konfigurieren…")
            return
        
        # Debounce: nicht bei jedem Key sofort suchen
        q = self.query_var.get().strip()
        if self._search_after_id is not None:
            try:
                self.root.after_cancel(self._search_after_id)
            except Exception:
                pass
  
        def _kickoff():
            self._search_token  += 1
            token = self._search_token
            threading.Thread(target=self.search, args=(q, token), daemon=True).start()

        self._search_after_id = self.root.after(180, _kickoff)

    def search(self, q, token):
        if not q:
            self.update_results([], token=token, q=q)
            return
        try:
            items = self.z.items(q=q, limit=MAX_RESULTS)
            self.update_results(items, token=token, q=q)
        except Exception as e:
            self.root.after(0, lambda: self.set_status(f"Suche fehlgeschlagen: {e}"))

    def update_results(self, items, token=None, q=None):
        # NUR Dicts behalten
        incoming = [it for it in (items or []) if isinstance(it, dict)]
        
        def _ui():
            # "Latest-only": ältere / überholte Such-Threads ignorieren (Flackern weg)
            if token is not None and token != self._search_token:
                return
            if q is not None and q != self.query_var.get().strip():
                return

            self.results = incoming
            self.listbox.delete(0, tk.END)
            for it in self.results:
                d = it.get("data", {})
                title = d.get("title", "(ohne Titel)")
                creators = d.get("creators", [])
                year = ""
                m = re.search(r"(\d{4})", d.get("date") or "")
                if m: year = m.group(1)
                author = ""
                if creators:
                    last = creators[0].get("lastName") or creators[0].get("name") or ""
                    if len(creators) == 1:
                        author = last
                    elif len(creators) == 2:
                        second = creators[1].get("lastName") or creators[1].get("name") or ""
                        author = f"{last} & {second}"
                    else:
                        author = f"{last} et al."
                self.listbox.insert(tk.END, f"{title} — {author} {year}".strip())
            self.set_status(f"{len(self.results)} Treffer")
        self.root.after(0, _ui)

    def current_item(self):
        idxs = self.listbox.curselection()
        return self.results[idxs[0]] if idxs else None

    def _format_authoryear_quick(self, item):
        data = item.get("data", {})
        creators = data.get("creators", [])
        names = []
        for c in creators:
            last = c.get("lastName") or c.get("name") or ""
            if last: names.append(last)
        if not names:
            names = [data.get("title", "o. A.")]
        author = names[0] if len(names) == 1 else (f"{names[0]} & {names[1]}" if len(names) == 2 else f"{names[0]} et al.")
        m = re.search(r"(\d{4})", data.get("date") or "")
        year = m.group(1) if m else ""
        return f"({author}, {year})" if year else f"({author}, n.d.)"

    def on_insert_click(self, event=None):
        # 0) Aktuellen Treffer holen
        it = self.current_item()
        if not it:
            self.set_status("Kein Eintrag ausgewählt.")
            return

        key = it.get("key")
        style = self.state.get("style", DEFAULT_STYLE)
        _debug(f"Insert click: style={style}, key={key}")

        try:
            # 1) Sonderfall IEEE (Platzhalter + Renummerierung)
            if style == "ieee":
                ok = insert_ieee_placeholder(key, parent=self.root)
                if ok:
                    self.set_status("IEEE citation inserted and renumbered.")
                else:
                    self.set_status("IEEE placeholder inserted. Configure Zotero to renumber & update bibliography.")
                return

            # 2) Ziel in PowerPoint bestimmen
            #    → wir brauchen das Shape für das Cite-Tagging
            slide, shp = _get_current_slide_and_shape()
            _debug(
                f"Insert target: slide={getattr(slide,'SlideID',None)}, "
                f"shape_id={getattr(shp,'Id',getattr(shp,'ID',None))}"
            )

            if shp is None or not getattr(shp, "HasTextFrame", False):
                raise RuntimeError("Bitte Textfeld auswählen und Cursor setzen.")

            # 3) Bereits vorhandene Zitate nach Zotero-Key sammeln
            #    → gleiche Quelle = exakt gleicher Cite-Text
            by_key = collect_all_cites_by_key()

            if style in ("apa", "harvard1") and key in by_key:
                # gleicher Zotero-Key → Cite 1:1 wiederverwenden
                cite = by_key[key].get("cite") or _format_authoryear_base_from_item(it)
                sig  = by_key[key].get("sig")  or _make_sig(it)
            else:
                # neue Quelle → neue Disambiguierungsgruppe (Autor + Jahr)
                sig = _make_sig(it)
                cite = _format_authoryear_base_from_item(it)

            # 4) Zitat an Cursorposition einfügen
            #    → ppt_insert_text_at_cursor wirft selbst einen Fehler,
            #      wenn kein Cursor vorhanden ist (NOK-Fall sauber abgefangen)
            ppt_insert_text_at_cursor(cite)
            _debug(f"Inserted cite: {cite}")

            # 5) Cite im Shape-Tag speichern (Key, Text, Sig)
            arr = _load_shape_cites(shp)
            arr.append({
                "key":  key,
                "cite": cite,
                "sig":  sig,
            })
            _save_shape_cites(shp, arr)

            # 6) APA / Harvard:
            #    Falls mehrere unterschiedliche Keys dieselbe sig teilen,
            #    a/b/... vergeben bzw. rückbauen
            if style in ("apa", "harvard1"):
                normalize_sig_group(sig)
                _debug(f"Normalized sig group: {sig}")

            # 7) Bibliographie-Keys aus dem Dokument neu ableiten
            #    (inkl. Bereinigung gelöschter Cites)
            self.state["bib_keys"] = resync_bibliography_keys_from_document(self.state)
            _debug(f"Resync keys: {len(self.state.get('bib_keys', []))}")

            # 8) Auto-Update der Bibliographie, falls ein Anker existiert
            if has_bibliography_anchor():
                try:
                    cfg = get_cfg(allow_prompt=False)
                except RuntimeError:
                    show_missing_zotero_config(self.root)
                    return
                update_bibliography(
                    self.state.get("bib_keys", []),
                    style,
                    cfg.api_key,
                    cfg.library_id,
                    cfg.library_type
                )
                _debug("Auto bibliography update triggered")
                self.set_status(f"Eingefügt: {cite} (Bibliographie aktualisiert)")
            else:
                self.set_status(
                    f"Eingefügt: {cite} "
                    f"(Bibliographie wird geschrieben, sobald Ziel gesetzt ist)"
                )

        except Exception as e:
            # 9) Saubere Fehlermeldung an den User
            messagebox.showerror("Fehler", str(e), parent=self.root)
        
    def on_set_anchor(self):
        # nicht mehrfach öffnen
        if getattr(self, "_anchor_win", None) and self._anchor_win.winfo_exists():
            try:
                self._anchor_win.lift()
            except Exception:
                pass
            return

        win = tk.Toplevel(self.root)
        self._anchor_win = win

        win.title("Bibliographie-Ziel setzen")
        win.transient(self.root)
        win.attributes("-topmost", True)

        ttk.Label(
            win,
            text=(
                "1) Wechsle jetzt zu PowerPoint.\n"
                "2) Klicke das Bibliographie-Textfeld an (Rahmen sichtbar).\n"
                "3) Komm zurück und klicke hier auf „Jetzt setzen“."
            ),
            justify="left",
            padding=10
        ).pack(fill="x")

        btn_row = ttk.Frame(win, padding=10)
        btn_row.pack(fill="x")

        def _do_set():
            try:
                # User hat jetzt wirklich Zeit gehabt, in PPT zu klicken
                set_bibliography_anchor_from_selection()

                # state frisch aus DocProps (mit bib_anchor/bib_guid)
                self.state = load_doc_state()

                # Keys scannen (speichert selbst in DocProps)
                self.state["bib_keys"] = resync_bibliography_keys_from_document(self.state)

                anchor_count = len(_resolve_anchor_list())

                if self.state.get("bib_keys"):
                    try:
                        cfg = get_cfg(allow_prompt=False)
                    except RuntimeError:
                        show_missing_zotero_config(self.root)
                        return
                    style = self.state.get("style", DEFAULT_STYLE)
                    
                    def _do_bib_update():
                        update_bibliography(
                            self.state.get("bib_keys", []),
                            style,
                            cfg.api_key,
                            cfg.library_id,
                            cfg.library_type
                        )
                        # Status-Update zurück in den UI-Thread
                        self.root.after(
                            0,
                            lambda: self.set_status(
                                f"Bibliographie-Ziel gesetzt. Gefundene Anker: {anchor_count} (Bibliographie aktualisiert)"
                            )
                        )

                    run_in_thread(
                        _do_bib_update,
                        on_error=lambda e: self.root.after(
                            0, lambda: messagebox.showerror("Fehler", str(e), parent=self.root)
                        )
                    )

                else:
                    self.set_status(
                        f"Bibliographie-Ziel gesetzt. Gefundene Anker: {anchor_count} (noch keine Zitate)"
                    )

                self._anchor_win = None
                win.destroy()

            except Exception as e:
                messagebox.showerror("Fehler", str(e), parent=self.root)

        ttk.Button(btn_row, text="Jetzt setzen", command=_do_set).pack(side="left")

        def _cancel():
            self._anchor_win = None
            win.destroy()

        ttk.Button(btn_row, text="Abbrechen", command=_cancel).pack(side="right")

        # --- Fenster über dem Picker zentrieren (NACH dem Layout!) ---
        self.root.update_idletasks()
        win.update_idletasks()

        confirm_w = win.winfo_width()
        confirm_h = win.winfo_height()

        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        # gewünschte Position (zentriert über Picker)
        x = root_x + (root_w - confirm_w) // 2
        y = root_y + (root_h - confirm_h) // 2

        # --- Clamp auf Bildschirm ---
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()

        x = max(0, min(x, screen_w - confirm_w))
        y = max(0, min(y, screen_h - confirm_h))

        win.geometry(f"+{x}+{y}")

        try:
            win.lift()
        except Exception:
            pass
        
        # PPT in den Vordergrund holen, damit der User klicken kann
        _activate_powerpoint()

        # Fenster soll beim Zurückkommen sicher über dem Picker liegen
        try:
            win.attributes("-topmost", True)
            win.lift()
        except Exception:
            pass

        _activate_powerpoint()

    def on_bib_update(self):
        # Keys aktualisieren
        self.state["bib_keys"] = resync_bibliography_keys_from_document(self.state)
        keys = self.state.get("bib_keys", [])
        if not keys:
            self.set_status("Bibliographie leer.")
            return

        if not has_bibliography_anchor():
            messagebox.showerror(
                "Fehlendes Ziel",
                "Kein Bibliographie-Ziel gesetzt.\n"
                "Bitte gehe zur Bibliographie-Folie, wähle das Textfeld und klicke "
                "„Bibliographie-Ziel festlegen“.",
                parent=self.root
            )
            return

        try:
            cfg = get_cfg(allow_prompt=False)
        except RuntimeError:
            show_missing_zotero_config(self.root)
            return
        style = self.state.get("style", DEFAULT_STYLE)

        # UI sofort: "läuft..."
        self.set_status("Bibliographie wird aktualisiert...")

        def _do_bib_update():
            update_bibliography(keys, style, cfg.api_key, cfg.library_id, cfg.library_type)
            self.root.after(
                0,
                lambda: self.set_status(f"Bibliographie aktualisiert ({style}).")
            )

        run_in_thread(
            _do_bib_update,
            on_error=lambda e: self.root.after(
                0, lambda: messagebox.showerror("Fehler", str(e), parent=self.root)
            )
        )
        
    def on_cleanup(self):
        _debug("Cleanup clicked")

        # 1) Keys neu aus dem Dokument ableiten
        new_keys = resync_bibliography_keys_from_document(self.state)
        style = self.state.get("style", DEFAULT_STYLE)
        _debug(f"Cleanup: keys_after_prune={len(new_keys)} style={style}")

        # 2) APA/Harvard: a/b-Gruppen komplett neu normalisieren (inkl. Rollback)
        if style in ("apa", "harvard1"):
            renormalize_all_sig_groups()
            _debug("Cleanup: renormalize_all_sig_groups done")
            # nach Textänderungen erneut Keys ableiten
            new_keys = resync_bibliography_keys_from_document(self.state)

        # 3) Basis-Status setzen (sofort sichtbar)
        if not new_keys:
            base_status = "Bereinigt: keine Zitate mehr im Dokument gefunden."
        else:
            base_status = f"Bereinigt: {len(new_keys)} Zitat(e) im Dokument."

        self.set_status(base_status)

        # 4) Falls kein Bibliographie-Ziel: fertig
        if not has_bibliography_anchor():
            self.set_status(base_status + " (Kein Bibliographie-Ziel gesetzt.)")
            return

        # 5) Bibliographie asynchron aktualisieren
        try:
            cfg = get_cfg(allow_prompt=False)
        except RuntimeError:
            show_missing_zotero_config(self.root)
            return                  

        # UI sofort informieren
        self.set_status(base_status + " (Bibliographie wird aktualisiert...)")

        def _do_bib_update():
            update_bibliography(new_keys, style, cfg.api_key, cfg.library_id, cfg.library_type)
            self.root.after(
                0,
                lambda: self.set_status(base_status)
            )

        run_in_thread(
            _do_bib_update,
            on_error=lambda e: self.root.after(
                0, lambda: messagebox.showerror("Fehler", str(e), parent=self.root)
            )
        )
            
def main():
    root = tk.Tk()

    # Load config early in the UI thread.
    # If config exists (config.json / env), no dialog appears.
    # get_cfg(allow_prompt=True, parent=root)

    PickerApp(root)
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        print(f"Fatal: {ex}", file=sys.stderr)
        sys.exit(1)
