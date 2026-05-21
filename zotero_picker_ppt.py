import os
import re
import sys
import json
import argparse
import html
import uuid
import time
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

# ===================== Configuration =====================
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
STYLE_CODES = {code for _label, code in STYLE_CHOICES}
AUTHOR_YEAR_STYLES = {"apa", "harvard1", "chicago-author-date"}
UNKNOWN_STYLE = "unknown"
PREF_FONT_SIZE = 14
MIN_FONT_SIZE  = 10

# Bibliography anchor (robust)
ALT_BIB_PREFIX = "ZP_BIB_GUID="

# Tags for stable anchor detection
TAG_BIB_GUID_KEY = "ZP_BIB_GUID"

# JSON list per Shape: [{"key": "...", "cite": "(...)"}, ...]
CITE_TAG = "ZP_CITES"
# =========================================================

# ===================== COM stabilization (Option A+) =====================
COM_LOCK = threading.RLock()

LOG = logging.getLogger("zotero_ppt")

@contextlib.contextmanager
def com_context(action: str = "", *, use_lock: bool = True):
    """
    Option A+:
    - Calls CoInitialize/CoUninitialize per thread context
    - Optionally serializes COM access via COM_LOCK
    """
    pythoncom.CoInitialize()
    try:
        if use_lock:
            with COM_LOCK:
                if action:
                    LOG.debug("COM enter: %s", action)
                try:
                    yield
                finally:
                    if action:
                        LOG.debug("COM exit: %s", action)
        else:
            if action:
                LOG.debug("COM enter(no-lock): %s", action)
            try:
                yield
            finally:
                if action:
                    LOG.debug("COM exit(no-lock): %s", action)
    finally:
        pythoncom.CoUninitialize()
# ========================================================================

# ======== Zero-width markers for non-IEEE citations ========
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
        LOG.debug(msg)
        print(f"[ZP] {msg}")
    except Exception:
        pass

_CFG: Optional[ZoteroConfig] = None


def reset_cfg_cache():
    """Allows reloading the config later without restarting the process."""
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


# ===================== PowerPoint helpers =================
def _get_presentation():
    app = win32.Dispatch("PowerPoint.Application")
    pres = app.ActivePresentation
    if not pres:
        raise RuntimeError("Keine aktive Präsentation.")
    return pres


def _iter_slides(pres):
    """Yield slides defensively using indexed access first."""
    try:
        count = int(pres.Slides.Count)
    except Exception:
        count = 0

    if count:
        for idx in range(1, count + 1):
            try:
                yield pres.Slides.Item(idx)
            except Exception:
                continue
        return

    try:
        for slide in pres.Slides:
            yield slide
    except Exception:
        return


def _get_slide_by_id(pres, slide_id):
    """Return a usable PowerPoint slide by SlideID, or None if it cannot be resolved."""
    try:
        wanted = int(slide_id)
    except Exception:
        return None

    # Prefer indexed iteration. In CLI/Ribbon mode, FindBySlideID can return
    # a COM proxy that exposes the matching id but fails later on .Shapes.
    for slide in _iter_slides(pres):
        try:
            if int(slide.SlideID) != wanted:
                continue
        except Exception:
            continue

        try:
            _ = slide.Shapes
            return slide
        except Exception as e:
            _debug(f"Resolve slide by id: matched slide unusable slide_id={wanted}: {e}")
            continue

    # Fallback to PowerPoint's direct lookup, but validate that the returned
    # object is really usable as a slide.
    try:
        slide = pres.Slides.FindBySlideID(wanted)
    except Exception as e:
        _debug(f"Resolve slide by id: FindBySlideID failed slide_id={wanted}: {e}")
        return None

    try:
        _ = slide.Shapes
        return slide
    except Exception as e:
        _debug(f"Resolve slide by id: FindBySlideID returned unusable slide_id={wanted}: {e}")

    # Some COM versions may return a range-like object. Try Item(1) defensively.
    try:
        slide = slide.Item(1)
        _ = slide.Shapes
        return slide
    except Exception as e:
        _debug(f"Resolve slide by id: Item(1) fallback failed slide_id={wanted}: {e}")

    return None


def _iter_shape_collection(shapes):
    """Yield shapes from a PowerPoint Shapes collection defensively."""
    try:
        count = int(shapes.Count)
    except Exception:
        count = 0

    if count:
        for idx in range(1, count + 1):
            try:
                yield shapes.Item(idx)
            except Exception:
                continue
        return

    try:
        for shp in shapes:
            yield shp
    except Exception:
        return


def _shape_has_usable_text_frame(shp) -> bool:
    """Return True if the shape exposes a usable TextFrame."""
    try:
        if not getattr(shp, "HasTextFrame", False):
            return False
        _ = shp.TextFrame
        return True
    except Exception:
        return False


def _get_shape_id(shp) -> int:
    """Return a stable PowerPoint shape id, or -1 if it cannot be read."""
    if shp is None:
        return -1

    for attr in ("Id", "ID", "id"):
        try:
            shape_id = int(getattr(shp, attr))
            if shape_id > 0:
                return shape_id
        except Exception:
            continue

    return -1


def _iter_notes_shapes(slide):
    """Yield shapes from a slide's notes page, if available."""
    try:
        notes_page = slide.NotesPage
        shapes = notes_page.Shapes
    except Exception:
        return

    for shp in _iter_shape_collection(shapes):
        yield shp


def _iter_citation_shapes_for_slide(slide, include_notes: bool = True):
    """Yield citation-relevant text shapes for one slide, including notes if requested."""
    try:
        shapes = slide.Shapes
    except Exception as e:
        _debug(f"Citation shape scan: slide shapes unavailable: {e}")
        shapes = None

    if shapes is not None:
        for shp in _iter_shape_collection(shapes):
            if _shape_has_usable_text_frame(shp):
                yield "slide", slide, shp

    if include_notes:
        for shp in _iter_notes_shapes(slide):
            if _shape_has_usable_text_frame(shp):
                yield "notes", slide, shp


def iter_citation_shapes(include_notes: bool = True):
    """
    Yield citation-relevant text shapes in document order.

    Order:
    - normal slide shapes
    - notes page shapes for the same slide
    """
    pres = _get_presentation()

    for slide in _iter_slides(pres):
        yield from _iter_citation_shapes_for_slide(slide, include_notes=include_notes)


def _activate_powerpoint():
    with com_context("_activate_powerpoint"):
        try:
            app = win32.Dispatch("PowerPoint.Application")
            app.Activate()
            if app.ActiveWindow is not None:
                app.ActiveWindow.Activate()
        except Exception:
            pass


def _get_current_slide_and_shape():
    with com_context("_get_current_slide_and_shape"):
        app = win32.Dispatch("PowerPoint.Application")
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
                                _ = shp.TextFrame  # ensure that accessing the text frame does not crash
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
    Inserts text strictly at the actual cursor position.
    Returns the shape corresponding to the active text cursor.

    Notes pane fallback:
    PowerPoint may expose Selection.TextRange in notes without a usable ShapeRange.
    In that case, insert a temporary unique marker, find the shape containing it,
    remove the marker, and return that shape.
    """
    with com_context("ppt_insert_text_at_cursor"):
        app = win32.Dispatch("PowerPoint.Application")
        win = app.ActiveWindow
        if not win:
            raise RuntimeError("Kein PowerPoint-Fenster aktiv.")

        sel = win.Selection
        ppSelectionText = 3  # PowerPoint constant

        try:
            sel_type = sel.Type
        except Exception:
            sel_type = None

        if sel_type != ppSelectionText:
            raise RuntimeError(
                "Kein Textcursor gefunden.\n"
                "Bitte klicke in das Textfeld (Cursor sichtbar) und versuche es erneut."
            )

        try:
            tr = sel.TextRange
        except Exception:
            tr = None

        if tr is None:
            raise RuntimeError(
                "Kein Textcursor gefunden.\n"
                "Bitte klicke in das Textfeld (Cursor sichtbar) und versuche es erneut."
            )

        slide = None
        try:
            slide = sel.SlideRange(1)
        except Exception:
            try:
                slide = win.View.Slide
            except Exception:
                slide = None

        # First try the normal slide-pane path.
        shp = None
        try:
            sr = sel.ShapeRange
            if sr is not None and sr.Count >= 1:
                candidate = sr.Item(1)
                if candidate is not None and _shape_has_usable_text_frame(candidate):
                    shp = candidate
        except Exception:
            shp = None

        # Fallback: walk TextRange parent chain.
        if shp is None:
            try:
                candidate = tr.Parent
            except Exception:
                candidate = None

            for _ in range(6):
                if candidate is None:
                    break

                if _shape_has_usable_text_frame(candidate):
                    shp = candidate
                    break

                try:
                    candidate = candidate.Parent
                except Exception:
                    break

        # If we have a shape, use the old clean path.
        if shp is not None:
            tr.InsertAfter(s)
            return shp

        # Notes pane fallback:
        # Insert an invisible unique marker, find the actual text shape,
        # remove the marker again, and return the found shape.
        marker = f"⟦zp-insert:{uuid.uuid4().hex}⟧"
        inserted = s + marker

        try:
            tr.InsertAfter(inserted)
        except Exception:
            raise RuntimeError(
                "Der Text konnte nicht an der aktuellen Cursorposition eingefügt werden.\n"
                "Bitte klicke erneut in das Textfeld und versuche es noch einmal."
            )

        found_shape = None

        if slide is not None:
            search_shapes = _iter_citation_shapes_for_slide(slide, include_notes=True)
        else:
            _debug("Insert fallback: no active slide resolved, scanning all citation shapes")
            search_shapes = iter_citation_shapes(include_notes=True)

        for _scope, _slide, candidate in search_shapes:
            try:
                txt = candidate.TextFrame.TextRange.Text or ""
            except Exception:
                continue

            if marker not in txt:
                continue

            try:
                candidate.TextFrame.TextRange.Text = txt.replace(marker, "", 1)
            except Exception:
                pass

            found_shape = candidate
            break

        if found_shape is None:
            _debug("Insert fallback failed: temporary marker was not found in slide or notes shapes")
            raise RuntimeError(
                "Das Textfeld konnte nach dem Einfügen nicht eindeutig erkannt werden.\n"
                "Das Zitat wurde möglicherweise eingefügt, aber nicht im Citation-State gespeichert.\n"
                "Bitte klicke erneut in das Notizen-Textfeld und versuche es noch einmal."
            )

        return found_shape


def _copy_font(src_font, dst_font):
    """Copies as many font properties as possible in a robust way."""
    props = ["Name", "Size", "Bold", "Italic", "Underline", "Color", "BaselineOffset"]
    for p in props:
        try:
            setattr(dst_font, p, getattr(src_font, p))
        except Exception:
            pass


def ppt_insert_hidden_marker(marker_text: str, trailing_text: str = " "):
    """
    Inserts a marker at the cursor position, hides it, and ensures that
    subsequent typing continues with the previous formatting
    (font, size, etc.). Optionally inserts a normal space afterwards.
    """
    with com_context("ppt_insert_hidden_marker"):
        app = win32.Dispatch("PowerPoint.Application")
        win = app.ActiveWindow
        if not win:
            raise RuntimeError("Kein PowerPoint-Fenster aktiv.")
        sel = win.Selection

        # Get base range from text cursor or selected shape
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

        # Remember current typing format from selection text range, fallback to base_range
        try:
            fmt_font = sel.TextRange.Font
        except Exception:
            fmt_font = base_range.Font

        insert_start_abs = base_range.Start + base_range.Length
        to_insert = marker_text + (trailing_text or "")
        base_range.InsertAfter(to_insert)

        # Relative start position within base_range, 1-based
        rel_start_1b = (insert_start_abs - base_range.Start) + 1

        # Marker range
        mr = base_range.Characters(rel_start_1b, len(marker_text))

        # Marker: use surrounding font, then hide it
        try:
            _copy_font(fmt_font, mr.Font)
        except Exception:
            pass
        try:
            mr.Font.Hidden = True
        except Exception:
            pass

        # Trailing text, e.g. a space: visible and formatted like the surrounding text
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
    year = year or "n.d."   # important: otherwise the signature would only contain the author
    return f"{author}|{year}"


def collect_all_cites_by_key():
    with com_context("collect_all_cites_by_key"):
        by_key = {}
        for _scope, _slide, shp in iter_citation_shapes():
            try:
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
    Assigns a/b/... to all *different Zotero keys* with the same author-year signature.
    Also removes suffixes again when only one key remains after deletion.
    Updates:
    - visible text in all shapes
    - stored citation tags
    """
    with com_context(f"normalize_sig_group sig={sig}"):

        # 1) Collect all occurrences of this signature via tags
        occ = []  # (scope, slide, shp, idx, key, old_cite)
        for scope, slide, shp in iter_citation_shapes():
            try:
                arr = prune_cites_in_shape(shp)
                for i, c in enumerate(arr):
                    if c.get("sig") == sig:
                        k = c.get("key")
                        oc = c.get("cite") or ""
                        if k and oc:
                            occ.append((scope, slide, shp, i, k, oc))
            except Exception:
                continue

        if not occ:
            return

        # Stable order: first occurrence in the document
        keys_in_order = []
        for _, _, _, _, k, _ in occ:
            if k not in keys_in_order:
                keys_in_order.append(k)

        def strip_suffix(cite: str) -> str:
            # (Author, 2020a) -> (Author, 2020)
            # (Author, n.d.a) -> (Author, n.d.)
            return re.sub(r"((?:\d{4})|n\.d\.)[a-z]\)$", r"\1)", cite)

        # Remember one base citation without suffix per key
        base_by_key = {}
        for _, _, _, _, k, oc in occ:
            if k not in base_by_key:
                base_by_key[k] = strip_suffix(oc)

        # 2) Determine target citation text per key
        letters = "abcdefghijklmnopqrstuvwxyz"
        new_by_key = {}

        if len(keys_in_order) == 1:
            # Rollback case: remove a/b suffix
            k = keys_in_order[0]
            new_by_key[k] = base_by_key.get(k) or strip_suffix(occ[0][5])
        else:
            # Assign a/b/... suffixes
            for idx, k in enumerate(keys_in_order):
                base = base_by_key.get(k) or "(o. A.)"
                new_by_key[k] = base[:-1] + letters[idx] + ")"

        # 3) Replace sequentially per shape without using the shape object as a dict key
        by_shape = {}  # (scope, slide_id, shape_id) -> {"shape": shp, "items":[(idx, key, old_cite),...]}
        for scope, slide, shp, i, k, old_cite in occ:
            try:
                slide_id = int(slide.SlideID)
                shape_id = _get_shape_id(shp)
            except Exception:
                continue
            key = (scope, slide_id, shape_id)
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

                    # Update tag; index matches the tag list as long as prune ran beforehand
                    if i < len(arr) and arr[i].get("key") == k:
                        arr[i]["cite"] = new_cite

                if changed:
                    tr.Text = txt
                    _save_shape_cites(shp, arr)

            except Exception:
                continue


def renormalize_all_sig_groups():
    with com_context("renormalize_all_sig_groups"):
        sigs = []
        for _scope, _slide, shp in iter_citation_shapes():
            try:
                arr = prune_cites_in_shape(shp)
                for c in arr:
                    s = c.get("sig")
                    if s and s not in sigs:
                        sigs.append(s)
            except Exception:
                continue

        for s in sigs:
            normalize_sig_group(s)


HTTP_RETRY_STATUS = {429, 500, 502, 503, 504}
HTTP_TRANSIENT_EXC = (requests.Timeout, requests.ConnectionError)


def _retry_delay(resp, attempt: int) -> float:
    """
    Uses server-provided retry hints if available, otherwise applies simple backoff.
    """
    for header in ("Retry-After", "Backoff"):
        raw = (resp.headers.get(header) or "").strip()
        if raw:
            try:
                return max(0.0, float(raw))
            except ValueError:
                pass
    return min(0.5 * (2 ** attempt), 8.0)


def _safe_get(url, *, headers=None, params=None, timeout=HTTP_TIMEOUT, retries=3, context=""):
    """
    Robust GET with retry handling for transient network/API errors.

    Retries on:
    - Timeout / ConnectionError
    - HTTP 429 / 500 / 502 / 503 / 504

    Does not retry on:
    - 400 / 401 / 403 / 404 etc.
    """
    last_exc = None

    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=headers, params=params, timeout=timeout)
            _debug(
                f"HTTP GET status={r.status_code} "
                f"attempt={attempt + 1}/{retries + 1} "
                f"context={context} url={r.url}"
            )

            if r.status_code in HTTP_RETRY_STATUS:
                if attempt >= retries:
                    raise requests.HTTPError(
                        f"{r.status_code} HTTP error for url: {r.url}",
                        response=r
                    )

                delay = _retry_delay(r, attempt)
                _debug(
                    f"HTTP RETRY status={r.status_code} "
                    f"delay={delay:.1f}s "
                    f"attempt={attempt + 1}/{retries + 1} "
                    f"context={context}"
                )
                time.sleep(delay)
                continue

            r.raise_for_status()
            return r

        except HTTP_TRANSIENT_EXC as e:
            last_exc = e

            if attempt >= retries:
                _debug(
                    f"HTTP FAIL exc={type(e).__name__} "
                    f"attempt={attempt + 1}/{retries + 1} "
                    f"context={context}: {e}"
                )
                raise

            delay = min(0.5 * (2 ** attempt), 8.0)
            _debug(
                f"HTTP RETRY exc={type(e).__name__} "
                f"delay={delay:.1f}s "
                f"attempt={attempt + 1}/{retries + 1} "
                f"context={context}: {e}"
            )
            time.sleep(delay)

        except requests.RequestException as e:
            _debug(
                f"HTTP FAIL exc={type(e).__name__} "
                f"attempt={attempt + 1}/{retries + 1} "
                f"context={context}: {e}"
            )
            raise

    raise last_exc or RuntimeError(f"HTTP GET failed without specific exception: {url}")


class BibliographyFetchError(RuntimeError):
    """Bibliography entry could not be loaded via the Zotero Web API."""
    pass
# =========================================================


# ===================== Document state =====================
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
    with com_context("load_doc_state"):
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
    with com_context("save_doc_state"):
        pres = _get_presentation()
        props = pres.CustomDocumentProperties
        payload = json.dumps(state, ensure_ascii=False)
        p = _get_docprop_by_name(props, DOCPROP_NAME)
        if p is not None:
            p.Value = payload
            return
        props.Add(DOCPROP_NAME, False, 4, payload)


# ============ Bibliography: stable anchor via tags ========
def _get_shape_tag(shape, key):
    # Tags can be unreliable depending on the Office version, so read them defensively
    try:
        # VBA-style: shape.Tags("key")
        return shape.Tags(key)
    except Exception:
        try:
            # Alternative access pattern: shape.Tags.Item("key")
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


# LEGACY/OPTIONAL: zero-width markers, currently unused for APA/Harvard but may be useful later
def collect_all_cite_texts():
    with com_context("collect_all_cite_texts"):
        out = []
        for _scope, _slide, shp in iter_citation_shapes():
            try:
                for c in prune_cites_in_shape(shp):
                    t = (c.get("cite") or "").strip()
                    if t:
                        out.append(t)
            except Exception:
                continue
        return out


def prune_cites_in_shape(shp):
    """Remove stored cites that no longer have a visible occurrence in the shape text."""
    arr = _load_shape_cites(shp)
    if not arr:
        return []

    try:
        txt = shp.TextFrame.TextRange.Text or ""
    except Exception:
        return arr

    remaining_counts = {}
    for c in arr:
        cite = c.get("cite") or ""
        if cite:
            remaining_counts[cite] = txt.count(cite)

    kept = []
    for c in arr:
        cite = c.get("cite") or ""
        if not cite:
            continue
        if remaining_counts.get(cite, 0) > 0:
            kept.append(c)
            remaining_counts[cite] -= 1

    if kept != arr:
        _save_shape_cites(shp, kept)

    return kept


def _format_authoryear_base_from_item(item):
    """Base citation without a/b suffix: (Author, Year), derived from a pyzotero item dict."""
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


def _clean_cite_part(value, fallback=""):
    text = re.sub(r"\s+", " ", value or "").strip()
    return text or fallback


def _shorten_cite_part(value, max_len=60):
    text = _clean_cite_part(value, "")
    if len(text) > max_len:
        return text[: max_len - 3].rstrip() + "..."
    return text


def _mla_label_parts_from_item(item):
    """
    Return minimal MLA label parts.

    label:
    - author / corporate author / author pair / et al.
    - fallback to shortTitle/title when no creator exists

    qualifier:
    - shortTitle preferred
    - title fallback
    - only used for creator-based labels during duplicate normalization
    """
    data = item.get("data", {})
    creators = data.get("creators", []) or []

    names = []
    for c in creators:
        last = c.get("lastName") or c.get("name") or ""
        if last:
            names.append(last)

    title = _clean_cite_part(
        data.get("shortTitle") or data.get("title") or "",
        "",
    )

    if names:
        if len(names) == 1:
            label = names[0]
        elif len(names) == 2:
            label = f"{names[0]} and {names[1]}"
        else:
            label = f"{names[0]} et al."

        label = _shorten_cite_part(label, 60)
        qualifier = _shorten_cite_part(title, 60) if title else ""

        if qualifier.casefold() == label.casefold():
            qualifier = ""

        return label, qualifier

    label = _shorten_cite_part(title or "o. T.", 60)
    return label, ""


def _format_mla_cite_from_parts(label, qualifier=""):
    label = _shorten_cite_part(label or "o. T.", 60)
    qualifier = _shorten_cite_part(qualifier or "", 60)

    if qualifier and qualifier.casefold() != label.casefold():
        return f"({label}, {qualifier})"

    return f"({label})"


def _format_mla_base_from_item(item):
    """Return a minimal MLA-plausible parenthetical citation without locator support."""
    label, _qualifier = _mla_label_parts_from_item(item)
    return _format_mla_cite_from_parts(label)


# LEGACY/OPTIONAL: zero-width markers, currently unused for APA/Harvard but may be useful later
def _disambiguate_authoryear(base, existing_cites):
    """Turns (Smith, 2020) into (Smith, 2020a/b/...) if needed."""
    if base not in existing_cites:
        return base
    if not base.endswith(")"):
        return base
    stem = base[:-1]  # without closing parenthesis
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


def _mla_label_from_cite_text(cite):
    cite = (cite or "").strip()
    if not (cite.startswith("(") and cite.endswith(")")):
        return ""

    inner = cite[1:-1].strip()
    if not inner:
        return ""

    # Legacy fallback only. New records store mla_label explicitly.
    return inner.split(",", 1)[0].strip()


def _is_probable_legacy_mla_record(c):
    style = c.get("style")
    if style == "mla":
        return True
    if style:
        return False

    cite = (c.get("cite") or "").strip()
    if not (cite.startswith("(") and cite.endswith(")")):
        return False

    # Avoid treating old APA/Harvard/Chicago author-date records as MLA.
    if re.search(r",\s*(?:\d{4}|n\.d\.)[a-z]?\)$", cite):
        return False

    return True




def infer_cite_record_style(cite_record, fallback_style=None):
    """Infer a citation record style without mutating legacy metadata."""
    if not isinstance(cite_record, dict):
        return UNKNOWN_STYLE

    explicit = (cite_record.get("style") or "").strip()
    if explicit in STYLE_CODES:
        return explicit
    if explicit:
        return UNKNOWN_STYLE

    cite = (cite_record.get("cite") or "").strip()
    if not cite:
        return UNKNOWN_STYLE

    if IEEE_CITE_RE.match(cite) or PH_RE.match(cite):
        return "ieee"

    if _is_probable_legacy_mla_record(cite_record):
        return "mla"

    fallback = fallback_style if fallback_style in STYLE_CODES else None
    if cite_record.get("sig") and fallback in AUTHOR_YEAR_STYLES:
        return fallback

    if re.search(r",\s*(?:\d{4}|n\.d\.)[a-z]?\)$", cite):
        if fallback in AUTHOR_YEAR_STYLES:
            return fallback
        return UNKNOWN_STYLE

    return UNKNOWN_STYLE


def collect_document_citation_styles(fallback_style=None):
    """Return distinct citation styles found in active slide and notes cite records."""
    with com_context("collect_document_citation_styles"):
        styles = []
        for _scope, _slide, shp in iter_citation_shapes():
            try:
                for c in prune_cites_in_shape(shp):
                    style = infer_cite_record_style(c, fallback_style=fallback_style)
                    if style not in styles:
                        styles.append(style)
            except Exception:
                continue
        return styles


def get_existing_document_style(fallback_style=None):
    """Return the single known document style, or None for empty/mixed/unknown state."""
    styles = collect_document_citation_styles(fallback_style=fallback_style)
    known = [s for s in styles if s != UNKNOWN_STYLE]
    if len(known) == 1 and len(styles) == 1:
        return known[0]
    return None


def find_mixed_style_conflicts(target_style, fallback_style=None):
    """
    Detect whether target_style would create or preserve an unsupported
    mixed/unknown citation-style state.
    """
    styles = collect_document_citation_styles(fallback_style=fallback_style)
    if not styles:
        return {
            "blocked": False,
            "reason": "empty",
            "styles": [],
            "document_style": None,
        }

    known = [s for s in styles if s != UNKNOWN_STYLE]
    has_unknown = UNKNOWN_STYLE in styles

    if has_unknown or len(known) != 1:
        return {
            "blocked": True,
            "reason": "mixed-or-unknown",
            "styles": styles,
            "document_style": None,
        }

    document_style = known[0]
    return {
        "blocked": document_style != target_style,
        "reason": "different-style" if document_style != target_style else "same-style",
        "styles": styles,
        "document_style": document_style,
    }


def _document_style_description(styles, document_style=None):
    if document_style:
        return f"im Stil „{code_to_label(document_style)}“"
    labels = [code_to_label(s) if s != UNKNOWN_STYLE else "unbekannt" for s in styles]
    if labels:
        return "mit mehreren oder nicht eindeutig erkennbaren Stilen (" + ", ".join(labels) + ")"
    return "mit nicht eindeutig erkennbarem Zitierstil"


def _show_message(parent, kind, title, message):
    show = messagebox.showwarning if kind == "warning" else messagebox.showerror
    if parent is not None:
        parent.after(0, lambda: show(title, message, parent=parent))
    else:
        show(title, message)


def block_mixed_style_insert_or_show_message(parent, target_style, fallback_style=None):
    conflict = find_mixed_style_conflicts(target_style, fallback_style=fallback_style)
    if not conflict["blocked"]:
        return False

    LOG.warning(
        "Mixed citation style blocked: action=insert target_style=%s document_styles=%s reason=%s",
        target_style,
        conflict["styles"],
        conflict["reason"],
    )

    document_style = conflict.get("document_style")
    if document_style:
        msg = (
            "Das Zitat wurde nicht eingefügt.\n\n"
            f"In dieser Präsentation sind bereits Zitate im Stil „{code_to_label(document_style)}“ vorhanden. "
            f"Der aktuell gewählte Stil ist „{code_to_label(target_style)}“.\n\n"
            "Unterschiedliche Zitierstile innerhalb einer Präsentation werden nicht unterstützt.\n\n"
            "Bitte verwende den bestehenden Stil oder stelle die Präsentation über den Zitierstil-Wechsel vollständig auf den neuen Stil um."
        )
    else:
        msg = (
            "Das Zitat wurde nicht eingefügt.\n\n"
            "In dieser Präsentation sind bereits Zitate mit mehreren oder nicht eindeutig erkennbaren Zitierstilen vorhanden.\n\n"
            "Unterschiedliche Zitierstile innerhalb einer Präsentation werden nicht unterstützt.\n\n"
            "Bitte stelle die Präsentation über den Zitierstil-Wechsel vollständig auf einen einheitlichen Stil um."
        )

    _show_message(parent, "warning", "Zitierstil-Konflikt", msg)
    return True


def normalize_mla_duplicate_labels():
    """
    MLA:
    Normalize duplicate visible labels across different Zotero keys.

    This is intentionally metadata-based:
    - no Zotero Web API calls
    - no CSL/style-engine refactor
    - no locator/page support

    Updates both:
    - visible citation text
    - ZP_CITES records
    """
    with com_context("normalize_mla_duplicate_labels"):
        occ = []  # (scope, slide, shp, idx, key, old_cite, label, qualifier)

        for scope, slide, shp in iter_citation_shapes():
            try:
                arr = prune_cites_in_shape(shp)
                for i, c in enumerate(arr):
                    if not _is_probable_legacy_mla_record(c):
                        continue

                    key = c.get("key")
                    old_cite = c.get("cite") or ""
                    label = c.get("mla_label") or _mla_label_from_cite_text(old_cite)
                    qualifier = c.get("mla_qualifier") or ""

                    if key and old_cite and label:
                        occ.append(
                            (scope, slide, shp, i, key, old_cite, label, qualifier)
                        )
            except Exception:
                continue

        if not occ:
            return

        labels_in_order = []
        keys_by_label = {}
        qualifier_by_label_key = {}

        for _scope, _slide, _shp, _idx, key, _old_cite, label, qualifier in occ:
            if label not in keys_by_label:
                keys_by_label[label] = []
                labels_in_order.append(label)

            if key not in keys_by_label[label]:
                keys_by_label[label].append(key)

            if qualifier and (label, key) not in qualifier_by_label_key:
                qualifier_by_label_key[(label, key)] = qualifier

        new_by_label_key = {}

        for label in labels_in_order:
            keys = keys_by_label[label]

            if len(keys) == 1:
                key = keys[0]
                new_by_label_key[(label, key)] = _format_mla_cite_from_parts(label)
                continue

            for key in keys:
                qualifier = qualifier_by_label_key.get((label, key), "")
                new_by_label_key[(label, key)] = _format_mla_cite_from_parts(
                    label,
                    qualifier,
                )

        by_shape = {}
        for scope, slide, shp, idx, key, old_cite, label, qualifier in occ:
            try:
                slide_id = int(slide.SlideID)
                shape_id = _get_shape_id(shp)
            except Exception:
                continue

            shape_key = (scope, slide_id, shape_id)
            by_shape.setdefault(shape_key, {"shape": shp, "items": []})
            by_shape[shape_key]["items"].append(
                (idx, key, old_cite, label, qualifier)
            )

        for _shape_key, pack in by_shape.items():
            shp = pack["shape"]
            items = pack["items"]

            try:
                tr = shp.TextFrame.TextRange
                txt = tr.Text or ""
                arr = _load_shape_cites(shp)

                text_changed = False
                tags_changed = False

                for idx, key, old_cite, label, qualifier in sorted(
                    items,
                    key=lambda x: x[0],
                ):
                    new_cite = new_by_label_key.get((label, key), old_cite)

                    if new_cite != old_cite:
                        txt, replaced = _replace_first(txt, old_cite, new_cite)
                        if replaced:
                            text_changed = True

                    if idx < len(arr) and arr[idx].get("key") == key:
                        if arr[idx].get("cite") != new_cite:
                            arr[idx]["cite"] = new_cite
                            tags_changed = True

                        if arr[idx].get("style") != "mla":
                            arr[idx]["style"] = "mla"
                            tags_changed = True

                        if not arr[idx].get("mla_label"):
                            arr[idx]["mla_label"] = label
                            tags_changed = True

                        if qualifier and not arr[idx].get("mla_qualifier"):
                            arr[idx]["mla_qualifier"] = qualifier
                            tags_changed = True

                if text_changed:
                    tr.Text = txt

                if text_changed or tags_changed:
                    _save_shape_cites(shp, arr)

            except Exception:
                continue


def set_bibliography_anchor_from_selection():
    with com_context("set_bibliography_anchor_from_selection"):
        app = win32.Dispatch("PowerPoint.Application")
        win = app.ActiveWindow
        if not win:
            raise RuntimeError("Kein PowerPoint-Fenster aktiv.")

        sel = win.Selection
        slide = None
        shp = None

        ppSelectionShapes = 2
        ppSelectionText = 3

        try:
            slide = sel.SlideRange(1)
        except Exception:
            try:
                slide = win.View.Slide
            except Exception:
                slide = None

        try:
            sel_type = sel.Type
        except Exception:
            sel_type = None

        if sel_type == ppSelectionShapes:
            try:
                sr = sel.ShapeRange
                if sr is not None and sr.Count >= 1:
                    shp = sr.Item(1)
            except Exception:
                shp = None

        elif sel_type == ppSelectionText:
            try:
                sr = sel.ShapeRange
                if sr is not None and sr.Count >= 1:
                    shp = sr.Item(1)
            except Exception:
                shp = None

        _debug(
            f"Anchor selection: slide={getattr(slide,'SlideID',None)}, "
            f"shape_id={_get_shape_id(shp)}, "
            f"hasTextFrame={getattr(shp,'HasTextFrame',False)}"
        )

        if slide is None:
            raise RuntimeError("Keine aktive Folie gefunden.")

        if shp is None or not getattr(shp, "HasTextFrame", False):
            raise RuntimeError(
                "Bitte wähle ein Textfeld aus.\n"
                "Hinweis: Textfeld anklicken (Rahmen sichtbar), "
                "nicht nur den Cursor setzen."
            )

        # Intentionally no `shp.TextFrame.HasText`: bibliography fields are allowed to be empty.
        # Capture PowerPoint IDs immediately while the selected COM objects are still connected.
        try:
            selected_slide_id = int(slide.SlideID)
            selected_shape_id = _get_shape_id(shp)
        except Exception as e:
            _debug(f"Anchor save: failed to read selected object IDs: {e}")
            selected_slide_id = -1
            selected_shape_id = -1

        st = load_doc_state()
        bib_guid = st.get("bib_guid") or str(uuid.uuid4())
        st["bib_guid"] = bib_guid

        # Hard anchor via captured IDs. This remains usable even if the original
        # selection COM proxy becomes disconnected before tagging.
        if selected_slide_id > 0 and selected_shape_id > 0:
            st["bib_anchor"] = {
                "slide_id": selected_slide_id,
                "shape_id": selected_shape_id,
            }
        else:
            _debug("Anchor save: selected shape has no stable SlideID/ShapeID")
            st.pop("bib_anchor", None)

        save_doc_state(st)
        _debug(f"Anchor saved: bib_guid={st.get('bib_guid')}, bib_anchor={st.get('bib_anchor')}")

        # 1) Try tags as a nice-to-have
        _set_shape_tag(shp, TAG_BIB_GUID_KEY, bib_guid)

        # 2) Robust path: set AlternativeText, which persists reliably
        try:
            shp.AlternativeText = ALT_BIB_PREFIX + bib_guid
        except Exception:
            # Some shapes may block this; in that case, keep at least the tag
            pass
        _debug("Anchor tagged: tag + AlternativeText set where possible")

        # 3) Immediately verify that the target can be resolved again
        if not has_bibliography_anchor():
            _debug("Anchor check FAIL: _resolve_anchor_list() returned 0")
            raise RuntimeError(
                "Bibliographie-Ziel konnte nicht gespeichert werden (PowerPoint hat den Anker nicht übernommen).\n"
                "Bitte: Textfeld einmal anklicken (Rahmen sichtbar), dann erneut „Bibliographie-Ziel…“."
            )


def _resolve_anchor_list():
    with com_context("_resolve_anchor_list"):
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

        # 1) Direct lookup via SlideID/ShapeID, most reliable path
        slide_id = anch.get("slide_id")
        shape_id = anch.get("shape_id")
        if slide_id and shape_id:
            try:
                slide_id = int(slide_id)
                shape_id = int(shape_id)
                slide = _get_slide_by_id(pres, slide_id)

                if slide is not None:
                    shapes = None

                    try:
                        shapes = slide.Shapes
                    except Exception as e:
                        _debug(f"Resolve anchors: direct slide shapes unavailable: {e}")

                        # PowerPoint COM may return a SlideRange-like object for FindBySlideID.
                        # In that case, the real slide is often Item(1).
                        try:
                            slide = slide.Item(1)
                            shapes = slide.Shapes
                            _debug("Resolve anchors: direct slide normalized via Item(1)")
                        except Exception as e2:
                            _debug(f"Resolve anchors: direct slide Item(1) fallback failed: {e2}")
                            shapes = None

                    if shapes is not None:
                        for shp in _iter_shape_collection(shapes):
                            sid = _get_shape_id(shp)
                            if sid == shape_id and getattr(shp, "HasTextFrame", False):
                                key = (int(slide.SlideID), sid)
                                if key not in seen:
                                    resolved.append((slide, shp))
                                    seen.add(key)
                                break
            except Exception as e:
                _debug(f"Resolve anchors: direct lookup failed: {e}")

        # 2) Additionally find all shapes with the GUID, including continuation slides or duplicates
        if bib_guid:
            for slide in _iter_slides(pres):
                try:
                    shapes = slide.Shapes
                except Exception as e:
                    _debug(f"Resolve anchors: slide shapes unavailable: {e}")
                    continue

                for shp in _iter_shape_collection(shapes):
                    try:
                        if not getattr(shp, "HasTextFrame", False):
                            continue

                        sid = _get_shape_id(shp)
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

        _debug(f"Resolve anchors: found={len(resolved)}")
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

    return f"Stil: {code_to_label(style)} | Zitate: {len(keys)} | Bibliographie-Ziel: {anchor_txt}"


def _is_title_placeholder(shp) -> bool:
    """Return True if the shape is very likely a title placeholder."""
    try:
        t = int(getattr(shp.PlaceholderFormat, "Type", -1))
        if t in (1, 3):  # Title + Center Title
            return True
    except Exception:
        pass

    # Additional fallback heuristic
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
    Finds a suitable text placeholder on a slide, preferring the same
    placeholder type as src_shape if src_shape is a placeholder.
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
            # Is this a placeholder object?
            _ = shp.PlaceholderFormat  # raises an exception if this is not a placeholder
        except Exception:
            continue

        try:
            # Skip title placeholders
            if _is_title_placeholder(shp):
                continue
        except Exception:
            pass

        # Prefer the same placeholder type as the source shape
        score = 0
        try:
            ptype = _placeholder_type(shp)
            if ptype not in allowed:
                continue

            if want_type is not None and ptype == want_type:
                score += 1000
        except Exception:
            continue

        # Secondary heuristic: largest text field
        try:
            score += int(float(shp.Width) * float(shp.Height))
        except Exception:
            score += 0

        if score > best_score:
            best = shp
            best_score = score

    if best is None:
        _debug("Find placeholder: none found (allowed types: {2,7})")

    return best


def _get_slide_title_text(slide):
    """Reads the text of a slide title placeholder."""
    try:
        for shp in slide.Shapes:
            try:
                if int(shp.PlaceholderFormat.Type) in (1, 3):  # Title or Center Title
                    if shp.TextFrame.HasText:
                        return shp.TextFrame.TextRange.Text
            except Exception:
                continue
    except Exception:
        pass
    return ""


def _set_slide_title_text(slide, text):
    """Sets the text of a slide title placeholder."""
    if not text:
        return
    try:
        for shp in slide.Shapes:
            try:
                if int(shp.PlaceholderFormat.Type) in (1, 3):  # Title or Center Title
                    shp.TextFrame.TextRange.Text = text
                    return
            except Exception:
                continue
    except Exception:
        pass


def _duplicate_anchor_to_new_slide_like(src_slide, src_shape):
    with com_context("_duplicate_anchor_to_new_slide_like"):
        pres = _get_presentation()

        # Remember the title of the source slide
        title_text = _get_slide_title_text(src_slide)

        # Create a new slide with the same layout
        new_slide = pres.Slides.AddSlide(pres.Slides.Count + 1, src_slide.CustomLayout)
        _debug(f"Created new bibliography slide (layout: {src_slide.CustomLayout.Name})")

        # Copy title
        if title_text:
            _set_slide_title_text(new_slide, title_text)
            _debug(f"Copied title: '{title_text}'")

        # 1) Try to find a suitable layout text placeholder
        new_shape = _find_best_text_placeholder(new_slide, src_shape=src_shape)

        if new_shape is not None:
            try:
                ptype = int(new_shape.PlaceholderFormat.Type)
            except Exception:
                ptype = "?"

            _debug(f"Found layout placeholder (Type={ptype}, Name='{getattr(new_shape, 'Name', '')}')")

            try:
                new_shape.TextFrame.TextRange.Text = ""
            except Exception:
                pass
        else:
            _debug("WARNING: no suitable text placeholder found in layout; using copy/paste fallback")

            # Fallback
            src_shape.Copy()
            pasted = new_slide.Shapes.Paste()
            new_shape = pasted.Item(1)
            try:
                new_shape.TextFrame.TextRange.Text = ""
            except Exception:
                pass

        # Tag with the same GUID; continuation slides belong to the same bibliography set
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


# ============ Zotero: bibliography entries ===============
def html_to_text(html_str):
    text = re.sub(r"<br\s*/?>", "\n", html_str, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    return re.sub(r"\s+\n", "\n", text).strip()


def _strip_ieee_bibliography_label(entry: str) -> str:
    """Remove numbering returned by Zotero for single IEEE bibliography entries."""
    return re.sub(r"^\s*(?:\[\d+\]|\d+\.)\s*", "", entry or "").strip()


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

    last_error = None

    for url, params in candidates:
        try:
            r = _safe_get(
                url,
                headers=headers,
                params=params,
                timeout=HTTP_TIMEOUT,
                retries=3,
                context=f"bib-html key={item_key}"
            )
        except requests.RequestException as e:
            last_error = f"{type(e).__name__}: {e}"
            continue

        ct = (r.headers.get("Content-Type", "") or "").lower()
        txt = r.text or ""

        looks_like_json = txt.lstrip().startswith("{") or txt.lstrip().startswith("[")
        if txt.strip() and (("text/html" in ct) or ("text/x-bibliography" in ct)) and not looks_like_json:
            return html_to_text(txt)

        last_error = f"unexpected bibliography response content-type={ct!r}"

    try:
        raw = _safe_get(
            f"{base}/items/{item_key}",
            headers={
                "Zotero-API-Key": api_key,
                "Zotero-API-Version": "3",
                "Accept": "application/json",
            },
            timeout=HTTP_TIMEOUT,
            retries=3,
            context=f"bib-json key={item_key}"
        )
    except requests.RequestException as e:
        raise BibliographyFetchError(
            f"Bibliographie für {item_key} konnte nicht geladen werden ({last_error or f'{type(e).__name__}: {e}'})"
        ) from e

    try:
        data = raw.json()
    except ValueError as e:
        raise BibliographyFetchError(
            f"Bibliographie für {item_key} konnte nicht geladen werden (invalid JSON response)"
        ) from e

    if isinstance(data, list) and data:
        data = data[0]

    if not isinstance(data, dict):
        raise BibliographyFetchError(
            f"Bibliographie für {item_key} konnte nicht geladen werden (unexpected JSON shape)"
        )

    d = data.get("data", {})
    title = d.get("title") or "[o. T.]"
    creators = d.get("creators") or []
    first = (creators[0].get("lastName") or creators[0].get("name")) if creators else ""
    m = re.search(r"(\d{4})", d.get("date") or "")
    year = m.group(1) if m else ""
    url = d.get("url") or ""
    pieces = [p for p in [first, f"({year})" if year else "", title, url] if p]

    if not pieces:
        raise BibliographyFetchError(
            f"Bibliographie für {item_key} konnte nicht geladen werden (empty fallback data)"
        )

    return " ".join(pieces)
# =========================================================


# ======== Bibliography: writing and pagination ===========
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
        try:
            tr.ParagraphFormat.Bullet.Visible = False
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
        _debug("Fit WARN: only 1 entry fits at minimum size")
        return 1, min_size

    return 0, preferred_size


def update_bibliography(keys, style, api_key, library_id, library_type, numbering=None):
    _debug(f"Bib update: keys={len(keys)} style={style}")

    # 1) Resolve anchors only for validation/clearing.
    #    Do not introduce a second anchor representation here:
    #    _resolve_anchor_list() is the single source of truth.
    with com_context("update_bibliography.resolve_anchors"):
        anchors = _resolve_anchor_list()
        _debug(f"Bib update: anchors={len(anchors)}")

        if not anchors:
            if keys:
                raise RuntimeError("Kein Bibliographie-Ziel gesetzt oder Ziel nicht auflösbar.")
            return

        # If there are no keys, clear all bibliography anchor fields and stop.
        if not keys:
            cleared = 0
            for _slide, shape in anchors:
                try:
                    if _shape_has_usable_text_frame(shape):
                        shape.TextFrame.TextRange.Text = ""
                        cleared += 1
                except Exception as e:
                    _debug(f"Bib clear: anchor skipped: {e}")

            _debug(f"Bib cleared: anchors={cleared} style={style}")
            return

    if style == "ieee" and numbering is None:
        numbering = build_ieee_numbering_from_document()
        key_set = set(keys)
        ordered_keys = [k for k in numbering.keys() if k in key_set]
        if ordered_keys:
            keys = ordered_keys
        _debug(f"Bib update IEEE numbering: keys={len(keys)} numbering={len(numbering)}")

    # 2) Build entries via Web API.
    entries = []
    failures = []

    for k in keys:
        try:
            entry = get_bibliography_entry_webapi(
                api_key,
                library_id,
                library_type,
                k,
                style,
            )
            if style == "ieee":
                entry = _strip_ieee_bibliography_label(entry)

            if numbering and k in numbering:
                entries.append(f"[{numbering[k]}] {entry}")
            else:
                entries.append(entry)

        except BibliographyFetchError as e:
            failures.append((k, str(e)))

    if failures:
        preview = "; ".join(f"{k}: {msg}" for k, msg in failures[:3])
        _debug(f"Bib update ABORT: failures={len(failures)} | {preview}")
        raise RuntimeError(
            f"Bibliographie nicht aktualisiert: {len(failures)} Einträge konnten nicht geladen werden. "
            f"Erste Fehler: {preview}"
        )

    remaining = entries[:]
    _debug(f"Bib entries generated: {len(entries)}")

    # 3) Write to the anchors that were resolved before the Web API calls.
    #    In CLI/Ribbon mode, a fresh lookup after network calls can fail even
    #    though the original anchor is still usable. Therefore, keep the
    #    originally resolved anchor as the primary path and only use a fresh
    #    lookup as a fallback.
    with com_context("update_bibliography.write_and_paginate"):
        _debug(f"Bib write: cached_anchors={len(anchors)}")

        usable_anchors = []
        for slide, shape in anchors:
            try:
                if _shape_has_usable_text_frame(shape):
                    usable_anchors.append((slide, shape))
                else:
                    _debug(
                        f"Bib write: cached anchor skipped, no usable TextFrame "
                        f"slide={getattr(slide, 'SlideID', None)} "
                        f"shape_id={_get_shape_id(shape)}"
                    )
            except Exception as e:
                _debug(f"Bib write: cached anchor skipped: {e}")

        if not usable_anchors:
            _debug("Bib write: cached anchors unavailable; trying fresh lookup")
            fresh_anchors = _resolve_anchor_list()
            _debug(f"Bib write: fresh_anchors={len(fresh_anchors)}")

            for slide, shape in fresh_anchors:
                try:
                    if _shape_has_usable_text_frame(shape):
                        usable_anchors.append((slide, shape))
                    else:
                        _debug(
                            f"Bib write: fresh anchor skipped, no usable TextFrame "
                            f"slide={getattr(slide, 'SlideID', None)} "
                            f"shape_id={_get_shape_id(shape)}"
                        )
                except Exception as e:
                    _debug(f"Bib write: fresh anchor skipped: {e}")

        anchors = usable_anchors

        if not anchors:
            _debug("Bib write ABORT: no usable bibliography text field")
            raise RuntimeError(
                "Bibliographie-Ziel konnte beim Schreiben nicht verwendet werden. "
                "Bitte nutze „Bibliographie neu schreiben“ erneut oder setze das Ziel neu."
            )

        for slide, shape in anchors:
            if remaining:
                used, _ = _try_fit_entries_into_shape(shape, remaining)

                if used <= 0:
                    _debug("Bib write ABORT: no entries could be placed on existing anchor")
                    raise RuntimeError(
                        f"Bibliographie nicht vollständig geschrieben: {len(remaining)} Einträge konnten nicht platziert werden."
                    )

                # Verify that PowerPoint actually accepted text in the target.
                try:
                    written_text = shape.TextFrame.TextRange.Text or ""
                    if not written_text.strip():
                        raise RuntimeError(
                            "Bibliographie-Ziel wurde beschrieben, enthält danach aber keinen Text."
                        )
                except RuntimeError:
                    raise
                except Exception as e:
                    _debug(f"Bib write: read-back verification unavailable: {e}")

                _debug(
                    f"Bib write: placed={used} "
                    f"remaining={max(0, len(remaining) - used)} "
                    f"slide={getattr(slide, 'SlideID', None)} "
                    f"shape_id={_get_shape_id(shape)}"
                )

                remaining = remaining[used:]

            else:
                try:
                    shape.TextFrame.TextRange.Text = ""
                except Exception:
                    pass

        while remaining:
            src_slide, src_shape = anchors[-1]
            new_slide, new_shape = _duplicate_anchor_to_new_slide_like(src_slide, src_shape)

            if new_shape is None or not _shape_has_usable_text_frame(new_shape):
                _debug("Bib pagination ABORT: new_shape invalid or without TextFrame")
                raise RuntimeError(
                    f"Bibliographie nicht vollständig geschrieben: {len(remaining)} Einträge konnten nicht platziert werden."
                )

            anchors.append((new_slide, new_shape))
            used, _ = _try_fit_entries_into_shape(new_shape, remaining)

            if used <= 0:
                _debug("Bib pagination ABORT: no entries could be placed on new anchor")
                raise RuntimeError(
                    f"Bibliographie nicht vollständig geschrieben: {len(remaining)} Einträge konnten nicht platziert werden."
                )

            _debug(
                f"Bib pagination: placed={used} "
                f"remaining={max(0, len(remaining) - used)} "
                f"slide={getattr(new_slide, 'SlideID', None)} "
                f"shape_id={_get_shape_id(new_shape)}"
            )

            remaining = remaining[used:]

        _debug(f"Bib update OK: entries={len(entries)} anchors={len(anchors)} style={style}")
# =========================================================


# =========== IEEE: placeholders and renumbering ==========
PH_RE = re.compile(r"⟦zp:([A-Za-z0-9]+)⟧")

def scan_all_placeholders():
    with com_context("scan_all_placeholders"):
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

IEEE_CITE_RE = re.compile(r"^\[\d+\]$")


def _is_ieee_cite_record(c):
    return c.get("style") == "ieee" or IEEE_CITE_RE.match(c.get("cite") or "")


def _collect_ieee_cites_in_shape(shp, arr=None):
    """
    Return IEEE cite records in visible text order for one shape.

    This is required because Shape Tags preserve insertion/appending order,
    not the actual text position after a user inserts a citation before
    existing IEEE citations.
    """
    if arr is None:
        arr = prune_cites_in_shape(shp)

    try:
        txt = shp.TextFrame.TextRange.Text or ""
    except Exception:
        return "", [], []

    next_start_by_token = {}
    positioned = []
    fallback = []

    for idx, c in enumerate(arr):
        if not _is_ieee_cite_record(c):
            continue

        tokens = []
        cite = c.get("cite") or ""
        key = c.get("key") or ""

        if cite:
            tokens.append(cite)
        if key:
            tokens.append(f"⟦zp:{key}⟧")

        # Preserve order while removing duplicate candidate tokens.
        tokens = list(dict.fromkeys(tokens))

        best = None
        for token in tokens:
            start_from = next_start_by_token.get(token, 0)
            pos = txt.find(token, start_from)
            if pos < 0:
                continue

            candidate = (pos, pos + len(token), token)
            if best is None or candidate[0] < best[0]:
                best = candidate

        if best is None:
            fallback.append({
                "idx": idx,
                "record": c,
                "start": None,
                "end": None,
                "old_text": "",
            })
            continue

        start, end, old_text = best
        next_start_by_token[old_text] = end

        positioned.append({
            "idx": idx,
            "record": c,
            "start": start,
            "end": end,
            "old_text": old_text,
        })

    positioned.sort(key=lambda hit: (hit["start"], hit["idx"]))
    return txt, positioned, fallback


def build_ieee_numbering_from_document():
    """Build stable IEEE numbering from stored cite tags in document order."""
    with com_context("build_ieee_numbering_from_document"):
        numbering = {}
        n = 1

        for _scope, _slide, shp in iter_citation_shapes():
            try:
                arr = prune_cites_in_shape(shp)
                _, positioned, fallback = _collect_ieee_cites_in_shape(shp, arr)

                for hit in positioned + fallback:
                    c = hit["record"]
                    key = c.get("key")
                    if key and key not in numbering:
                        numbering[key] = n
                        n += 1
            except Exception:
                continue

        return numbering


def resync_bibliography_keys_from_document(state=None):
    with com_context("resync_bibliography_keys_from_document"):
        keys = []

        for _scope, _slide, shp in iter_citation_shapes():
            try:
                kept = prune_cites_in_shape(shp)
                for c in kept:
                    k = c.get("key")
                    if k and k not in keys:
                        keys.append(k)
            except Exception:
                continue

        if state is None:
            state = load_doc_state()

        state["bib_keys"] = keys
        save_doc_state(state)
        return keys


def insert_ieee_placeholder(key, *, parent=None) -> bool:
    """
    IEEE: insert placeholder, persist citation metadata, renumber,
    and update bibliography if possible.
    Runs asynchronously in a worker so the UI does not block.
    """
    def _work():
        state = load_doc_state()
        style_local = state.get("style", DEFAULT_STYLE) or DEFAULT_STYLE
        if block_mixed_style_insert_or_show_message(parent, "ieee", fallback_style=style_local):
            if parent is not None:
                parent.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Zotero PowerPoint",
                        "Abgebrochen: anderer Zitierstil im Dokument erkannt.",
                        parent=parent,
                    ),
                )
            return

        placeholder = f"⟦zp:{key}⟧"
        shp = ppt_insert_text_at_cursor(f" {placeholder}")

        arr = _load_shape_cites(shp)
        arr.append({
            "key": key,
            "cite": placeholder,
            "style": "ieee",
        })
        _save_shape_cites(shp, arr)

        ok = renumber_ieee_and_update(parent=parent)

        if parent is not None:
            parent.after(
                0,
                lambda: messagebox.showinfo(
                    "IEEE",
                    "IEEE-Platzhalter eingefügt und renummeriert."
                    if ok else
                    "IEEE-Platzhalter eingefügt. Zotero konfigurieren, um Renummerierung/Bibliographie zu aktualisieren.",
                    parent=parent
                )
            )

    run_in_thread("IEEEInsert", _work, ui_parent=parent)
    return True


def renumber_ieee_document_citations():
    """Renumber IEEE citations in visible text and ZP_CITES without writing bibliography."""
    with com_context("renumber_ieee_document_citations"):
        state = load_doc_state()
        numbering = build_ieee_numbering_from_document()

        for _scope, _slide, shp in iter_citation_shapes():
            try:
                if not getattr(shp, "HasTextFrame", False):
                    continue

                tr = shp.TextFrame.TextRange
                arr = prune_cites_in_shape(shp)
                txt, positioned, fallback = _collect_ieee_cites_in_shape(shp, arr)

                replacements = []
                tags_changed = False

                for hit in positioned:
                    c = hit["record"]
                    key = c.get("key")
                    if not key or key not in numbering:
                        continue

                    old_text = hit["old_text"]
                    new_cite = f"[{numbering[key]}]"

                    if old_text and old_text != new_cite:
                        replacements.append((hit["start"], hit["end"], new_cite))

                    if c.get("cite") != new_cite or c.get("style") != "ieee":
                        c["cite"] = new_cite
                        c["style"] = "ieee"
                        tags_changed = True

                # Fallback records have no reliable text position, but their tags
                # can still be normalized if their key is part of the numbering.
                for hit in fallback:
                    c = hit["record"]
                    key = c.get("key")
                    if not key or key not in numbering:
                        continue

                    new_cite = f"[{numbering[key]}]"
                    if c.get("cite") != new_cite or c.get("style") != "ieee":
                        c["cite"] = new_cite
                        c["style"] = "ieee"
                        tags_changed = True

                if replacements:
                    for start, end, new_cite in sorted(
                        replacements,
                        key=lambda item: item[0],
                        reverse=True,
                    ):
                        txt = txt[:start] + new_cite + txt[end:]

                    tr.Text = txt

                if replacements or tags_changed:
                    _save_shape_cites(shp, arr)

            except Exception:
                continue

        state["bib_keys"] = list(numbering.keys())
        save_doc_state(state)
        return numbering


def renumber_ieee_and_update(*, parent=None) -> bool:
    numbering = renumber_ieee_document_citations()
    style = "ieee"

    if not has_bibliography_anchor():
        return True

    try:
        cfg = get_cfg(allow_prompt=False, parent=parent)
    except RuntimeError:
        if parent is not None:
            parent.after(0, lambda: show_missing_zotero_config(parent))
        return False

    update_bibliography(
        list(numbering.keys()),
        style,
        cfg.api_key,
        cfg.library_id,
        cfg.library_type,
        numbering=numbering,
    )
    return True
# =========================================================


# ===================== CLI / Ribbon actions =====================
def _get_action_style(state):
    return state.get("style", DEFAULT_STYLE) or DEFAULT_STYLE


def _missing_zotero_config_error() -> RuntimeError:
    return RuntimeError(
        "Zotero ist nicht konfiguriert. "
        "Bitte starte zuerst den Picker und speichere die Zotero-Zugangsdaten."
    )


def run_document_update_workflow(parent=None) -> str:
    """
    Shared implementation for the document update workflow.

    Used by:
    - PickerApp.on_cleanup() / PickerApp.on_document_update()
    - CLI/Ribbon action --action update-document
    """
    try:
        cfg = get_cfg(allow_prompt=False, parent=parent)
    except RuntimeError as e:
        raise _missing_zotero_config_error() from e

    state = load_doc_state()
    style = _get_action_style(state)
    bibliography_already_updated = False

    styles = collect_document_citation_styles(fallback_style=style)
    if UNKNOWN_STYLE in styles or len([s for s in styles if s != UNKNOWN_STYLE]) > 1:
        LOG.warning(
            "Mixed citation style detected during document update: action=update-document state_style=%s document_styles=%s",
            style,
            styles,
        )

    new_keys = resync_bibliography_keys_from_document(state)
    _debug(f"DocumentUpdate: keys_after_prune={len(new_keys)} style={style}")

    if style in ("apa", "harvard1"):
        renormalize_all_sig_groups()
        _debug("DocumentUpdate: renormalize_all_sig_groups done")
        new_keys = resync_bibliography_keys_from_document(state)

    if style == "mla":
        normalize_mla_duplicate_labels()
        _debug("DocumentUpdate: normalize_mla_duplicate_labels done")
        new_keys = resync_bibliography_keys_from_document(state)

    if style == "ieee":
        had_bibliography_anchor = has_bibliography_anchor()
        ieee_update_ok = renumber_ieee_and_update(parent=parent)
        bibliography_already_updated = had_bibliography_anchor and ieee_update_ok
        new_keys = resync_bibliography_keys_from_document(state)
        _debug("DocumentUpdate: renumber_ieee_and_update done")

    if not new_keys:
        base_status = "Aktualisiert: keine Zitate mehr im Dokument gefunden."
    else:
        base_status = f"Aktualisiert: {len(new_keys)} Zitat(e) im Dokument."

    if not has_bibliography_anchor():
        return base_status + " (Kein Bibliographie-Ziel gesetzt.)"

    if not bibliography_already_updated:
        update_bibliography(
            new_keys,
            style,
            cfg.api_key,
            cfg.library_id,
            cfg.library_type,
        )

    return base_status


def run_rewrite_bibliography_workflow(parent=None) -> str:
    """
    Shared implementation for rewriting the bibliography.

    Used by:
    - PickerApp.on_bib_update()
    - CLI/Ribbon action --action rewrite-bibliography
    """
    try:
        cfg = get_cfg(allow_prompt=False, parent=parent)
    except RuntimeError as e:
        raise _missing_zotero_config_error() from e

    state = load_doc_state()
    style = _get_action_style(state)
    styles = collect_document_citation_styles(fallback_style=style)
    if UNKNOWN_STYLE in styles or len([s for s in styles if s != UNKNOWN_STYLE]) > 1:
        LOG.warning(
            "Mixed citation style detected during bibliography rewrite: action=rewrite-bibliography state_style=%s document_styles=%s",
            style,
            styles,
        )
    state["bib_keys"] = resync_bibliography_keys_from_document(state)
    keys = state.get("bib_keys", [])

    if not has_bibliography_anchor():
        if not keys:
            return "Bibliographie leer."
        raise RuntimeError(
            "Kein Bibliographie-Ziel gesetzt.\n"
            "Bitte gehe zur Bibliographie-Folie, wähle das Textfeld und nutze "
            "„Bibliographie-Ziel festlegen“."
        )

    update_bibliography(
        keys,
        style,
        cfg.api_key,
        cfg.library_id,
        cfg.library_type,
    )

    if keys:
        return f"Bibliographie aktualisiert ({code_to_label(style)})."
    return "Bibliographie geleert."


def run_set_bibliography_target_workflow(parent=None) -> str:
    """
    Shared implementation for setting the bibliography target.

    Used by:
    - PickerApp.on_set_anchor()
    - CLI/Ribbon action --action set-bibliography-target

    Important:
    Setting the bibliography target must not force an immediate full citation
    resync. In headless CLI/Ribbon mode, that resync can block while PowerPoint
    still exposes the newly selected target shape. The document update and
    bibliography rewrite workflows remain responsible for full resync.
    """
    set_bibliography_anchor_from_selection()

    state = load_doc_state()
    keys = state.get("bib_keys", []) or []
    anchor_count = len(_resolve_anchor_list())
    style = _get_action_style(state)

    _debug(
        f"SetBibTarget: keys_from_state={len(keys)} "
        f"anchors={anchor_count} style={style}"
    )

    if not anchor_count:
        raise RuntimeError(
            "Bibliographie-Ziel wurde gesetzt, ist danach aber nicht auflösbar."
        )

    if keys:
        try:
            cfg = get_cfg(allow_prompt=False, parent=parent)
        except RuntimeError as e:
            raise RuntimeError(
                f"Bibliographie-Ziel gesetzt. Gefundene Anker: {anchor_count}. "
                "Zotero ist nicht konfiguriert; die Bibliographie wurde noch nicht aktualisiert."
            ) from e

        update_bibliography(
            keys,
            style,
            cfg.api_key,
            cfg.library_id,
            cfg.library_type,
        )
        return f"Bibliographie-Ziel gesetzt. Gefundene Anker: {anchor_count} (Bibliographie aktualisiert)."

    return f"Bibliographie-Ziel gesetzt. Gefundene Anker: {anchor_count} (noch keine Zitate)."


def run_document_update_action(parent=None) -> str:
    """CLI/Ribbon wrapper for the shared document update workflow."""
    return run_document_update_workflow(parent=parent)


def run_rewrite_bibliography_action(parent=None) -> str:
    """CLI/Ribbon wrapper for the shared bibliography rewrite workflow."""
    return run_rewrite_bibliography_workflow(parent=parent)


def run_set_bibliography_target_action(parent=None) -> str:
    """CLI/Ribbon wrapper for the shared bibliography target workflow."""
    return run_set_bibliography_target_workflow(parent=parent)


ACTION_CHOICES = {
    "update-document": run_document_update_action,
    "rewrite-bibliography": run_rewrite_bibliography_action,
    "set-bibliography-target": run_set_bibliography_target_action,
}


def run_cli_action(action: str) -> int:
    """
    Run one Ribbon/CLI action without opening the picker GUI.

    Keep the Tk event loop alive and run the actual workflow in a worker thread,
    matching the PickerApp button execution model more closely than a synchronous
    CLI call on the main thread.
    """
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("zotero_ppt.log", encoding="utf-8"),
        ],
    )

    fn = ACTION_CHOICES.get(action)
    if fn is None:
        print(f"Unbekannte Aktion: {action}", file=sys.stderr)
        return 1

    root = tk.Tk()
    root.withdraw()

    result_box = {
        "exit_code": 1,
        "result": None,
        "error": None,
    }

    def _work():
        try:
            with com_context(f"cli-action-worker:{action}", use_lock=False):
                result = fn(parent=root)

            result_box["exit_code"] = 0
            result_box["result"] = result
            LOG.info("CLI action completed: action=%s result=%s", action, result)

        except Exception as e:
            result_box["exit_code"] = 1
            result_box["error"] = e
            LOG.exception("CLI action failed: action=%s", action)

        finally:
            try:
                root.after(0, root.quit)
            except Exception:
                pass

    threading.Thread(
        target=_work,
        name=f"ZP-CLI-{action}",
        daemon=True,
    ).start()

    try:
        root.mainloop()

        if result_box["exit_code"] == 0:
            result = result_box["result"] or "Aktion abgeschlossen."
            messagebox.showinfo("Zotero PowerPoint", result, parent=root)
            return 0

        err = result_box["error"]
        msg = str(err).strip() if err is not None else ""
        if not msg:
            msg = "Unbekannter Fehler. Details siehe zotero_ppt.log."

        messagebox.showerror("Zotero PowerPoint", msg, parent=root)
        print(msg, file=sys.stderr)
        return 1

    finally:
        try:
            root.destroy()
        except Exception:
            pass

# =========================================================

def code_to_label(code):
    return next((name for name, c in STYLE_CHOICES if c == code), code)


def label_to_code(label):
    return next((c for name, c in STYLE_CHOICES if name == label), DEFAULT_STYLE)




def fetch_zotero_item_by_key(key, cfg):
    """Fetch one Zotero item by key using the existing retry-aware HTTP helper."""
    base = f"https://api.zotero.org/{cfg.library_type}s/{cfg.library_id}"
    try:
        r = _safe_get(
            f"{base}/items/{key}",
            headers={
                "Zotero-API-Key": cfg.api_key,
                "Zotero-API-Version": "3",
                "Accept": "application/json",
            },
            timeout=HTTP_TIMEOUT,
            retries=3,
            context=f"item-json key={key}",
        )
    except requests.RequestException as e:
        raise RuntimeError(f"Zotero-Eintrag {key} konnte nicht geladen werden.") from e

    try:
        item = r.json()
    except ValueError as e:
        raise RuntimeError(f"Zotero-Eintrag {key} konnte nicht gelesen werden (ungültige JSON-Antwort).") from e

    if not isinstance(item, dict) or not item.get("data"):
        raise RuntimeError(f"Zotero-Eintrag {key} konnte nicht gelesen werden (unerwartete Antwort).")

    return item


def fetch_zotero_items_by_key(keys, cfg):
    items = {}
    for key in keys:
        items[key] = fetch_zotero_item_by_key(key, cfg)
    return items


def _format_cite_record_for_style(item, target_style):
    key = item.get("key")
    record = {"key": key, "style": target_style}

    if target_style == "ieee":
        record["cite"] = f"⟦zp:{key}⟧"
        return record

    if target_style == "mla":
        label, qualifier = _mla_label_parts_from_item(item)
        record.update({
            "cite": _format_mla_base_from_item(item),
            "sig": _make_sig(item),
            "mla_label": label,
        })
        if qualifier:
            record["mla_qualifier"] = qualifier
        return record

    record.update({
        "cite": _format_authoryear_base_from_item(item),
        "sig": _make_sig(item),
    })
    return record


def _collect_active_citation_occurrences():
    occurrences = []
    keys = []
    for scope, slide, shp in iter_citation_shapes():
        try:
            arr = prune_cites_in_shape(shp)
            for idx, c in enumerate(arr):
                key = c.get("key")
                cite = c.get("cite") or ""
                if not key or not cite:
                    continue
                occurrences.append({
                    "scope": scope,
                    "slide": slide,
                    "shape": shp,
                    "idx": idx,
                    "record": c,
                })
                if key not in keys:
                    keys.append(key)
        except Exception:
            continue
    return occurrences, keys


def convert_document_citation_style(target_style, *, parent=None) -> str:
    """Convert all active citation records and visible citations to target_style."""
    if target_style not in STYLE_CODES:
        raise RuntimeError(f"Unbekannter Zitierstil: {target_style}")

    state = load_doc_state()
    source_style = state.get("style", DEFAULT_STYLE) or DEFAULT_STYLE
    occurrences, keys = _collect_active_citation_occurrences()

    if not occurrences:
        state["style"] = target_style
        state["bib_keys"] = []
        save_doc_state(state)
        LOG.info(
            "Document citation style conversion skipped: empty document target_style=%s",
            target_style,
        )
        return f"Stil gesetzt: {code_to_label(target_style)}."

    LOG.info(
        "Document citation style conversion started: source_style=%s target_style=%s keys=%s",
        source_style,
        target_style,
        len(keys),
    )

    try:
        cfg = get_cfg(allow_prompt=False, parent=parent)
    except RuntimeError as e:
        raise RuntimeError(
            "Konvertierung fehlgeschlagen: Zotero ist nicht konfiguriert. Details siehe zotero_ppt.log."
        ) from e

    items_by_key = fetch_zotero_items_by_key(keys, cfg)
    new_records_by_key = {
        key: _format_cite_record_for_style(items_by_key[key], target_style)
        for key in keys
    }

    converted_shapes = 0
    try:
        by_shape = {}
        for occ in occurrences:
            slide = occ["slide"]
            shp = occ["shape"]
            shape_key = (occ["scope"], int(slide.SlideID), _get_shape_id(shp))
            by_shape.setdefault(shape_key, {"shape": shp, "items": []})["items"].append(occ)

        for _shape_key, pack in by_shape.items():
            shp = pack["shape"]
            arr = _load_shape_cites(shp)
            tr = shp.TextFrame.TextRange
            txt = tr.Text or ""
            text_changed = False
            tags_changed = False

            for occ in sorted(pack["items"], key=lambda item: item["idx"]):
                idx = occ["idx"]
                old = occ["record"]
                key = old.get("key")
                old_cite = old.get("cite") or ""
                fresh = dict(new_records_by_key[key])
                new_cite = fresh.get("cite") or ""

                txt, replaced = _replace_first(txt, old_cite, new_cite)
                if not replaced:
                    raise RuntimeError(
                        f"Zitat konnte nicht ersetzt werden: key={key} cite={old_cite!r}"
                    )
                text_changed = True

                if idx >= len(arr) or arr[idx].get("key") != key:
                    raise RuntimeError(
                        f"Citation-State konnte nicht eindeutig aktualisiert werden: key={key}"
                    )

                arr[idx].clear()
                arr[idx].update(fresh)
                if target_style != "mla":
                    arr[idx].pop("mla_label", None)
                    arr[idx].pop("mla_qualifier", None)
                tags_changed = True

            if text_changed:
                tr.Text = txt
                converted_shapes += 1
            if text_changed or tags_changed:
                _save_shape_cites(shp, arr)

        numbering = None
        if target_style in ("apa", "harvard1"):
            renormalize_all_sig_groups()
        elif target_style == "mla":
            normalize_mla_duplicate_labels()
        elif target_style == "ieee":
            numbering = renumber_ieee_document_citations()

        new_keys = resync_bibliography_keys_from_document(state)
        state["style"] = target_style
        state["bib_keys"] = new_keys if target_style != "ieee" else list((numbering or {}).keys())
        save_doc_state(state)

        if has_bibliography_anchor():
            update_bibliography(
                state.get("bib_keys", []),
                target_style,
                cfg.api_key,
                cfg.library_id,
                cfg.library_type,
                numbering=numbering,
            )

        LOG.info(
            "Document citation style conversion completed: target_style=%s keys=%s shapes=%s",
            target_style,
            len(keys),
            converted_shapes,
        )
        return f"Präsentation wurde auf „{code_to_label(target_style)}“ umgestellt."

    except Exception:
        LOG.exception(
            "Document citation style conversion failed: source_style=%s target_style=%s",
            source_style,
            target_style,
        )
        raise RuntimeError("Konvertierung fehlgeschlagen. Details siehe zotero_ppt.log.")


def confirm_document_style_conversion(parent, source_style, target_style, document_styles):
    description = _document_style_description(document_styles, source_style)
    msg = (
        f"In dieser Präsentation sind bereits Zitate {description} vorhanden.\n\n"
        f"Du hast „{code_to_label(target_style)}“ ausgewählt.\n\n"
        "Eine reine Stilumschaltung würde zu unterschiedlichen Zitierstilen innerhalb derselben Präsentation führen. "
        "Das wird nicht unterstützt.\n\n"
        f"Die Präsentation kann stattdessen vollständig auf „{code_to_label(target_style)}“ umgestellt werden. "
        "Dabei werden alle gespeicherten Zitate, sichtbaren Zitierstellen und die Bibliographie neu geschrieben.\n\n"
        f"Möchtest du die Präsentation jetzt auf „{code_to_label(target_style)}“ umstellen?"
    )
    return messagebox.askyesno("Zitierstil umstellen", msg, parent=parent)


def run_in_thread(name: str, fn, *, on_error=None, ui_parent=None):
    """
    Option A+ Worker:
    - unique thread name
    - COM initialized/uninitialized once per worker thread
    - COM locking only in actual COM helpers via com_context(...)
    - exceptions logged with full stack trace
    - no UI blocking
    """
    def _wrap():
        threading.current_thread().name = f"ZP-{name}"
        LOG.debug("Worker start: %s", name)
        try:
            # Initialize COM only; do NOT hold the global COM lock
            # for the entire worker lifetime.
            with com_context(f"worker:{name}", use_lock=False):
                fn()
            LOG.debug("Worker done: %s", name)
        except Exception as e:
            try:
                LOG.exception("Worker failed: %s", name)
            except Exception:
                _debug(f"Worker failed: {name}: {e}\n{traceback.format_exc()}")

            if on_error:
                try:
                    on_error(e)
                except Exception:
                    pass
            elif ui_parent is not None:
                try:
                    msg = str(e).strip() or "Unbekannter Fehler. Details siehe zotero_ppt.log."
                    ui_parent.after(0, lambda: messagebox.showerror("Fehler", msg, parent=ui_parent))
                except Exception:
                    pass

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

        # Optional: minimum size so the UI does not look cramped
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
            self._ppt_startup_error = str(e)   # remember for later

        self.results = []

        # Search: debounce + latest-only handling to avoid flicker and out-of-order threads
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

        # Result list with vertical and horizontal scrollbars
        listfrm = ttk.Frame(frm)
        listfrm.pack(fill="both", expand=True, pady=8)

        self.listbox = tk.Listbox(listfrm, height=16)
        self.listbox.bind("<Button-1>", lambda e: self.listbox.focus_set())

        vsb = ttk.Scrollbar(listfrm, orient="vertical", command=self.listbox.yview)
        hsb = ttk.Scrollbar(listfrm, orient="horizontal", command=self.listbox.xview)

        # Connect listbox and scrollbars
        self.listbox.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout: large listbox, vertical scrollbar on the right, horizontal scrollbar below
        self.listbox.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, columnspan=2, sticky="ew")

        # Frame resize rules
        listfrm.rowconfigure(0, weight=1)
        listfrm.columnconfigure(0, weight=1)

        self.listbox.bind("<Double-Button-1>", self.on_insert_click)

        # --- Optional UX: horizontal scrolling with Shift + mouse wheel ---
        def _on_mousewheel(ev):
            # Windows/macOS: ev.delta; Shift pressed -> horizontal, otherwise vertical
            shift = ((ev.state & 0x0001) != 0) or ((ev.state & 0x0004) != 0)  # depending on Tk/platform
            if shift:
                # X-scroll depending on wheel direction
                step = -1 if ev.delta > 0 else 1
                self.listbox.xview_scroll(step, "units")
            else:
                step = -1 if ev.delta > 0 else 1
                self.listbox.yview_scroll(step, "units")
            return "break"

        def _on_shift_mousewheel(ev):
            # Explicit Shift variant in case Tk emits it separately
            step = -1 if ev.delta > 0 else 1
            self.listbox.xview_scroll(step, "units")
            return "break"

        # Windows / macOS
        self.listbox.bind("<MouseWheel>", _on_mousewheel)
        self.listbox.bind("<Shift-MouseWheel>", _on_shift_mousewheel)

        # Linux (X11): mouse wheel events are button events
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

        # Buttons: 2x2 grid instead of one row so they do not disappear at small window heights
        row2 = ttk.Frame(frm)
        row2.pack(fill="x", pady=4)

        buttons = [
            ("Zitation einfügen", self.on_insert_click),
            ("Dokument aktualisieren", self.on_document_update),
            ("Bibliographie-Ziel festlegen", self.on_set_anchor),
            ("Bibliographie neu schreiben", self.on_bib_update),
        ]

        for i, (label, cmd) in enumerate(buttons):
            btn = ttk.Button(row2, text=label, command=cmd)

            # Optional: calmer layout with similar widths and clean tab focus behavior
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

        # Check config only after the GUI has started; never block in __init__
        self.root.after(0, self._ensure_cfg_ready)

        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(300, lambda: self.root.attributes("-topmost", False))

    def set_status(self, msg=""):
        if msg:
            self.status.config(text=msg)
            return

        try:
            self.status.config(text=get_status_summary())
        except Exception as e:
            self.status.config(text=f"(Status nicht verfügbar: {e})")

    def _ensure_cfg_ready(self):
        # 1) First try without prompting, non-blocking
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

        # 2) The GUI is alive now, so a modal prompt is allowed
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
        target_style = label_to_code(self.style_var.get())
        previous_style = self.state.get("style", DEFAULT_STYLE) or DEFAULT_STYLE

        if target_style == previous_style:
            self.set_status(f"Stil gesetzt: {code_to_label(previous_style)}")
            return

        document_styles = collect_document_citation_styles(fallback_style=previous_style)

        if not document_styles:
            self.state["style"] = target_style
            save_doc_state(self.state)
            self.set_status(f"Stil gesetzt: {code_to_label(target_style)}")
            LOG.info("Style set in empty document: target_style=%s", target_style)
            return

        existing_style = get_existing_document_style(fallback_style=previous_style)
        if existing_style == target_style:
            self.state["style"] = target_style
            save_doc_state(self.state)
            self.set_status(f"Stil synchronisiert: {code_to_label(target_style)}")
            LOG.info("Style synchronized with document citations: target_style=%s", target_style)
            return

        if not confirm_document_style_conversion(
            self.root,
            existing_style,
            target_style,
            document_styles,
        ):
            self.style_var.set(code_to_label(previous_style))
            self.set_status("Stilwechsel abgebrochen.")
            LOG.info(
                "Document citation style conversion cancelled: source_style=%s target_style=%s document_styles=%s",
                previous_style,
                target_style,
                document_styles,
            )
            return

        self.style_var.set(code_to_label(previous_style))
        self.set_status(f"Konvertierung auf {code_to_label(target_style)} läuft...")

        def _work():
            result = convert_document_citation_style(target_style, parent=self.root)

            def _finish_ui():
                self.state = load_doc_state()
                self.style_var.set(code_to_label(self.state.get("style", target_style)))
                self.set_status(result)

            self.root.after(0, _finish_ui)

        def _on_error(e):
            def _finish_error():
                self.style_var.set(code_to_label(previous_style))
                self.set_status("Konvertierung fehlgeschlagen. Details siehe zotero_ppt.log.")
                messagebox.showerror(
                    "Zitierstil umstellen",
                    str(e) or "Konvertierung fehlgeschlagen. Details siehe zotero_ppt.log.",
                    parent=self.root,
                )

            self.root.after(0, _finish_error)

        run_in_thread("StyleConversion", _work, on_error=_on_error, ui_parent=self.root)

    def on_key(self, event=None):
        if self.z is None:
            self.set_status("Bitte Zotero konfigurieren…")
            return

        # Debounce: do not search immediately on every key press
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
        # Keep dicts only
        incoming = [it for it in (items or []) if isinstance(it, dict)]

        def _ui():
            # Latest-only: ignore older/outdated search threads to avoid flicker
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
        # 0) Get the current result item on the UI thread
        it = self.current_item()
        if not it:
            self.set_status("Kein Eintrag ausgewählt.")
            return

        key = it.get("key")
        style = self.state.get("style", DEFAULT_STYLE)

        if not key:
            self.set_status("Ungültiger Eintrag (kein Key).")
            return

        _debug(f"Insert click: style={style}, key={key}")

        # Inform the UI immediately
        self.set_status("Einfügen läuft...")

        def _work():
            # Worker: all COM work and optional Web API calls
            state = load_doc_state()
            style_local = state.get("style", style) or DEFAULT_STYLE

            if block_mixed_style_insert_or_show_message(self.root, style_local, fallback_style=style_local):
                self.root.after(
                    0,
                    lambda: self.set_status("Abgebrochen: anderer Zitierstil im Dokument erkannt."),
                )
                return

            # IEEE: placeholder + renumbering, handled entirely in the worker
            if style_local == "ieee":
                placeholder = f"⟦zp:{key}⟧"
                shp = ppt_insert_text_at_cursor(f" {placeholder}")

                arr = _load_shape_cites(shp)
                arr.append({
                    "key": key,
                    "cite": placeholder,
                    "style": "ieee",
                })
                _save_shape_cites(shp, arr)

                ok = renumber_ieee_and_update(parent=self.root)

                def _finish_ieee():
                    # Update state if possible; renumber_ieee_and_update writes doc_state
                    try:
                        self.state = load_doc_state()
                    except Exception:
                        pass
                    if ok:
                        self.set_status("IEEE: Platzhalter eingefügt und renummeriert.")
                    else:
                        self.set_status("IEEE: Platzhalter eingefügt. Zotero konfigurieren für Renummerierung/Bibliographie.")
                self.root.after(0, _finish_ieee)
                return

            # 2) Collect existing citations by Zotero key
            by_key = collect_all_cites_by_key()

            if style_local in ("apa", "harvard1") and key in by_key:
                cite = by_key[key].get("cite") or _format_authoryear_base_from_item(it)
                sig = by_key[key].get("sig") or _make_sig(it)
                record_extra = {"style": style_local}

            elif style_local == "mla":
                sig = _make_sig(it)
                label, qualifier = _mla_label_parts_from_item(it)

                if key in by_key and _is_probable_legacy_mla_record(by_key[key]):
                    cite = by_key[key].get("cite") or _format_mla_base_from_item(it)
                else:
                    cite = _format_mla_base_from_item(it)

                record_extra = {
                    "style": "mla",
                    "mla_label": label,
                }
                if qualifier:
                    record_extra["mla_qualifier"] = qualifier

            else:
                sig = _make_sig(it)
                cite = _format_authoryear_base_from_item(it)
                record_extra = {
                    "style": style_local,
                }

            # 3) Insert citation at cursor position; strictly requires a real text cursor
            shp = ppt_insert_text_at_cursor(cite)
            _debug(f"Inserted cite: {cite}")

            # 4) Store citation in the tag of the actual shape used
            arr = _load_shape_cites(shp)
            record = {
                "key": key,
                "cite": cite,
                "sig": sig,
            }
            record.update(record_extra)
            arr.append(record)
            _save_shape_cites(shp, arr)

            # 5) Style-specific citation normalization
            if style_local in ("apa", "harvard1"):
                normalize_sig_group(sig)
                _debug(f"Normalized sig group: {sig}")

            elif style_local == "mla":
                normalize_mla_duplicate_labels()
                refreshed = collect_all_cites_by_key()
                cite = (refreshed.get(key) or {}).get("cite") or cite
                _debug("Normalized MLA duplicate labels")

            # 6) Rebuild bibliography keys from the document and write doc_state
            state["bib_keys"] = resync_bibliography_keys_from_document(state)
            _debug(f"Resync keys: {len(state.get('bib_keys', []))}")

            # 7) Auto-update bibliography if an anchor exists
            did_bib = False
            if has_bibliography_anchor():
                try:
                    cfg = get_cfg(allow_prompt=False, parent=self.root)
                except RuntimeError:
                    # Show UI dialog on the UI thread
                    self.root.after(0, lambda: show_missing_zotero_config(self.root))
                    cfg = None

                if cfg is not None:
                    update_bibliography(
                        state.get("bib_keys", []),
                        style_local,
                        cfg.api_key,
                        cfg.library_id,
                        cfg.library_type
                    )
                    did_bib = True
                    _debug("Auto bibliography update triggered")

            # UI-Finish
            def _finish_ui():
                self.state = state
                if did_bib:
                    self.set_status(f"Eingefügt: {cite} (Bibliographie aktualisiert)")
                else:
                    self.set_status(
                        f"Eingefügt: {cite} "
                        f"(Bibliographie wird geschrieben, sobald Ziel gesetzt ist)"
                    )
            self.root.after(0, _finish_ui)

        run_in_thread(
            "Insert",
            _work,
            ui_parent=self.root
        )

    def on_set_anchor(self):
        # Do not open multiple anchor windows
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

        def _finish_after_set(result):
            try:
                self.state = load_doc_state()
            except Exception:
                pass

            self._anchor_win = None
            try:
                if win.winfo_exists():
                    win.destroy()
            except Exception:
                pass

            self.set_status(result)

        def _do_set():
            def _work():
                result = run_set_bibliography_target_workflow(parent=self.root)
                self.root.after(0, lambda: _finish_after_set(result))

            run_in_thread(
                "SetBibAnchor",
                _work,
                ui_parent=self.root
            )

        ttk.Button(btn_row, text="Jetzt setzen", command=_do_set).pack(side="left")

        def _cancel():
            self._anchor_win = None
            win.destroy()

        ttk.Button(btn_row, text="Abbrechen", command=_cancel).pack(side="right")

        # --- Center window above the picker after layout is complete ---
        self.root.update_idletasks()
        win.update_idletasks()

        confirm_w = win.winfo_width()
        confirm_h = win.winfo_height()

        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        # Desired position, centered above the picker
        x = root_x + (root_w - confirm_w) // 2
        y = root_y + (root_h - confirm_h) // 2

        # --- Clamp to screen bounds ---
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()

        x = max(0, min(x, screen_w - confirm_w))
        y = max(0, min(y, screen_h - confirm_h))

        win.geometry(f"+{x}+{y}")

        try:
            win.lift()
        except Exception:
            pass

        # Bring PowerPoint to the foreground so the user can click the target field
        _activate_powerpoint()

        # Ensure the anchor window is above the picker when the user returns
        try:
            win.attributes("-topmost", True)
            win.lift()
        except Exception:
            pass

        _activate_powerpoint()

    def on_document_update(self):
        """
        Primary user workflow:
        - resyncs slide and notes citations with stored citation metadata
        - applies style-specific repairs where needed:
            - APA/Harvard: author-year disambiguation rollback/rebuild
            - IEEE: slide and notes citation renumbering
        - updates the bibliography if a bibliography target exists
        """
        return self.on_cleanup()

    def on_bib_update(self):
        try:
            get_cfg(allow_prompt=False, parent=self.root)
        except RuntimeError:
            show_missing_zotero_config(self.root)
            return

        # Update UI immediately: running state
        self.set_status("Bibliographie wird aktualisiert...")

        def _work():
            result = run_rewrite_bibliography_workflow(parent=self.root)

            def _finish_ui():
                try:
                    self.state = load_doc_state()
                except Exception:
                    pass
                self.set_status(result)

            self.root.after(0, _finish_ui)

        # Run bibliography update in the background
        run_in_thread("BibUpdate", _work, ui_parent=self.root)

    def on_cleanup(self):
        _debug("Document update clicked")

        try:
            get_cfg(allow_prompt=False, parent=self.root)
        except RuntimeError:
            show_missing_zotero_config(self.root)
            return

        # Update UI immediately: running state
        self.set_status("Dokument wird aktualisiert...")

        def _work():
            result = run_document_update_workflow(parent=self.root)

            def _finish_ui():
                try:
                    self.state = load_doc_state()
                except Exception:
                    pass
                self.set_status(result)

            self.root.after(0, _finish_ui)

        # Run document update workflow in background worker
        run_in_thread("DocumentUpdate", _work, ui_parent=self.root)


def main():
    parser = argparse.ArgumentParser(
        description="Zotero PowerPoint picker and Ribbon actions."
    )
    parser.add_argument(
        "--action",
        choices=sorted(ACTION_CHOICES.keys()),
        help="Run a PowerPoint action without opening the picker GUI.",
    )
    args = parser.parse_args()

    if args.action:
        return run_cli_action(args.action)

    logging.basicConfig(
        level=logging.DEBUG,  # useful for COM stabilization; later possibly INFO
        format="%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("zotero_ppt.log", encoding="utf-8"),
        ],
    )

    root = tk.Tk()

    PickerApp(root)
    root.mainloop()
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as ex:
        print(f"Fatal: {ex}", file=sys.stderr)
        sys.exit(1)
