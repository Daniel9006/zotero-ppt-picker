from __future__ import annotations

import json
import os
import webbrowser

from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

# Optional: .env Support (nur wenn installiert)
try:
    from dotenv import load_dotenv  # type: ignore
except Exception:
    load_dotenv = None

# Optional: platformdirs für saubere, plattformneutrale Pfade
try:
    from platformdirs import user_config_dir  # type: ignore
except Exception:
    user_config_dir = None

import tkinter as tk
from tkinter import ttk


ENV_API_KEY = "ZOTERO_API_KEY"
ENV_LIBRARY_ID = "ZOTERO_LIBRARY_ID"
ENV_LIBRARY_TYPE = "ZOTERO_LIBRARY_TYPE"

DEFAULT_LIBRARY_TYPE = "user"

APP_NAME = "ZoteroPowerPoint"
APP_AUTHOR = "YourOrg"  # kann auch dein Repo-/Teamname sein


class ConfigError(Exception):
    """Konfigurationsfehler, die user-facing behandelt werden sollen."""
    pass


@dataclass(frozen=True)
class ZoteroConfig:
    api_key: str
    library_id: str
    library_type: str = DEFAULT_LIBRARY_TYPE


# -------------------------
# Storage (User Config File)
# -------------------------

def get_user_config_path() -> str:
    """
    Plattfomneutraler Speicherort pro User.
    Liegt NICHT im Repo, daher kein Git-Risiko.
    """
    if user_config_dir:
        base = user_config_dir(APP_NAME, APP_AUTHOR)
    else:
        # Fallback ohne Abhängigkeit:
        if os.name == "nt" and os.environ.get("APPDATA"):
            base = os.path.join(os.environ["APPDATA"], APP_NAME)
        else:
            base = os.path.join(os.path.expanduser("~"), ".config", APP_NAME)

    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "config.json")


def read_user_config() -> Dict[str, Any]:
    path = get_user_config_path()
    if not os.path.exists(path):
        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError) as e:
        raise ConfigError(f"Lokale Config-Datei ist unlesbar:\n{path}") from e


def write_user_config(cfg: ZoteroConfig) -> str:
    validate_config(cfg)
    path = get_user_config_path()
    tmp = path + ".tmp"

    data = {
        "api_key": cfg.api_key,
        "library_id": cfg.library_id,
        "library_type": cfg.library_type,
    }

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    os.replace(tmp, path)
    return path


# -------------------------
# Loading / Merging
# -------------------------

def load_from_env(*, load_dotenv_file: bool = True) -> Dict[str, str]:
    """
    Liest ENV + optional .env (wenn python-dotenv installiert).
    .env bleibt optional und sollte in .gitignore sein.
    """
    if load_dotenv_file and load_dotenv is not None:
        # lädt .env im aktuellen Arbeitsverzeichnis (Repo-Root), wenn vorhanden
        load_dotenv(override=False)

    out: Dict[str, str] = {}
    if os.environ.get(ENV_API_KEY):
        out["api_key"] = os.environ[ENV_API_KEY].strip()
    if os.environ.get(ENV_LIBRARY_ID):
        out["library_id"] = os.environ[ENV_LIBRARY_ID].strip()
    if os.environ.get(ENV_LIBRARY_TYPE):
        out["library_type"] = os.environ[ENV_LIBRARY_TYPE].strip()
    return out


def load_from_user_file() -> Dict[str, str]:
    data = read_user_config()
    return {
        "api_key": str(data.get("api_key") or "").strip(),
        "library_id": str(data.get("library_id") or "").strip(),
        "library_type": str(data.get("library_type") or DEFAULT_LIBRARY_TYPE).strip(),
    }


def validate_config(cfg: ZoteroConfig) -> None:
    if not cfg.api_key or not cfg.api_key.strip():
        raise ConfigError("Zotero API-Key fehlt.")
    if not cfg.library_id or not cfg.library_id.strip():
        raise ConfigError("Zotero Library ID fehlt.")
    if not cfg.library_id.isdigit():
        raise ConfigError("Zotero Library ID muss numerisch sein.")
    if cfg.library_type not in ("user", "group"):
        raise ConfigError("ZOTERO_LIBRARY_TYPE muss 'user' oder 'group' sein.")


def build_config(merged: Dict[str, str]) -> ZoteroConfig:
    return ZoteroConfig(
        api_key=merged.get("api_key", "").strip(),
        library_id=merged.get("library_id", "").strip(),
        library_type=(merged.get("library_type", DEFAULT_LIBRARY_TYPE).strip() or DEFAULT_LIBRARY_TYPE),
    )


def load_zotero_config(*, allow_prompt: bool, parent: Optional[tk.Misc] = None) -> ZoteroConfig:
    """
    Priorität:
      1) User-Config-Datei
      2) ENV (+ optional .env) überschreibt
      3) Prompt (wenn allow_prompt=True) als Fallback

    Rückgabe: validierte ZoteroConfig
    """
    merged: Dict[str, str] = {"library_type": DEFAULT_LIBRARY_TYPE}

    # 1) User-Config
    file_cfg = load_from_user_file()
    for k, v in file_cfg.items():
        if v:
            merged[k] = v

    # 2) ENV override
    env_cfg = load_from_env(load_dotenv_file=True)
    for k, v in env_cfg.items():
        if v:
            merged[k] = v

    cfg = build_config(merged)

    try:
        validate_config(cfg)
        return cfg
    except ConfigError:
        if not allow_prompt:
            raise

    # 3) GUI Prompt
    cfg2, action = prompt_zotero_config(parent=parent, initial=cfg)

    if action == "cancel":
        raise ConfigError("Abgebrochen: Zotero-Zugangsdaten wurden nicht konfiguriert.")

    validate_config(cfg2)

    if action == "save":
        write_user_config(cfg2)

    # action == "session": nur zurückgeben, nicht persistieren
    return cfg2


# -------------------------
# Tkinter Prompt Dialog
# -------------------------

def prompt_zotero_config(
    *,
    parent: Optional[tk.Misc],
    initial: Optional[ZoteroConfig] = None
) -> Tuple[ZoteroConfig, str]:
    """
    Zeigt einen modalen Dialog.
    Return: (ZoteroConfig, action)
      action in {"save", "session", "cancel"}
    """
    # Root/Parent Handling
    owns_root = False
    root = None
    if parent is None:
        owns_root = True
        root = tk.Tk()
        # NICHT withdraw(): kann unter Windows Child-Toplevels "unsichtbar" machen
        root.title("")
        root.geometry("1x1+0+0")
        try:
            root.attributes("-alpha", 0.0)  # praktisch unsichtbar, aber echtes Owner-Window
        except Exception:
            root.iconify()  # Fallback
        parent = root

    initial = initial or ZoteroConfig(api_key="", library_id="", library_type=DEFAULT_LIBRARY_TYPE)

    win = tk.Toplevel(parent)

    win.title("Zotero Zugangsdaten konfigurieren")
    win.resizable(True, False)   # horizontal optional, vertikal fix

    # Wenn wir selbst ein Root erzeugen (Testlauf/Standalone),
    # dann kein transient() erzwingen – in echten Apps ist parent bereits gesetzt.
    if not owns_root and parent is not None:
        win.transient(parent)

    # Initiale Breite setzen (Höhe später dynamisch)
    ww = 560
        
    # --- HARD FIX: sichtbar + vorne + fokus ---
    try:
        win.state("normal")
    except Exception:
        pass

    win.update_idletasks()
    try:
        win.deiconify()
    except Exception:
        pass

    try:
        win.lift()
        win.attributes("-topmost", True)
        win.focus_force()
        win.after(350, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

    win.update_idletasks()
    win.wait_visibility()
    win.grab_set()  # modal

    api_var = tk.StringVar(value=initial.api_key)
    id_var = tk.StringVar(value=initial.library_id)
    type_var = tk.StringVar(value=initial.library_type or DEFAULT_LIBRARY_TYPE)
    err_var = tk.StringVar(value="")

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="x", expand=False)

    ttk.Label(frm, text="Bitte trage deine Zotero-Zugangsdaten ein.", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    ttk.Label(
        frm,
        text="Hinweis: Du kannst sie lokal speichern (im Benutzerprofil) oder nur für diese Sitzung verwenden.",
        wraplength=490
    ).pack(anchor="w", pady=(4, 10))

    grid = ttk.Frame(frm)
    grid.pack(fill="x", expand=False)

    ttk.Label(grid, text="Library Type:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
    type_box = ttk.Combobox(grid, textvariable=type_var, values=["user", "group"], state="readonly", width=10)
    type_box.grid(row=0, column=1, sticky="w", pady=4)

    # Dynamisches Label für ID (User ID vs Group ID)
    id_label_var = tk.StringVar(value="User ID (Zotero Library ID):")
    id_hint_var = tk.StringVar(value="")      # wird durch _refresh_id_labels gesetzt
    id_link_var = tk.StringVar(value="")      # URL oder leer

    id_label = ttk.Label(grid, textvariable=id_label_var)
    id_label.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)

    id_entry = ttk.Entry(grid, textvariable=id_var, width=30)
    id_entry.grid(row=1, column=1, sticky="w", pady=4)

    # Hint (Text) unter dem Eingabefeld
    id_hint = ttk.Label(grid, textvariable=id_hint_var, wraplength=340, justify="left")
    id_hint.grid(row=2, column=1, sticky="w", pady=(0, 2))

    # Klickbarer Link (nur sichtbar, wenn id_link_var gesetzt ist)
    link_lbl = ttk.Label(grid, text="", wraplength=340, justify="left", cursor="")
    link_lbl.grid(row=3, column=1, sticky="w", pady=(0, 6))

    
    default_link_fg = link_lbl.cget("foreground")

    def _open_link(_event=None):
        url = (id_link_var.get() or "").strip()
        if url:
            webbrowser.open_new_tab(url)

    link_lbl.bind("<Button-1>", _open_link)

    def _refresh_id_labels(*_):
        lt = (type_var.get() or DEFAULT_LIBRARY_TYPE).strip().lower()
        if lt == "group":
            id_label_var.set("Group ID (Zotero Library ID):")
            id_hint_var.set(
                "Für Gruppenbibliotheken ist die Library ID die Group ID.\n"
                "Du findest sie auf der Gruppen-Seite – meistens in der URL als Zahl."
            )
            # Ergänzter Hint + Link (klickbar)
            id_link_var.set("https://www.zotero.org/groups")
        else:
            id_label_var.set("User ID (Zotero Library ID):")
            id_hint_var.set(
                "Für persönliche Bibliotheken entspricht die Library ID der User ID.\n"
                "Zotero Web → Settings → Security → Applications"
            )
            id_link_var.set("https://www.zotero.org/settings/security/Applications")

        # Link-Label optisch wie Link darstellen oder „ausblenden“
        if id_link_var.get().strip():
            try:
                link_lbl.configure(foreground="blue")
            except Exception:
                pass
            link_lbl.configure(text=id_link_var.get())
            link_lbl.configure(cursor="hand2")
        else:
            link_lbl.configure(text="")
            link_lbl.configure(cursor="")
            try:
                link_lbl.configure(foreground=default_link_fg)
            except Exception:
                pass

    # initial setzen + bei Änderung updaten
    _refresh_id_labels()
    try:
        type_var.trace_add("write", _refresh_id_labels)
    except Exception:
        # Fallback für sehr alte Tk-Versionen
        type_var.trace("w", _refresh_id_labels)

    # API Key (nach ID + Hint + Link)
    ttk.Label(grid, text="API Key:").grid(row=4, column=0, sticky="w", padx=(0, 10), pady=4)
    api_entry = ttk.Entry(grid, textvariable=api_var, width=44, show="•")
    api_entry.grid(row=4, column=1, sticky="w", pady=4)


    err_lbl = ttk.Label(frm, textvariable=err_var, foreground="#b00020", wraplength=490)
    err_lbl.pack(anchor="w", pady=(10, 6))
    err_lbl.pack_forget()

    btns = ttk.Frame(frm)
    btns.pack(fill="x", pady=(8, 0), side="bottom")

    action = {"value": "cancel"}

    def _make_cfg() -> ZoteroConfig:
        return ZoteroConfig(
            api_key=api_var.get().strip(),
            library_id=id_var.get().strip(),
            library_type=(type_var.get().strip() or DEFAULT_LIBRARY_TYPE),
        )

    def _validate_show() -> bool:
        try:
            validate_config(_make_cfg())
            err_var.set("")
            err_lbl.pack_forget()
            return True
        except ConfigError as e:
            err_var.set(str(e))
            err_lbl.pack(anchor="w", pady=(10, 6))
            return False

    def on_save():
        if not _validate_show():
            return
        action["value"] = "save"
        win.destroy()

    def on_session():
        if not _validate_show():
            return
        action["value"] = "session"
        win.destroy()

    def on_cancel():
        action["value"] = "cancel"
        win.destroy()

    ttk.Button(btns, text="Speichern & Weiter", command=on_save).pack(side="left")
    ttk.Button(btns, text="Nur diese Sitzung", command=on_session).pack(side="left", padx=8)
    ttk.Button(btns, text="Abbrechen", command=on_cancel).pack(side="right")

    win.protocol("WM_DELETE_WINDOW", on_cancel)

    # Tastaturbedienung
    win.bind("<Return>", lambda e: on_save())
    win.bind("<Escape>", lambda e: on_cancel())
  
    # Fokus sinnvoll setzen
    if not initial.library_id:
        id_entry.focus_set()
    elif not initial.api_key:
        api_entry.focus_set()
    else:
        id_entry.focus_set()

    # --- Fensterhöhe exakt an Inhalt anpassen (kein Leerraum) ---
    win.update_idletasks()

    req_h = win.winfo_reqheight()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max(0, (sw - ww) // 2)
    y = max(0, (sh - req_h) // 2)

    win.geometry(f"{ww}x{req_h}+{x}+{y}")
    win.minsize(ww, req_h)
    
    win.wait_window()

    if owns_root and root is not None:
        try:
            root.destroy()
        except Exception:
            pass

    cfg = _make_cfg()
    return cfg, action["value"]
