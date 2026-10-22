#!/usr/bin/env python3
"""DocuBuilder – Buttons, die hinterlegte Excel-Listen öffnen.

Läuft unter Windows und Linux, benötigt nur Python 3 (tkinter ist enthalten).
Anmeldung mit Benutzerkonten (Admin fest); Einstellungen verschlüsselt und versteckt (siehe README).
"""
import hashlib
import hmac
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

APP_NAME = "DocuBuilder"
# Version und Kanal werden von tools/version.py gepflegt (Nightly-Builds setzen sie nur beim Bauen).
APP_VERSION = "1.1.1.0"
APP_CHANNEL = "stable"
APP_TITLE = APP_NAME + (" Nightly" if APP_CHANNEL == "nightly" else "")
# Update-Funktion erst freischalten, wenn DocuBuilder signiert ausgeliefert wird (siehe CODE_SIGNING.md):
# unsignierte Updates blockiert Smart App Control.
UPDATES_ENABLED = False

ADMIN_NAME = "Admin"                       # festes Konto, kann nicht gelöscht werden
MIN_PW_LEN = 4
# Tastenkombinationen, die das Default-Fenster nach der ersten Einrichtung öffnen.
# Die zweite gilt als Ersatz, falls das System ALT+F11 selbst belegt.
DEFAULT_SHORTCUTS = ("<Alt-F11>", "<Control-Shift-F11>")
EXCEL_EXT = (".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".csv", ".ods")
MAX_ITEMS_LIMIT = 200

# Farbschemata nach den Farben des Icons (Blau #3654E8, Dunkelblau #1F2A44, Orange #FFB020).
# Jedes Schema hat eine helle und eine dunkle Palette. Ein neues Schema ist ein neuer Eintrag in
# COLOR_SCHEMES; DEFAULT_SCHEME legt fest, was gilt, solange ein Benutzer nichts anderes wählt.
# Optionale Schlüssel: primary/primary_fg (Hauptknöpfe wie "Anmelden"), title (Programmtitel).
BRAND = {
    "light": dict(bg="#F4F6FB", fg="#1F2A44", panel="#FFFFFF", line="#C5CDE0", mute="#56607A",
                  btn="#E6EAF5", btn_hover="#D9E0FC", disabled="#98A1B6", tab="#E2E7F3",
                  accent="#3654E8", accent_fg="#FFFFFF"),
    "dark": dict(bg="#141B2C", fg="#EEF1F8", panel="#1F2A44", line="#46557E", mute="#AFB9D2",
                 btn="#29355A", btn_hover="#354471", disabled="#6F7A98", tab="#1A2338",
                 accent="#6F86FF", accent_fg="#0E1426"),
}


def _scheme(light=None, dark=None, base=BRAND):
    return {"light": dict(base["light"], **(light or {})), "dark": dict(base["dark"], **(dark or {}))}


ORANGE = dict(primary="#FFB020", primary_fg="#1F2A44")
COLOR_SCHEMES = {
    "blau_orange": ("Blau mit orangem Hauptknopf", _scheme(ORANGE, ORANGE)),
    "blau": ("Blau", _scheme()),
    "orange": ("Orange", _scheme(dict(accent="#FFB020", accent_fg="#1F2A44", title="#3654E8",
                                      btn_hover="#FFEBC2"),
                                 dict(accent="#FFB020", accent_fg="#1F2A44", title="#8DA0FF"))),
    "gruen": ("Grün (bis Version 1.1)", {
        "light": dict(bg="#f3f4f2", fg="#1d2a24", panel="#ffffff", accent="#1f6b4a",
                      accent_fg="#ffffff", line="#b4beb8", mute="#5c6962",
                      btn="#e4e9e5", btn_hover="#d3e2d8", disabled="#8f9a94", tab="#dde3df"),
        "dark": dict(bg="#1a201d", fg="#f0f4f1", panel="#29332e", accent="#5cc493",
                     accent_fg="#0b130f", line="#56685e", mute="#b3c2ba",
                     btn="#34413a", btn_hover="#435249", disabled="#7f8e86", tab="#2e3833")}),
}
DEFAULT_SCHEME = "blau_orange"

DEFAULT_SETTINGS = {"theme": "system", "scheme": DEFAULT_SCHEME, "max_items": 20, "files_dir": "",
                    "setup_done": False, "users": [], "update_check": False, "categories": []}
# Bereiche: Einträge gehören zu einem Bereich (oder zu keinem = für alle sichtbar), Benutzer sehen
# nur die Bereiche, die der Admin ihnen freigibt. Vorschlag bei der Ersteinrichtung:
DEFAULT_CATEGORIES = ["Arbeitsvorbereitung", "Angebote", "Abrechnung"]
ALL_LABEL = "(alle)"


# ---------------------------------------------------------------- Konfiguration
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource(name):
    """Mitgelieferte Datei aus assets/ (in der .exe entpackt PyInstaller sie nach _MEIPASS)."""
    return os.path.join(getattr(sys, "_MEIPASS", None) or app_dir(), "assets", name)


# Gleiche ID wie die Verknüpfungen des Installationsprogramms: Windows gruppiert das Fenster
# dann in der Taskleiste unter dem DocuBuilder-Icon statt unter Python.
APP_USER_MODEL_ID = "PatrickLange.DocuBuilder"


WIN = sys.platform.startswith("win")
MAGIC = b"DBC1"
# Eingebetteter Schlüsselanteil: verhindert Lesen/Bearbeiten im Editor, ist aber kein Schutz
# gegen Entschlossene, die den Quelltext kennen (siehe README).
_SEED = b"DocuBuilder/settings/v1/7c2e9b41"
PORTABLE_FLAG = "portable.txt"
# Netzwerk-Version: server.txt neben dem Programm (oder unter Linux in /etc/docubuilder) nennt den
# gemeinsamen Ordner auf dem Server. Erste Zeile: Ordner, z. B. \\Server\DocuBuilder.
# Weitere Zeilen "Windows-Pfad = Linux-Pfad" übersetzen Pfade für Linux-Arbeitsplätze.
SERVER_FILE = "server.txt"
LINUX_SHARED = "/etc/docubuilder"
OFFLINE_TEXT = ("Der Server ist nicht erreichbar. Bitte Netzwerk prüfen und erneut anmelden.")
NO_RIGHTS = ("Nur der Admin oder Benutzer mit dem Recht „Darf ändern“ dürfen Einträge und "
             "Default-Einstellungen ändern.")
# Gemeinsame Einstellungen für alle Benutzer; die Ansicht ("theme") speichert jeder Benutzer selbst.
MACHINE_KEYS = ("max_items", "files_dir", "setup_done", "users", "categories")
STATE = {"legacy": None, "error": False, "stamp": None, "conflict": False, "offline": False}


def _set_win_attr(path, attr):
    try:
        import ctypes
        ctypes.windll.kernel32.SetFileAttributesW(str(path), attr)
    except Exception:  # noqa: BLE001
        pass


def protect_file(path, shared=False):
    """Datei schreibgeschützt und (unter Windows) versteckt ablegen."""
    try:
        if WIN:
            _set_win_attr(path, 0x1 | 0x2)   # READONLY | HIDDEN
        else:
            os.chmod(path, 0o666 if shared else 0o400)
    except OSError:
        pass


def unprotect_file(path):
    try:
        if os.path.exists(path):
            if WIN:
                _set_win_attr(path, 0x80)    # NORMAL
            else:
                os.chmod(path, 0o600)
    except OSError:
        pass


def is_portable():
    return os.path.exists(os.path.join(app_dir(), PORTABLE_FLAG))


def is_admin():
    try:
        if WIN:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        return os.geteuid() == 0
    except Exception:  # noqa: BLE001
        return False


def _run(cmd):
    try:
        flags = 0x08000000 if WIN else 0     # CREATE_NO_WINDOW
        return subprocess.run(cmd, capture_output=True, creationflags=flags).returncode == 0
    except OSError:
        return False


def prepare_machine_dir(d):
    """Legt den gemeinsamen Ordner an und setzt die Rechte: Alle Benutzer dürfen die
    Einstellungsdatei lesen und überschreiben (wer was ändern darf, prüft die Anmeldung im
    Programm), aber nicht löschen oder ersetzen. Sonst ließe sich durch Löschen eine neue
    Ersteinrichtung mit eigenem Admin erzwingen."""
    if is_portable() or (not WIN and d not in (LINUX_SHARED, NETWORK["dir"])):
        return
    if is_unc(d):                            # Freigabe: die Rechte setzt der Server
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            pass
        return
    created = not os.path.isdir(d)
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        return
    if not (created or is_admin()):
        return
    if WIN:
        _run(["icacls", d, "/setowner", "*S-1-5-32-544", "/T", "/C", "/Q"])   # nur als Admin
        # SIDs statt Namen, damit es auch auf deutschen Systemen (Administratoren/Benutzer) passt.
        _run(["icacls", d, "/inheritance:r", "/grant:r", "*S-1-5-18:(OI)(CI)F",
              "*S-1-5-32-544:(OI)(CI)F", "*S-1-5-32-545:(RX,WD)", "/Q"])
        _run(["icacls", d, "/grant", "*S-1-5-32-545:(OI)(IO)(R,W)", "/Q"])
        _run(["icacls", os.path.join(d, "*"), "/reset", "/C", "/Q"])
        _set_win_attr(d, 0x2)                # HIDDEN
    else:
        try:
            os.chmod(d, 0o1777)              # Sticky-Bit: nur der Eigentümer darf löschen
        except OSError:
            pass


def user_dir():
    """Versteckter Ordner im Benutzerprofil für die persönliche Ansicht."""
    try:
        if WIN:
            base = os.environ.get("APPDATA") or os.path.expanduser("~")
            d = os.path.join(base, "DocuBuilder")
        else:
            base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
            d = os.path.join(base, "docubuilder")
        os.makedirs(d, exist_ok=True)
        if WIN:
            _set_win_attr(d, 0x2)            # HIDDEN
        else:
            os.chmod(d, 0o700)
        return d
    except OSError:
        return app_dir()


def is_unc(path):
    return path.startswith(("\\\\", "//"))


def read_server_file():
    """Netzwerk-Version: {"dir": gemeinsamer Ordner, "map": [(von, nach), ...]} aus server.txt.
    dir ist leer, wenn es keine server.txt gibt (Einzelplatz)."""
    net = {"dir": "", "map": [], "file": ""}
    for d in (app_dir(), LINUX_SHARED) if not WIN else (app_dir(),):
        path = os.path.join(d, SERVER_FILE)
        try:
            with open(path, encoding="utf-8-sig") as f:
                lines = [ln.strip() for ln in f if ln.strip() and not ln.lstrip().startswith("#")]
        except OSError:
            continue
        net["file"] = path
        for ln in lines:
            if "=" in ln:
                src, dst = (s.strip() for s in ln.split("=", 1))
                if src and dst:
                    net["map"].append((src, dst))
            elif not net["dir"]:
                net["dir"] = os.path.expandvars(os.path.expanduser(ln))
        break
    if net["dir"]:
        net["dir"] = map_path(net["dir"], net["map"])
    return net


def map_path(path, mapping):
    """Pfad nach den Zeilen "von = nach" aus server.txt übersetzen (z. B. \\\\Server\\Vorlagen →
    /mnt/vorlagen auf Linux-Arbeitsplätzen). Ohne passende Zeile bleibt der Pfad unverändert."""
    norm = path.replace("\\", "/")
    for src, dst in mapping:
        s = src.replace("\\", "/").rstrip("/")
        if norm.lower() == s.lower() or norm.lower().startswith(s.lower() + "/"):
            rest = norm[len(s):].lstrip("/")
            joined = dst.rstrip("/\\") + ("/" + rest if rest else "")
            return os.path.normpath(joined)
    return path


NETWORK = read_server_file()


def machine_dir():
    """Gemeinsamer Ordner für alle Benutzer. Netzwerk-Version: Ordner aus server.txt.
    Mit portable.txt neben dem Programm: Programmordner."""
    if NETWORK["dir"]:
        return NETWORK["dir"]
    if is_portable():
        return app_dir()
    if WIN:
        return os.path.join(os.environ.get("ProgramData") or r"C:\ProgramData", "DocuBuilder")
    # Unter Linux kann nur root /etc/docubuilder anlegen; bis dahin pro Benutzer.
    return LINUX_SHARED if os.path.isdir(LINUX_SHARED) or is_admin() else user_dir()


MACHINE_DIR = machine_dir()
CONFIG = os.path.join(MACHINE_DIR, ".docubuilder.dat")
PREFS = os.path.join(user_dir(), ".docubuilder-user.dat")
LEGACY_FILES = [os.path.join(user_dir(), ".docubuilder.dat"),     # 1.3: pro Benutzer
                os.path.join(app_dir(), "config.json"),
                os.path.join(os.path.expanduser("~"), ".docubuilder", "config.json")]


def _keys(salt):
    dk = hashlib.pbkdf2_hmac("sha256", _SEED, salt, 50_000, dklen=64)
    return dk[:32], dk[32:]


def _stream(key, nonce, n):
    out, counter = bytearray(), 0
    while len(out) < n:
        out += hmac.new(key, nonce + counter.to_bytes(8, "big"), hashlib.sha256).digest()
        counter += 1
    return bytes(out[:n])


def _xor(data, ks):
    return bytes(a ^ b for a, b in zip(data, ks))


def encrypt_blob(data):
    salt, nonce = os.urandom(16), os.urandom(16)
    ek, mk = _keys(salt)
    ct = _xor(data, _stream(ek, nonce, len(data)))
    mac = hmac.new(mk, MAGIC + salt + nonce + ct, hashlib.sha256).digest()
    return MAGIC + salt + nonce + mac + ct


def decrypt_blob(blob):
    if blob[:4] != MAGIC or len(blob) < 68:
        raise ValueError("Ungültiges Dateiformat")
    salt, nonce, mac, ct = blob[4:20], blob[20:36], blob[36:68], blob[68:]
    ek, mk = _keys(salt)
    if not hmac.compare_digest(mac, hmac.new(mk, MAGIC + salt + nonce + ct, hashlib.sha256).digest()):
        raise ValueError("Datei verändert oder beschädigt")
    return _xor(ct, _stream(ek, nonce, len(ct)))


def hash_password(pw):
    salt, iters = os.urandom(16), 200_000
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, iters)
    return f"pbkdf2${iters}${salt.hex()}${dk.hex()}"


def verify_password(pw, stored):
    try:
        _, iters, salt, dk = stored.split("$")
        calc = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), int(iters))
        return hmac.compare_digest(calc.hex(), dk)
    except ValueError:
        return False


def clean_settings(raw):
    s = dict(DEFAULT_SETTINGS, users=[])
    if isinstance(raw, dict):
        if raw.get("theme") in ("system", "light", "dark"):
            s["theme"] = raw["theme"]
        try:
            s["max_items"] = min(MAX_ITEMS_LIMIT, max(1, int(raw.get("max_items"))))
        except (TypeError, ValueError):
            pass
        s["files_dir"] = str(raw.get("files_dir") or "")
        s["setup_done"] = bool(raw.get("setup_done", False))
        s["categories"] = clean_categories(raw.get("categories"))
        s["users"] = clean_users(raw.get("users"), s["categories"])
        h = raw.get("admin_hash")              # bis 1.3: Passwort des Default-Fensters
        if not s["users"] and isinstance(h, str) and h.startswith("pbkdf2$"):
            s["users"] = [new_user(ADMIN_NAME, h, admin=True)]
    return s


def clean_categories(raw):
    cats, seen = [], set()
    for c in raw if isinstance(raw, list) else []:
        c = str(c).strip()
        if c and c.lower() not in seen and c != ALL_LABEL:
            seen.add(c.lower())
            cats.append(c)
    return cats


def new_user(name, pw_hash, admin=False, can_edit=False, cats=()):
    return {"name": name, "hash": pw_hash, "admin": admin, "can_edit": admin or can_edit,
            "cats": list(cats)}


def clean_users(raw, categories=()):
    """Gültige Konten; nur das Konto ADMIN_NAME ist Admin, Namen sind eindeutig.
    Konten aus Versionen ohne Bereiche sehen weiterhin alle Bereiche."""
    users, seen = [], set()
    for u in raw if isinstance(raw, list) else []:
        if not isinstance(u, dict):
            continue
        name, h = str(u.get("name", "")).strip(), u.get("hash")
        if not name or name.lower() in seen or not (isinstance(h, str) and h.startswith("pbkdf2$")):
            continue
        seen.add(name.lower())
        is_adm = name.lower() == ADMIN_NAME.lower()
        cats = u.get("cats") if isinstance(u.get("cats"), list) else categories
        cats = [c for c in categories if c in cats]
        users.append(new_user(ADMIN_NAME if is_adm else name, h, is_adm, bool(u.get("can_edit")),
                              cats))
    users.sort(key=lambda u: not u["admin"])  # Admin immer zuerst
    return users


def sees_all(user):
    """Admin und Benutzer mit „Darf ändern“ sehen alle Bereiche (sie pflegen die Einträge)."""
    return bool(user and (user["admin"] or user["can_edit"]))


def visible(item, user):
    return sees_all(user) or not item.get("cat") or item["cat"] in user["cats"]


def find_user(users, name):
    return next((u for u in users if u["name"].lower() == name.strip().lower()), None)


def has_admin(settings):
    return any(u["admin"] for u in settings["users"])


def parse_config(data):
    raw_items = data if isinstance(data, list) else data.get("items", []) if isinstance(data, dict) else []
    raw_settings = data.get("settings") if isinstance(data, dict) else None
    items = [{"name": str(x.get("name", "")), "path": str(x.get("path", "")),
              "cat": str(x.get("cat") or "")}
             for x in raw_items if isinstance(x, dict)]
    settings = clean_settings(raw_settings)
    for it in items:                           # Bereich gelöscht: Eintrag gilt wieder für alle
        if it["cat"] not in settings["categories"]:
            it["cat"] = ""
    if not (isinstance(raw_settings, dict) and "setup_done" in raw_settings):
        settings["setup_done"] = bool(items)   # ältere Dateien: Einrichtung gilt als erledigt
    return settings, items


def read_store(path, allow_plain=False):
    with open(path, "rb") as f:
        raw = f.read()
    if allow_plain and raw[:4] != MAGIC:
        return json.loads(raw.decode("utf-8-sig"))
    return json.loads(decrypt_blob(raw).decode("utf-8"))


class ConfigChanged(Exception):
    """Die gemeinsame Konfiguration wurde seit dem Einlesen an einem anderen Platz geändert."""


def file_stamp(raw):
    """Kennung eines Dateistands ("" = keine Datei)."""
    return hashlib.sha256(raw).hexdigest() if raw else ""


# Gesperrt wird ein Byte weit hinter dem Dateiende: Unter Windows sind Sperren verbindlich, eine
# Sperre am Dateianfang würde sonst alle Lesenden blockieren. So schließen sich nur Speichernde aus.
LOCK_OFFSET = 0x7FFFFFF0


def _lock(f, tries=50):
    for _ in range(tries):
        try:
            if WIN:
                import msvcrt
                f.seek(LOCK_OFFSET)
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.lockf(f, fcntl.LOCK_EX | fcntl.LOCK_NB, 1, LOCK_OFFSET)
            return True
        except OSError:
            time.sleep(0.1)
    return False


def _unlock(f):
    try:
        if WIN:
            import msvcrt
            f.seek(LOCK_OFFSET)
            msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.lockf(f, fcntl.LOCK_UN, 1, LOCK_OFFSET)
    except OSError:
        pass


def write_store(path, obj, shared=False, expect=None):
    """Speichert verschlüsselt. shared: gemeinsame Datei; mit expect (Kennung beim Einlesen)
    wird ConfigChanged ausgelöst, wenn inzwischen jemand anderes gespeichert hat.
    Liefert die Kennung des neuen Stands."""
    data = encrypt_blob(json.dumps(obj, ensure_ascii=False).encode("utf-8"))
    if shared:
        # Gemeinsame Datei direkt überschreiben: Benutzer dürfen sie ändern, aber nicht
        # löschen oder ersetzen (siehe prepare_machine_dir).
        unprotect_file(path)
        try:
            with open(path, "r+b" if os.path.exists(path) else "x+b") as f:
                if not _lock(f):
                    raise OSError("Die Einstellungen werden gerade an einem anderen Platz "
                                  "gespeichert. Bitte gleich noch einmal versuchen.")
                try:
                    f.seek(0)
                    if expect is not None and file_stamp(f.read()) != expect:
                        raise ConfigChanged()
                    f.seek(0)
                    f.write(data)
                    f.truncate()
                    f.flush()
                finally:
                    _unlock(f)
        finally:
            protect_file(path, shared)
        return file_stamp(data)
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    unprotect_file(path)
    try:
        os.replace(tmp, path)
    finally:
        protect_file(path, shared)
    return file_stamp(data)


def server_offline():
    """Netzwerk-Version: gemeinsamer Ordner nicht erreichbar (Server aus, kein Netz, keine Rechte)."""
    return bool(NETWORK["dir"]) and not os.path.isdir(MACHINE_DIR)


def read_config_raw():
    with open(CONFIG, "rb") as f:
        return f.read()


def load_config():
    """Liest die gemeinsame Konfiguration. Ältere Dateien (1.3, config.json) werden übernommen."""
    STATE.update(error=False, conflict=False, offline=server_offline())
    if STATE["offline"]:
        return clean_settings(None), []
    if os.path.exists(CONFIG):
        # Mehrere Versuche: im Netzwerk kann gerade ein anderer Platz die Datei schreiben.
        for _ in range(5):
            try:
                raw = read_config_raw()
                data = json.loads(decrypt_blob(raw).decode("utf-8"))
                STATE["stamp"] = file_stamp(raw)
                return parse_config(data)
            except (OSError, ValueError):
                time.sleep(0.2)
        STATE["error"] = True
        return clean_settings(None), []
    STATE["stamp"] = ""
    if NETWORK["dir"]:                         # ältere Einzelplatz-Dateien nicht ins Netz übernehmen
        return clean_settings(None), []
    for path in LEGACY_FILES:
        if os.path.isfile(path):
            try:
                data = read_store(path, allow_plain=True)
                STATE["legacy"] = path
                return parse_config(data)
            except (OSError, ValueError):
                continue
    return clean_settings(None), []


def load_prefs(settings):
    """Übernimmt die persönlichen Einstellungen (Ansicht, Farbschema, Update-Suche) in settings."""
    try:
        prefs = read_store(PREFS)
        theme = prefs.get("theme")
    except (OSError, ValueError, AttributeError):
        return
    if theme in ("system", "light", "dark"):
        settings["theme"] = theme
    if prefs.get("scheme") in COLOR_SCHEMES:
        settings["scheme"] = prefs["scheme"]
    settings["update_check"] = bool(prefs.get("update_check", False))


def save_config(settings, items):
    try:
        STATE["stamp"] = write_store(
            CONFIG, {"settings": {k: settings[k] for k in MACHINE_KEYS}, "items": items},
            shared=True, expect=STATE["stamp"])
        return True
    except ConfigChanged:
        STATE["conflict"] = True
        messagebox.showwarning(APP_NAME, "Die Einstellungen wurden inzwischen an einem anderen "
                                         "Platz geändert. Deine Änderung wurde nicht gespeichert.\n\n"
                                         "DocuBuilder lädt jetzt den aktuellen Stand. Bitte die "
                                         "Änderung danach noch einmal vornehmen.")
    except PermissionError:
        if NETWORK["dir"]:
            messagebox.showerror(APP_NAME, "Keine Schreibrechte für die gemeinsame Konfiguration:\n"
                                 + CONFIG + "\n\nBitte die Rechte der Freigabe auf dem Server "
                                 "prüfen (siehe README, Netzwerk-Version).")
        else:
            messagebox.showerror(APP_NAME, "Windows erlaubt das Speichern nicht:\n" + CONFIG +
                                 "\n\nBitte DocuBuilder einmal als Administrator starten "
                                 "(Rechtsklick → Als Administrator ausführen), damit die Rechte "
                                 "gesetzt werden.")
    except OSError as e:
        messagebox.showerror(APP_NAME, f"Einstellungen konnten nicht gespeichert werden:\n{e}")
    return False


def _remove(path):
    """Datei oder Ordner löschen (Schreibschutz vorher aufheben). True, wenn danach weg."""
    if not os.path.exists(path):
        return True
    try:
        if os.path.isdir(path):
            for root, _, files in os.walk(path):
                for f in files:
                    unprotect_file(os.path.join(root, f))
            shutil.rmtree(path)
        else:
            unprotect_file(path)
            os.remove(path)
    except OSError:
        pass
    return not os.path.exists(path)


def _own_dir(d):
    """Nur eigene DocuBuilder-Ordner ganz löschen, nie z. B. den Programmordner."""
    return os.path.basename(os.path.normpath(d)).lower() == "docubuilder" and \
        os.path.normcase(os.path.abspath(d)) != os.path.normcase(app_dir())


def _run_elevated(params):
    """cmd.exe mit Administratorrechten (UAC-Abfrage) ausführen und warten. False bei Ablehnung."""
    import ctypes
    from ctypes import wintypes

    class ShellExecuteInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("fMask", ctypes.c_ulong),
                    ("hwnd", wintypes.HWND), ("lpVerb", wintypes.LPCWSTR),
                    ("lpFile", wintypes.LPCWSTR), ("lpParameters", wintypes.LPCWSTR),
                    ("lpDirectory", wintypes.LPCWSTR), ("nShow", ctypes.c_int),
                    ("hInstApp", wintypes.HINSTANCE), ("lpIDList", ctypes.c_void_p),
                    ("lpClass", wintypes.LPCWSTR), ("hkeyClass", wintypes.HKEY),
                    ("dwHotKey", wintypes.DWORD), ("hIcon", wintypes.HANDLE),
                    ("hProcess", wintypes.HANDLE)]

    info = ShellExecuteInfo(cbSize=ctypes.sizeof(ShellExecuteInfo), fMask=0x40,  # NOCLOSEPROCESS
                            lpVerb="runas", lpFile="cmd.exe", lpParameters=params, nShow=0)
    if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(info)):
        return False
    if info.hProcess:
        ctypes.windll.kernel32.WaitForSingleObject(info.hProcess, 60_000)
        ctypes.windll.kernel32.CloseHandle(info.hProcess)
    return True


def remove_all_data():
    """Löscht die gemeinsame Konfiguration (Einträge, Benutzer) und die eigene Ansicht.
    Liefert None bei Erfolg, sonst einen Fehlertext."""
    user = os.path.dirname(PREFS)
    for p in [user] if _own_dir(user) else [PREFS, PREFS + ".tmp"] + LEGACY_FILES:
        _remove(p)
    # Netzwerk-Version: nur die Konfigurationsdatei, nie den Ordner auf dem Server
    target = MACHINE_DIR if _own_dir(MACHINE_DIR) and not NETWORK["dir"] else CONFIG
    if _remove(target):
        return None
    if NETWORK["dir"]:
        return ("Die gemeinsame Konfiguration auf dem Server konnte nicht gelöscht werden. Bitte "
                f"auf dem Server löschen lassen:\n{target}")
    if WIN:
        if os.path.isdir(target):
            cmd = f'/c attrib -R -H "{target}\\*" & rmdir /s /q "{target}"'
        else:
            cmd = f'/c attrib -R -H "{target}" & del /f /q "{target}"'
        if _run_elevated(cmd) and not os.path.exists(target):
            return None
        if os.path.exists(CONFIG):
            protect_file(CONFIG, shared=True)
        return ("Die gemeinsame Konfiguration konnte nicht gelöscht werden, weil keine "
                f"Administratorrechte erteilt wurden:\n{target}")
    return f"Bitte mit Administratorrechten löschen:\nsudo rm -rf {target}"


def save_prefs(settings):
    try:
        write_store(PREFS, {"theme": settings["theme"], "scheme": settings["scheme"],
                            "update_check": settings["update_check"]})
    except OSError as e:
        messagebox.showerror(APP_NAME, f"Ansicht konnte nicht gespeichert werden:\n{e}")


# ---------------------------------------------------------------- Update
GITHUB_REPO = "plange-dev/DocuBuilder"
RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases?per_page=30"
RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases"
DOWNLOAD_PREFIX = f"https://github.com/{GITHUB_REPO}/releases/download/"


def parse_version(text):
    m = re.search(r"(\d+)\.(\d+)\.(\d+)\.(\d+)$", text or "")
    return tuple(int(x) for x in m.groups()) if m else None


def fmt_version(v):
    return ".".join(map(str, v))


def _http(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}",
                                               "Accept": "application/vnd.github+json"})
    return urllib.request.urlopen(req, timeout=timeout)


def find_update(channel=APP_CHANNEL, current=APP_VERSION, releases=None):
    """Neuestes Release, das neuer ist als die laufende Version, sonst None.
    Stable sieht nur Stable-Releases, Nightly zusätzlich die Vorabversionen."""
    if releases is None:
        with _http(RELEASES_API) as r:
            releases = json.loads(r.read().decode("utf-8"))
    cur = parse_version(current)
    best = None
    for rel in releases:
        if rel.get("draft") or (rel.get("prerelease") and channel != "nightly"):
            continue
        v = parse_version(rel.get("tag_name"))
        if v and v > cur and (best is None or v > best["version"]):
            best = {"version": v, "nightly": bool(rel.get("prerelease")),
                    "notes": (rel.get("body") or "").strip(),
                    "page": rel.get("html_url") or RELEASES_PAGE,
                    "assets": {a["name"]: a["browser_download_url"] for a in rel.get("assets", [])}}
    return best


def install_mode():
    """setup: mit dem Installationsprogramm installiert; replace: einzelne Programmdatei;
    source: als Python-Skript gestartet (kein automatisches Update)."""
    if not getattr(sys, "frozen", False):
        return "source"
    if WIN and os.path.exists(os.path.join(app_dir(), "unins000.exe")):
        return "setup"
    return "replace"


def update_asset(update, mode):
    if mode == "setup":
        return f"DocuBuilder-Setup-{fmt_version(update['version'])}.exe"
    return "DocuBuilder.exe" if WIN else "DocuBuilder"


def download_update(update, name, dest_dir, progress=None):
    """Lädt eine Datei des Releases und prüft sie gegen <name>.sha256. Liefert den Pfad."""
    url, sha_url = update["assets"].get(name), update["assets"].get(name + ".sha256")
    if not url or not sha_url:
        raise ValueError(f"Im Release fehlt die Datei {name} oder ihre Prüfsumme.")
    for u in (url, sha_url):
        if not u.startswith(DOWNLOAD_PREFIX):
            raise ValueError(f"Unerwartete Download-Adresse:\n{u}")
    with _http(sha_url) as r:
        expected = r.read().decode("ascii", "replace").split()[0].lower()
    path = os.path.join(dest_dir, name)
    digest = hashlib.sha256()
    with _http(url, timeout=60) as r, open(path, "wb") as f:
        total, done = int(r.headers.get("Content-Length") or 0), 0
        while True:
            chunk = r.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
            digest.update(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    if digest.hexdigest() != expected:
        os.remove(path)
        raise ValueError("Die Prüfsumme der heruntergeladenen Datei stimmt nicht.\n"
                         "Das Update wurde verworfen.")
    return path


def clean_child_env():
    """PyInstaller-Variablen entfernen, bevor ein neuer Prozess startet (Setup oder neue .exe).
    Sonst hält sich das neue DocuBuilder für einen Unterprozess dieser .exe und sucht seine Dateien
    in deren temporärem Ordner _MEI…, der beim Beenden gelöscht wird („Failed to load Python DLL“)."""
    mei = os.path.normcase(getattr(sys, "_MEIPASS", "") or "")
    for key in list(os.environ):
        value = os.path.normcase(os.environ[key])
        if key.startswith("_PYI_") or key == "_MEIPASS2" or (mei and value.startswith(mei)):
            del os.environ[key]


def program_pid():
    """Prozess, der beendet sein muss, bevor die Programmdatei ersetzt werden kann. Bei einer
    einzelnen PyInstaller-.exe ist das der Startprozess (Elternprozess dieses Python-Prozesses)."""
    mei = getattr(sys, "_MEIPASS", None)
    exe_dir = os.path.normcase(os.path.dirname(sys.executable))
    # Einzeldatei: _MEIPASS ist ein Temp-Ordner. Programmordner (--onedir): _MEIPASS liegt im
    # Programmordner (_internal), es gibt keinen Startprozess.
    if getattr(sys, "frozen", False) and mei and not os.path.normcase(mei).startswith(exe_dir):
        return os.getppid()
    return os.getpid()


def start_setup(path):
    """Installationsprogramm ohne Rückfragen starten; Windows fragt nach Adminrechten (UAC).
    /WAITPID: das Setup wartet, bis dieses DocuBuilder beendet ist, bevor es Dateien ersetzt.
    /UPDATE=1: das Setup startet DocuBuilder danach wieder (siehe installer/DocuBuilder.iss)."""
    os.startfile(path, "open", "/SILENT /SP- /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS "
                               f"/UPDATE=1 /WAITPID={program_pid()}")


def install_error_text(e):
    code = getattr(e, "winerror", None)
    if code == 1223:
        return "Die Administratorrechte wurden nicht erteilt. Das Update wurde nicht installiert."
    if code == 4551:                                 # Smart App Control / Anwendungssteuerung
        return ("Windows hat das Update blockiert (Smart App Control bzw. eine "
                "Anwendungssteuerungsrichtlinie), weil die Datei nicht digital signiert ist.\n\n"
                "Bis DocuBuilder signiert ausgeliefert wird, lässt sich das Update nur installieren, "
                "wenn Smart App Control ausgeschaltet ist (Windows-Sicherheit → App- & "
                "Browsersteuerung). Das ist eine Sicherheitsentscheidung – im Zweifel bitte die IT fragen.")
    return f"Das Update konnte nicht installiert werden:\n{e}"


def replace_program(exe, new_file):
    """Laufende Programmdatei austauschen: Windows erlaubt Umbenennen, aber nicht Überschreiben.
    Die alte Datei (.old) wird beim nächsten Start entfernt."""
    old = exe + ".old"
    if os.path.exists(old):
        os.remove(old)
    os.replace(exe, old)
    try:
        shutil.move(new_file, exe)
    except OSError:
        os.replace(old, exe)
        raise
    if not WIN:
        os.chmod(exe, 0o755)


def cleanup_old_program():
    if getattr(sys, "frozen", False):
        try:
            os.remove(sys.executable + ".old")
        except OSError:
            pass


# ---------------------------------------------------------------- Öffnen
def open_target(path, base=""):
    path = path.strip()
    if not path:
        messagebox.showinfo(APP_NAME, "Für diesen Button ist kein Pfad hinterlegt.")
        return
    if re.match(r"^https?://", path, re.I):
        webbrowser.open(path)
        return
    full = map_path(os.path.expandvars(os.path.expanduser(path)), NETWORK["map"])
    if base and not (os.path.isabs(full) or is_unc(full)):
        # relative Pfade ab dem Dateiordner (Default)
        full = os.path.join(map_path(base, NETWORK["map"]), full)
    if not os.path.exists(full):
        messagebox.showerror(APP_NAME, f"Datei nicht gefunden:\n{full}")
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(full)  # öffnet mit dem Standardprogramm (Excel)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", full])
        else:
            subprocess.Popen(["xdg-open", full], stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
    except Exception as e:  # noqa: BLE001
        messagebox.showerror(APP_NAME, f"Datei konnte nicht geöffnet werden:\n{e}")


def list_excel_files(folder):
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    files = [os.path.join(folder, n) for n in names
             if os.path.splitext(n)[1].lower() in EXCEL_EXT and not n.startswith("~$")
             and os.path.isfile(os.path.join(folder, n))]
    return sorted(files, key=lambda f: os.path.basename(f).lower())


def fill_from_folder(items, folder, n):
    """Trägt Excel-Dateien aus dem Ordner in leere Plätze bzw. am Ende ein (max. n Einträge)."""
    existing = {os.path.normcase(os.path.normpath(i["path"])) for i in items if i["path"]}
    added = 0
    for path in list_excel_files(folder):
        key = os.path.normcase(os.path.normpath(path))
        if key in existing:
            continue
        entry = {"name": os.path.splitext(os.path.basename(path))[0], "path": path, "cat": ""}
        slot = next((i for i, it in enumerate(items) if not it["path"]), None)
        if slot is not None:
            items[slot] = entry
        elif len(items) < n:
            items.append(entry)
        else:
            break
        existing.add(key)
        added += 1
    return added


# ---------------------------------------------------------------- Design
def detect_system_theme():
    """Liefert 'dark' oder 'light' passend zum Betriebssystem."""
    try:
        if sys.platform.startswith("win"):
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
                return "light" if winreg.QueryValueEx(k, "AppsUseLightTheme")[0] else "dark"
        for key in ("color-scheme", "gtk-theme"):
            out = subprocess.run(["gsettings", "get", "org.gnome.desktop.interface", key],
                                 capture_output=True, text=True, timeout=2).stdout.lower()
            if "dark" in out:
                return "dark"
        if "dark" in os.environ.get("GTK_THEME", "").lower():
            return "dark"
    except Exception:  # noqa: BLE001
        pass
    return "light"


def set_dark_titlebar(win, dark):
    """Dunkle Titelleiste unter Windows 10/11 (rein kosmetisch, Fehler werden ignoriert)."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), 4)
    except Exception:  # noqa: BLE001
        pass


def setup_style(root, p):
    s = ttk.Style(root)
    s.theme_use("clam")
    root.configure(bg=p["bg"])
    # Hover-, Druck- und Sperrzustände explizit setzen: "clam" färbt sie sonst hellgrau,
    # was im Dunkelmodus kaum lesbar ist.
    dis = [("disabled", p["disabled"])]
    s.configure(".", background=p["bg"], foreground=p["fg"], fieldbackground=p["panel"],
                bordercolor=p["line"], lightcolor=p["line"], darkcolor=p["line"],
                troughcolor=p["bg"], focuscolor=p["accent"], selectbackground=p["accent"],
                selectforeground=p["accent_fg"], insertcolor=p["fg"])
    s.map(".", foreground=dis)
    s.configure("TFrame", background=p["bg"])
    s.configure("TLabel", background=p["bg"], foreground=p["fg"])
    s.configure("Mute.TLabel", foreground=p["mute"])
    s.configure("Title.TLabel", font=(None, 15, "bold"), foreground=p.get("title", p["fg"]))
    s.configure("TLabelframe", background=p["bg"], bordercolor=p["line"])
    s.configure("TLabelframe.Label", background=p["bg"], foreground=p["fg"])
    s.configure("TButton", background=p["btn"], foreground=p["fg"], bordercolor=p["line"],
                lightcolor=p["btn"], darkcolor=p["btn"], padding=(10, 5))
    s.map("TButton", foreground=dis,
          background=[("disabled", p["bg"]), ("pressed", p["btn_hover"]), ("active", p["btn_hover"])],
          bordercolor=[("focus", p["accent"])],
          lightcolor=[("disabled", p["bg"]), ("active", p["btn_hover"])],
          darkcolor=[("disabled", p["bg"]), ("active", p["btn_hover"])])
    # Hauptknopf (z. B. "Anmelden"); eigene Farbe optional, sonst die Akzentfarbe
    prim, prim_fg = p.get("primary", p["accent"]), p.get("primary_fg", p["accent_fg"])
    s.configure("Accent.TButton", background=prim, foreground=prim_fg,
                bordercolor=prim, lightcolor=prim, darkcolor=prim)
    s.map("Accent.TButton", foreground=dis,
          background=[("disabled", p["btn"]), ("active", prim)],
          lightcolor=[("disabled", p["btn"]), ("active", prim)],
          darkcolor=[("disabled", p["btn"]), ("active", prim)])
    field = dict(fieldbackground=p["panel"], foreground=p["fg"], insertcolor=p["fg"],
                 bordercolor=p["line"], lightcolor=p["panel"], darkcolor=p["panel"],
                 selectbackground=p["accent"], selectforeground=p["accent_fg"])
    field_map = dict(foreground=dis, fieldbackground=[("disabled", p["bg"]), ("readonly", p["bg"])],
                     bordercolor=[("focus", p["accent"])], lightcolor=[("focus", p["accent"])])
    s.configure("TEntry", **field)
    s.map("TEntry", **field_map)
    s.configure("TSpinbox", background=p["btn"], arrowcolor=p["fg"], **field)
    s.map("TSpinbox", background=[("active", p["btn_hover"])], arrowcolor=dis, **field_map)
    s.configure("TCombobox", background=p["btn"], arrowcolor=p["fg"], **field)
    s.map("TCombobox", background=[("active", p["btn_hover"])], arrowcolor=dis,
          **dict(field_map, fieldbackground=[("disabled", p["bg"]), ("readonly", p["panel"])]))
    root.option_add("*TCombobox*Listbox.background", p["panel"])
    root.option_add("*TCombobox*Listbox.foreground", p["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", p["accent"])
    root.option_add("*TCombobox*Listbox.selectForeground", p["accent_fg"])
    for kind in ("TRadiobutton", "TCheckbutton"):
        s.configure(kind, background=p["bg"], foreground=p["fg"], indicatorbackground=p["panel"],
                    indicatorforeground=p["accent_fg"], upperbordercolor=p["line"],
                    lowerbordercolor=p["line"])
        s.map(kind, background=[("active", p["bg"])], foreground=dis,
              indicatorbackground=[("disabled", p["bg"]), ("selected", p["accent"]),
                                   ("pressed", p["btn_hover"])],
              upperbordercolor=[("selected", p["accent"]), ("active", p["accent"])],
              lowerbordercolor=[("selected", p["accent"]), ("active", p["accent"])])
    s.configure("TNotebook", background=p["bg"], bordercolor=p["line"],
                lightcolor=p["panel"], darkcolor=p["line"])
    s.configure("TNotebook.Tab", background=p["tab"], foreground=p["mute"], bordercolor=p["line"],
                lightcolor=p["tab"], padding=(14, 6))
    s.map("TNotebook.Tab", background=[("selected", p["bg"]), ("active", p["btn_hover"])],
          foreground=[("selected", p["fg"]), ("active", p["fg"])],
          lightcolor=[("selected", p["accent"])])
    s.configure("Treeview", background=p["panel"], fieldbackground=p["panel"],
                foreground=p["fg"], bordercolor=p["line"], rowheight=26)
    s.map("Treeview", background=[("selected", p["accent"])],
          foreground=[("selected", p["accent_fg"])])
    s.configure("Treeview.Heading", background=p["btn"], foreground=p["fg"], bordercolor=p["line"],
                lightcolor=p["btn"], darkcolor=p["btn"])
    s.map("Treeview.Heading", background=[("active", p["btn_hover"])],
          lightcolor=[("active", p["btn_hover"])], darkcolor=[("active", p["btn_hover"])])
    s.configure("TScrollbar", background=p["btn"], troughcolor=p["bg"], arrowcolor=p["fg"],
                bordercolor=p["line"], lightcolor=p["btn"], darkcolor=p["btn"])
    s.map("TScrollbar", background=[("pressed", p["accent"]), ("active", p["btn_hover"])],
          arrowcolor=dis)


# ---------------------------------------------------------------- Einstellungen
class Settings(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Einstellungen")
        self.geometry("760x620")
        self.minsize(620, 440)
        self.configure(bg=app.pal["bg"])
        self.transient(app)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        self.nb = nb = ttk.Notebook(outer)
        nb.grid(row=0, column=0, sticky="nsew")
        self.tab_items = ttk.Frame(nb, padding=12)
        self.tab_general = ttk.Frame(nb, padding=16)
        nb.add(self.tab_items, text="Einträge")
        nb.add(self.tab_general, text="Allgemein")
        ttk.Button(outer, text="Schließen", command=self.destroy).grid(
            row=1, column=0, sticky="e", pady=(12, 0))

        self.build_items_tab()
        self.build_general_tab()

    # ---- Tab Einträge
    def build_items_tab(self):
        f = self.tab_items
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(f, columns=("name", "cat", "path"), show="headings",
                                 selectmode="browse", height=8)
        self.tree.heading("name", text="Name")
        self.tree.heading("cat", text="Bereich")
        self.tree.heading("path", text="Pfad / Link")
        self.tree.column("name", width=160)
        self.tree.column("cat", width=130)
        self.tree.column("path", width=280)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        edit = []                                 # wird ohne Schreibrecht gesperrt
        side = ttk.Frame(f)
        side.grid(row=0, column=2, sticky="n", padx=(10, 0))
        for text, cmd in (("Nach oben", lambda: self.move(-1)),
                          ("Nach unten", lambda: self.move(1)),
                          ("Löschen", self.delete)):
            edit.append(ttk.Button(side, text=text, command=cmd, width=12))
            edit[-1].pack(pady=2)
        self.counter = ttk.Label(side, style="Mute.TLabel")
        self.counter.pack(pady=(14, 0))

        form = ttk.LabelFrame(f, text="Eintrag", padding=10)
        form.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        form.columnconfigure(1, weight=1)
        self.name_var = tk.StringVar()
        self.path_var = tk.StringVar()
        ttk.Label(form, text="Name:").grid(row=0, column=0, sticky="w", pady=3)
        edit.append(ttk.Entry(form, textvariable=self.name_var))
        edit[-1].grid(row=0, column=1, columnspan=2, sticky="ew", pady=3)
        ttk.Label(form, text="Pfad / Link:").grid(row=1, column=0, sticky="w", pady=3)
        edit.append(ttk.Entry(form, textvariable=self.path_var))
        edit[-1].grid(row=1, column=1, sticky="ew", pady=3)
        edit.append(ttk.Button(form, text="Durchsuchen …", command=self.browse))
        edit[-1].grid(row=1, column=2, padx=(6, 0))
        ttk.Label(form, text="Bereich:").grid(row=2, column=0, sticky="w", pady=3)
        self.cat_var = tk.StringVar(value=ALL_LABEL)
        edit.append(ttk.Combobox(form, textvariable=self.cat_var, state="readonly", width=28,
                                 values=[ALL_LABEL] + self.app.settings["categories"]))
        edit[-1].grid(row=2, column=1, sticky="w", pady=3)
        btns = ttk.Frame(form)
        btns.grid(row=3, column=0, columnspan=3, sticky="e", pady=(8, 0))
        edit.append(ttk.Button(btns, text="Neuer Eintrag", command=self.new))
        edit[-1].pack(side="left", padx=4)
        edit.append(ttk.Button(btns, text="Speichern", style="Accent.TButton",
                               command=self.save_entry))
        edit[-1].pack(side="left")
        if not self.app.can_edit:
            for w in edit:
                w.state(["disabled"])
            ttk.Label(f, text=NO_RIGHTS, style="Mute.TLabel", wraplength=600,
                      justify="left").grid(row=2, column=0, columnspan=3, sticky="w",
                                           pady=(10, 0))
        self.refresh()

    def refresh(self, select=None):
        items = self.app.items
        self.tree.delete(*self.tree.get_children())
        for i, it in enumerate(items):
            if visible(it, self.app.user):           # fremde Bereiche auch hier nicht zeigen
                self.tree.insert("", "end", iid=str(i),
                                 values=(it["name"], it["cat"] or ALL_LABEL, it["path"]))
        if select is not None and self.tree.exists(str(select)):
            self.tree.selection_set(str(select))
            self.tree.see(str(select))
        self.counter.configure(text=f"{len(items)} von {self.app.settings['max_items']}")

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def on_select(self, _event=None):
        i = self.selected()
        if i is not None:
            self.name_var.set(self.app.items[i]["name"])
            self.path_var.set(self.app.items[i]["path"])
            self.cat_var.set(self.app.items[i]["cat"] or ALL_LABEL)

    def new(self):
        self.tree.selection_remove(self.tree.selection())
        self.name_var.set("")
        self.path_var.set("")
        self.cat_var.set(ALL_LABEL)

    def browse(self):
        p = filedialog.askopenfilename(
            parent=self, title="Excel-Datei wählen",
            initialdir=self.app.settings["files_dir"] or None,
            filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls *.xlsb *.xltx *.xltm *.csv *.ods"),
                       ("Alle Dateien", "*.*")])
        if p:
            self.path_var.set(os.path.normpath(p))
            if not self.name_var.get().strip():
                self.name_var.set(os.path.splitext(os.path.basename(p))[0])

    def save_entry(self):
        items = self.app.items
        name, path = self.name_var.get().strip(), self.path_var.get().strip()
        cat = self.cat_var.get() if self.cat_var.get() in self.app.settings["categories"] else ""
        if not name or not path:
            messagebox.showinfo(APP_NAME, "Bitte Name und Pfad eintragen.", parent=self)
            return
        i = self.selected()
        if i is None:
            if len(items) >= self.app.settings["max_items"]:
                messagebox.showinfo(
                    APP_NAME, "Die maximale Anzahl an Einträgen ist erreicht.\n"
                              "Sie ist in den Default-Einstellungen festgelegt.", parent=self)
                return
            items.append({"name": name, "path": path, "cat": cat})
            i = len(items) - 1
        else:
            items[i] = {"name": name, "path": path, "cat": cat}
        self.commit(i)

    def delete(self):
        i = self.selected()
        if i is None:
            return
        if messagebox.askyesno(APP_NAME, f"„{self.app.items[i]['name']}“ löschen?", parent=self):
            del self.app.items[i]
            self.new()
            self.commit(None)

    def move(self, d):
        items = self.app.items
        i = self.selected()
        if i is None or not 0 <= i + d < len(items):
            return
        items[i], items[i + d] = items[i + d], items[i]
        self.commit(i + d)

    def commit(self, select):
        self.app.save()
        self.app.build(force=True)
        self.refresh(select)

    # ---- Tab Allgemein
    def build_general_tab(self):
        f = self.tab_general
        f.columnconfigure(0, weight=1)

        view = ttk.LabelFrame(f, text="Ansicht", padding=12)
        view.grid(row=0, column=0, sticky="ew")
        view.columnconfigure(1, weight=1)
        modes = ttk.Frame(view)
        modes.grid(row=0, column=0, sticky="nw", padx=(0, 40))
        ttk.Label(modes, text="Hell / Dunkel", style="Mute.TLabel").pack(anchor="w", pady=(0, 2))
        self.theme_var = tk.StringVar(value=self.app.settings["theme"])
        for value, text in (("system", "An das System anpassen"),
                            ("light", "Hell"), ("dark", "Dunkel")):
            ttk.Radiobutton(modes, text=text, value=value, variable=self.theme_var,
                            command=self.apply_theme_choice).pack(anchor="w", pady=2)
        schemes = ttk.Frame(view)
        schemes.grid(row=0, column=1, sticky="nw")
        ttk.Label(schemes, text="Farbschema", style="Mute.TLabel").pack(anchor="w", pady=(0, 2))
        self.scheme_var = tk.StringVar(value=self.app.settings["scheme"])
        for value, (text, _) in COLOR_SCHEMES.items():
            ttk.Radiobutton(schemes, text=text + (" (Standard)" if value == DEFAULT_SCHEME else ""),
                            value=value, variable=self.scheme_var,
                            command=self.apply_theme_choice).pack(anchor="w", pady=2)
        ttk.Button(view, text="Ansicht auf Standard zurücksetzen",
                   command=self.reset_defaults).grid(row=1, column=0, sticky="w", pady=(8, 0))

        upd = ttk.LabelFrame(f, text="Updates", padding=12)
        upd.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        if UPDATES_ENABLED:
            self.upd_var = tk.BooleanVar(value=self.app.settings["update_check"])
            ttk.Checkbutton(upd, text="Beim Anmelden nach Updates suchen (fragt GitHub ab)",
                            variable=self.upd_var, command=self.apply_update_choice).pack(anchor="w")
            ttk.Button(upd, text="Jetzt nach Updates suchen …",
                       command=self.app.open_update).pack(anchor="w", pady=(8, 0))
        else:
            ttk.Label(upd, style="Mute.TLabel", wraplength=560, justify="left",
                      text="Die Update-Funktion wird freigeschaltet, sobald DocuBuilder "
                           "signiert ausgeliefert wird.").pack(anchor="w")

        if self.app.user["admin"]:
            box = ttk.LabelFrame(f, text="Daten entfernen (nur Admin)", padding=10)
            box.grid(row=2, column=0, sticky="ew", pady=(14, 0))
            text = ("Löscht alle Einträge, Benutzer und Passwörter auf dem Server – für alle "
                    "Arbeitsplätze – sowie deine Ansicht. Danach wird DocuBuilder beendet."
                    if NETWORK["dir"] else
                    "Löscht alle Einträge, Benutzer und Passwörter dieses Rechners sowie "
                    "deine Ansicht. Danach wird DocuBuilder beendet und kann gelöscht "
                    "werden. Windows fragt nach Administratorrechten.")
            ttk.Label(box, style="Mute.TLabel", wraplength=560, justify="left",
                      text=text).pack(anchor="w")
            ttk.Button(box, text="DocuBuilder-Daten entfernen …",
                       command=self.remove_data).pack(anchor="w", pady=(8, 0))

    def apply_update_choice(self):
        self.app.settings["update_check"] = self.upd_var.get()
        save_prefs(self.app.settings)

    def apply_theme_choice(self):
        self.app.settings["theme"] = self.theme_var.get()
        self.app.settings["scheme"] = self.scheme_var.get()
        save_prefs(self.app.settings)
        self.app.apply_theme()
        self.app.build(force=True)

    def reset_defaults(self):
        if not messagebox.askyesno(APP_NAME, "Ansicht auf Standard zurücksetzen?", parent=self):
            return
        self.app.settings["theme"] = DEFAULT_SETTINGS["theme"]
        self.app.settings["scheme"] = DEFAULT_SETTINGS["scheme"]
        self.theme_var.set(self.app.settings["theme"])
        self.scheme_var.set(self.app.settings["scheme"])
        save_prefs(self.app.settings)
        self.app.apply_theme()
        self.app.build(force=True)

    # ---- Daten entfernen (nur Admin, im Tab Allgemein)
    def remove_data(self):
        if not messagebox.askyesno(
                APP_NAME, "Wirklich ALLE DocuBuilder-Daten auf diesem Rechner löschen?\n\n"
                          "Einträge, Benutzer und Passwörter gehen verloren. Das kann nicht "
                          "rückgängig gemacht werden.", icon="warning", default="no", parent=self):
            return
        pw = simpledialog.askstring(APP_NAME, f"Zur Bestätigung das {ADMIN_NAME}-Passwort eingeben:",
                                    show="•", parent=self)
        if pw is None:
            return
        admin = next((u for u in self.app.settings["users"] if u["admin"]), None)
        if not (admin and verify_password(pw, admin["hash"])):
            messagebox.showerror(APP_NAME, "Falsches Passwort. Es wurde nichts gelöscht.", parent=self)
            return
        err = remove_all_data()
        if err:
            messagebox.showerror(APP_NAME, err, parent=self)
            return
        messagebox.showinfo(
            APP_NAME, "Alle DocuBuilder-Daten wurden entfernt.\n\nDocuBuilder wird jetzt beendet. "
                      "Danach kannst du die Programmdatei löschen.\n\nAndere Windows-Benutzer "
                      "können ihre Ansicht unter %APPDATA%\\DocuBuilder selbst löschen.", parent=self)
        self.app.destroy()


# ---------------------------------------------------------------- Info
class InfoDialog(tk.Toplevel):
    """Eigenes Fenster „Info zu DocuBuilder“ (aus dem Menü)."""

    def __init__(self, app):
        super().__init__(app)
        self.title("Info zu DocuBuilder")
        self.resizable(False, False)
        self.configure(bg=app.pal["bg"])
        self.transient(app)
        set_dark_titlebar(self, app.is_dark)
        self.bind("<Escape>", lambda e: self.destroy())

        f = ttk.Frame(self, padding=20)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, columnspan=2,
                                                               sticky="w")
        ttk.Label(f, text="Buttons, die hinterlegte Excel-Listen öffnen.",
                  style="Mute.TLabel").grid(row=1, column=0, columnspan=2, sticky="w",
                                            pady=(0, 14))
        user = app.user
        rows = (("Version", APP_VERSION + (" (Nightly)" if APP_CHANNEL == "nightly" else "")),
                ("Lizenz", "MIT"),
                ("Betrieb", f"Netzwerk (Server: {MACHINE_DIR})" if NETWORK["dir"] else "Einzelplatz"),
                ("Konfiguration", CONFIG),
                ("Angemeldet als", user["name"] + (
                    " (Admin)" if user["admin"] else
                    " (darf ändern)" if app.can_edit else " (nur ansehen)")),
                ("Bereiche", "alle" if sees_all(user) else ", ".join(user["cats"]) or "keine"),
                ("System", platform.platform()),
                ("Python", platform.python_version()))
        for i, (k, v) in enumerate(rows, start=2):
            ttk.Label(f, text=k + ":", style="Mute.TLabel").grid(row=i, column=0, sticky="nw",
                                                                 padx=(0, 12), pady=2)
            ttk.Label(f, text=v, wraplength=420, justify="left").grid(row=i, column=1,
                                                                       sticky="w", pady=2)
        ttk.Button(f, text="Schließen", style="Accent.TButton", command=self.destroy).grid(
            row=len(rows) + 2, column=0, columnspan=2, sticky="e", pady=(18, 0))


# ---------------------------------------------------------------- Default
def check_password(p1, p2):
    """Fehlertext oder None."""
    if p1 != p2:
        return "Die Passwörter stimmen nicht überein."
    if len(p1) < MIN_PW_LEN:
        return f"Das Passwort braucht mindestens {MIN_PW_LEN} Zeichen."
    return None


def category_checks(parent, cats, chosen):
    """Häkchen je Bereich; liefert eine Funktion, die die gewählten Bereiche zurückgibt."""
    vars_ = []
    if not cats:
        ttk.Label(parent, style="Mute.TLabel", text="Es sind noch keine Bereiche angelegt.").pack(
            anchor="w")
    for c in cats:
        v = tk.BooleanVar(value=c in chosen)
        ttk.Checkbutton(parent, text=c, variable=v).pack(anchor="w", pady=1)
        vars_.append((c, v))
    return lambda: [c for c, v in vars_ if v.get()]


class AccessDialog(tk.Toplevel):
    """Sichtbare Bereiche eines Benutzers festlegen."""

    def __init__(self, parent, app, user, cats):
        super().__init__(parent)
        self.result = None
        self.title("Bereiche")
        self.configure(bg=app.pal["bg"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, wraplength=320, justify="left",
                  text=f"Welche Bereiche sieht „{user['name']}“?").pack(anchor="w", pady=(0, 8))
        self.get = category_checks(f, cats, user["cats"])
        ttk.Label(f, style="Mute.TLabel", wraplength=320, justify="left",
                  text="Einträge ohne Bereich sieht jeder. Mit „Darf ändern“ sieht ein Benutzer "
                       "ohnehin alle Bereiche.").pack(anchor="w", pady=(10, 0))
        row = ttk.Frame(f)
        row.pack(anchor="e", pady=(12, 0))
        ttk.Button(row, text="Abbrechen", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(row, text="OK", style="Accent.TButton", command=self.ok).pack(side="left")
        self.bind("<Escape>", lambda ev: self.destroy())

    def ok(self):
        self.result = self.get()
        self.destroy()


class CategoryDialog(tk.Toplevel):
    """Bereiche anlegen, umbenennen, löschen. Ergebnis: [(Name, Name beim Öffnen oder None)]."""

    def __init__(self, parent, app, cats):
        super().__init__(parent)
        self.result = None
        self.rows = [(c, c) for c in cats]
        self.title("Bereiche verwalten")
        self.configure(bg=app.pal["bg"])
        self.geometry("460x380")
        self.transient(parent)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)
        ttk.Label(f, style="Mute.TLabel", wraplength=420, justify="left",
                  text="Bereiche fassen Einträge zusammen, z. B. Arbeitsvorbereitung, Angebote, "
                       "Abrechnung. Pro Benutzer legst du fest, welche Bereiche er sieht.").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self.tree = ttk.Treeview(f, columns=("name",), show="headings", selectmode="browse")
        self.tree.heading("name", text="Bereich")
        self.tree.grid(row=1, column=0, sticky="nsew")
        side = ttk.Frame(f)
        side.grid(row=1, column=1, sticky="n", padx=(10, 0))
        for text, cmd in (("Hinzufügen …", self.add), ("Umbenennen …", self.rename),
                          ("Nach oben", lambda: self.move(-1)), ("Nach unten", lambda: self.move(1)),
                          ("Löschen", self.delete)):
            ttk.Button(side, text=text, command=cmd, width=14).pack(pady=2)
        row = ttk.Frame(f)
        row.grid(row=2, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(row, text="Abbrechen", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(row, text="OK", style="Accent.TButton", command=self.ok).pack(side="left")
        self.bind("<Escape>", lambda ev: self.destroy())
        self.refresh()

    def refresh(self, select=None):
        self.tree.delete(*self.tree.get_children())
        for i, (name, _) in enumerate(self.rows):
            self.tree.insert("", "end", iid=str(i), values=(name,))
        if select is not None and 0 <= select < len(self.rows):
            self.tree.selection_set(str(select))

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def ask_name(self, title, initial=""):
        name = simpledialog.askstring(title, "Name des Bereichs:", initialvalue=initial, parent=self)
        name = (name or "").strip()
        if not name:
            return None
        if name == ALL_LABEL or len(name) > 40:
            messagebox.showinfo(APP_NAME, "Bitte einen anderen Namen (max. 40 Zeichen) wählen.",
                                parent=self)
            return None
        if any(n.lower() == name.lower() and n != initial for n, _ in self.rows):
            messagebox.showinfo(APP_NAME, "Diesen Bereich gibt es schon.", parent=self)
            return None
        return name

    def add(self):
        name = self.ask_name("Bereich hinzufügen")
        if name:
            self.rows.append((name, None))
            self.refresh(len(self.rows) - 1)

    def rename(self):
        i = self.selected()
        if i is None:
            return
        name = self.ask_name("Bereich umbenennen", self.rows[i][0])
        if name:
            self.rows[i] = (name, self.rows[i][1])
            self.refresh(i)

    def move(self, d):
        i = self.selected()
        if i is None or not 0 <= i + d < len(self.rows):
            return
        self.rows[i], self.rows[i + d] = self.rows[i + d], self.rows[i]
        self.refresh(i + d)

    def delete(self):
        i = self.selected()
        if i is not None and messagebox.askyesno(
                APP_NAME, f"Bereich „{self.rows[i][0]}“ löschen?\n\nSeine Einträge sind danach für "
                          "alle sichtbar.", parent=self):
            del self.rows[i]
            self.refresh()

    def ok(self):
        self.result = self.rows
        self.destroy()


class UserDialog(tk.Toplevel):
    """Benutzer anlegen (user=None) oder Passwort eines Benutzers neu setzen."""

    def __init__(self, parent, app, taken, user=None, cats=()):
        super().__init__(parent)
        self.result, self.taken, self.user = None, taken, user
        self.get_cats = lambda: []
        self.title("Passwort setzen" if user else "Benutzer anlegen")
        self.configure(bg=app.pal["bg"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        self.name_var, self.pw1, self.pw2 = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.edit_var = tk.BooleanVar(value=False)
        ttk.Label(f, text="Benutzername:").grid(row=0, column=0, sticky="w", pady=4)
        if user:
            ttk.Label(f, text=user["name"]).grid(row=0, column=1, sticky="w", padx=(8, 0))
        else:
            e = ttk.Entry(f, textvariable=self.name_var, width=30)
            e.grid(row=0, column=1, sticky="ew", pady=4, padx=(8, 0))
            e.focus_set()
        ttk.Label(f, text="Passwort:").grid(row=1, column=0, sticky="w", pady=4)
        pw = ttk.Entry(f, textvariable=self.pw1, show="•", width=30)
        pw.grid(row=1, column=1, sticky="ew", pady=4, padx=(8, 0))
        if user:
            pw.focus_set()
        ttk.Label(f, text="Wiederholen:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.pw2, show="•", width=30).grid(row=2, column=1, sticky="ew",
                                                                     pady=4, padx=(8, 0))
        if not user:
            ttk.Checkbutton(f, text="Darf ändern (Einträge und Default-Fenster)",
                            variable=self.edit_var).grid(row=3, column=0, columnspan=2,
                                                         sticky="w", pady=(6, 0))
            box = ttk.LabelFrame(f, text="Sichtbare Bereiche", padding=8)
            box.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
            self.get_cats = category_checks(box, cats, ())
        self.msg = ttk.Label(f, wraplength=320, justify="left")
        self.msg.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        row = ttk.Frame(f)
        row.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(row, text="Abbrechen", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(row, text="OK", style="Accent.TButton", command=self.ok).pack(side="left")
        self.bind("<Return>", lambda ev: self.ok())
        self.bind("<Escape>", lambda ev: self.destroy())

    def ok(self):
        name = self.user["name"] if self.user else self.name_var.get().strip()
        if not self.user:
            if not name or len(name) > 40:
                return self.msg.configure(text="Bitte einen Benutzernamen (max. 40 Zeichen) eingeben.")
            if name.lower() in self.taken:
                return self.msg.configure(text="Diesen Benutzernamen gibt es schon.")
        err = check_password(self.pw1.get(), self.pw2.get())
        if err:
            return self.msg.configure(text=err)
        self.result = new_user(name, hash_password(self.pw1.get()),
                               can_edit=self.user["can_edit"] if self.user else self.edit_var.get(),
                               cats=self.user["cats"] if self.user else self.get_cats())
        self.destroy()


class DefaultDialog(tk.Toplevel):
    """Anzahl der Buttons, Dateiordner, Benutzer. Erststart automatisch, danach per ALT+F11."""

    def __init__(self, app, first_run=False):
        super().__init__(app)
        self.app = app
        self.first_run = first_run
        self.manage_users = first_run or app.user["admin"]
        self.title("Default – Erste Einrichtung" if first_run else "Default")
        self.geometry("720x820" if self.manage_users else "640x480")
        self.minsize(600, 700 if self.manage_users else 420)
        self.configure(bg=app.pal["bg"])
        self.transient(app)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Default", style="Title.TLabel").grid(row=0, column=0, columnspan=3,
                                                                sticky="w")
        ttk.Label(f, style="Mute.TLabel", wraplength=580, justify="left",
                  text="Lege fest, wie viele Buttons es gibt und in welchem Ordner deine Dateien "
                       "liegen. Dieses Fenster erscheint nur beim ersten Start. Später öffnest du "
                       "es mit ALT+F11 (alternativ STRG+UMSCHALT+F11).").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(4, 16))

        ttk.Label(f, text="Anzahl der Buttons:").grid(row=2, column=0, sticky="w", pady=6)
        self.count_var = tk.StringVar(value=str(app.settings["max_items"]))
        ttk.Spinbox(f, from_=1, to=MAX_ITEMS_LIMIT, width=6, textvariable=self.count_var).grid(
            row=2, column=1, sticky="w", pady=6)

        ttk.Label(f, text="Ordner mit den Dateien:").grid(row=3, column=0, sticky="w", pady=6)
        self.dir_var = tk.StringVar(value=app.settings["files_dir"])
        ttk.Entry(f, textvariable=self.dir_var).grid(row=3, column=1, sticky="ew", pady=6)
        ttk.Button(f, text="Durchsuchen …", command=self.browse).grid(row=3, column=2,
                                                                       padx=(6, 0))

        self.import_var = tk.BooleanVar(value=first_run)
        ttk.Checkbutton(f, text="Excel-Dateien aus dem Ordner als Buttons übernehmen",
                        variable=self.import_var).grid(row=4, column=1, columnspan=2,
                                                       sticky="w", pady=(2, 0))

        self.pw1, self.pw2 = tk.StringVar(), tk.StringVar()
        admin = next((u for u in app.settings["users"] if u["admin"]), None)
        self.admin_hash = admin["hash"] if admin else ""
        self.users = [dict(u, cats=list(u["cats"])) for u in app.settings["users"] if not u["admin"]]
        self.cats = list(app.settings["categories"] or (DEFAULT_CATEGORIES if first_run else []))
        self.cat_origin = {c: c for c in app.settings["categories"]}   # aktueller → alter Name
        if self.manage_users:
            self.build_users(f)
        else:
            ttk.Label(f, style="Mute.TLabel", text="Benutzer verwaltet nur der Admin.").grid(
                row=5, column=0, columnspan=3, sticky="w", pady=(16, 0))

        ttk.Label(f, style="Mute.TLabel", wraplength=580, justify="left",
                  text="Nicht belegte Buttons heißen „Auswahl 1, 2, …“ und können unter "
                       "Einstellungen → Einträge zugeordnet werden. Relative Pfade werden ab dem "
                       "Ordner gesucht.").grid(row=6, column=0, columnspan=3, sticky="w",
                                               pady=(14, 0))
        self.msg = ttk.Label(f, wraplength=580, justify="left")
        self.msg.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))

        btns = ttk.Frame(f)
        btns.grid(row=8, column=0, columnspan=3, sticky="se", pady=(16, 0))
        f.rowconfigure(8, weight=1)
        ttk.Button(btns, text="Abbrechen", command=self.cancel).pack(side="left", padx=6)
        ttk.Button(btns, text="Übernehmen", style="Accent.TButton",
                   command=self.apply).pack(side="left")

    def build_users(self, f):
        box = ttk.LabelFrame(f, text="Benutzer", padding=10)
        box.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(16, 0))
        f.rowconfigure(5, weight=1)
        box.columnconfigure(1, weight=1)
        box.rowconfigure(3, weight=1)
        ttk.Label(box, style="Mute.TLabel", wraplength=560, justify="left", text=(
            f"Das Konto „{ADMIN_NAME}“ besteht immer und darf alles. Bitte ein Passwort festlegen."
            if self.first_run else
            f"Passwort für „{ADMIN_NAME}“ ändern (leer lassen = unverändert):")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(box, text=f"{ADMIN_NAME}-Passwort:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(box, textvariable=self.pw1, show="•").grid(row=1, column=1, columnspan=2,
                                                             sticky="ew", pady=3, padx=(8, 0))
        ttk.Label(box, text="Wiederholen:").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(box, textvariable=self.pw2, show="•").grid(row=2, column=1, columnspan=2,
                                                             sticky="ew", pady=3, padx=(8, 0))

        self.utree = ttk.Treeview(box, columns=("name", "edit", "cats"), show="headings",
                                  selectmode="browse", height=5)
        self.utree.heading("name", text="Benutzer")
        self.utree.heading("edit", text="Darf ändern")
        self.utree.heading("cats", text="Sichtbare Bereiche")
        self.utree.column("name", width=150)
        self.utree.column("edit", width=90, anchor="center")
        self.utree.column("cats", width=220)
        self.utree.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        side = ttk.Frame(box)
        side.grid(row=3, column=2, sticky="n", padx=(10, 0), pady=(10, 0))
        for text, cmd in (("Hinzufügen …", self.add_user),
                          ("Passwort …", self.reset_password),
                          ("Recht an/aus", self.toggle_edit),
                          ("Bereiche …", self.edit_access),
                          ("Löschen", self.delete_user)):
            ttk.Button(side, text=text, command=cmd, width=14).pack(pady=2)
        cats = ttk.Frame(box)
        cats.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        cats.columnconfigure(0, weight=1)
        self.cats_label = ttk.Label(cats, style="Mute.TLabel", wraplength=440, justify="left")
        self.cats_label.grid(row=0, column=0, sticky="w")
        ttk.Button(cats, text="Bereiche verwalten …", command=self.manage_categories).grid(
            row=0, column=1, sticky="e", padx=(10, 0))
        self.refresh_users()

    def user_cats_text(self, u):
        if u["can_edit"]:
            return "alle (darf ändern)"
        return ", ".join(u["cats"]) if u["cats"] else "nur Einträge ohne Bereich"

    def refresh_users(self, select=None):
        self.utree.delete(*self.utree.get_children())
        self.utree.insert("", "end", iid="admin", values=(f"{ADMIN_NAME} (fest)", "ja", "alle"))
        for i, u in enumerate(self.users):
            self.utree.insert("", "end", iid=str(i), values=(
                u["name"], "ja" if u["can_edit"] else "nein", self.user_cats_text(u)))
        if select is not None:
            self.utree.selection_set(select)
        self.cats_label.configure(text="Bereiche: " + (", ".join(self.cats) or "keine"))

    def manage_categories(self):
        dlg = CategoryDialog(self, self.app, self.cats)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        local = {old: new for new, old in dlg.result if old}
        for u in self.users:
            u["cats"] = [local[c] for c in u["cats"] if c in local]
        self.cat_origin = {new: self.cat_origin.get(old) for new, old in dlg.result}
        self.cats = [new for new, _ in dlg.result]
        self.refresh_users()

    def edit_access(self):
        i = self.selected_user()
        if i is None:
            return
        dlg = AccessDialog(self, self.app, self.users[i], self.cats)
        self.wait_window(dlg)
        if dlg.result is not None:
            self.users[i]["cats"] = dlg.result
            self.refresh_users(str(i))

    def selected_user(self):
        sel = self.utree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Bitte einen Benutzer auswählen.", parent=self)
            return None
        if sel[0] == "admin":
            messagebox.showinfo(APP_NAME, f"Das Konto „{ADMIN_NAME}“ kann nicht gelöscht oder "
                                          "eingeschränkt werden. Sein Passwort änderst du oben.",
                                parent=self)
            return None
        return int(sel[0])

    def add_user(self):
        taken = {ADMIN_NAME.lower()} | {u["name"].lower() for u in self.users}
        dlg = UserDialog(self, self.app, taken, cats=self.cats)
        self.wait_window(dlg)
        if dlg.result:
            self.users.append(dlg.result)
            self.refresh_users(str(len(self.users) - 1))

    def reset_password(self):
        i = self.selected_user()
        if i is None:
            return
        dlg = UserDialog(self, self.app, set(), user=self.users[i])
        self.wait_window(dlg)
        if dlg.result:
            self.users[i] = dlg.result
            self.refresh_users(str(i))

    def toggle_edit(self):
        i = self.selected_user()
        if i is not None:
            self.users[i]["can_edit"] = not self.users[i]["can_edit"]
            self.refresh_users(str(i))

    def delete_user(self):
        i = self.selected_user()
        if i is not None and messagebox.askyesno(
                APP_NAME, f"Benutzer „{self.users[i]['name']}“ löschen?", parent=self):
            del self.users[i]
            self.refresh_users()

    def browse(self):
        d = filedialog.askdirectory(parent=self, title="Ordner mit den Dateien wählen",
                                    initialdir=self.dir_var.get() or None)
        if d:
            self.dir_var.set(os.path.normpath(d))

    def cancel(self):
        if self.first_run and not messagebox.askyesno(
                APP_NAME, "Ohne Einrichtung kann DocuBuilder nicht benutzt werden.\n"
                          "Programm beenden?", parent=self):
            return
        self.grab_release()
        self.destroy()
        if self.first_run:
            self.app.destroy()

    def apply(self):
        try:
            n = int(self.count_var.get())
        except ValueError:
            n = 0
        if not 1 <= n <= MAX_ITEMS_LIMIT:
            self.msg.configure(text=f"Bitte eine Zahl von 1 bis {MAX_ITEMS_LIMIT} eingeben.")
            return
        folder = self.dir_var.get().strip()
        if folder:
            folder = os.path.normpath(os.path.expandvars(os.path.expanduser(folder)))
            if not os.path.isdir(map_path(folder, NETWORK["map"])):
                self.msg.configure(text="Der Ordner wurde nicht gefunden.")
                return
        elif self.import_var.get():
            self.msg.configure(text="Zum Einlesen bitte einen Ordner angeben.")
            return

        users = self.app.settings["users"]
        cats = self.app.settings["categories"]
        if self.manage_users:
            cats = self.cats
            p1, p2 = self.pw1.get(), self.pw2.get()
            admin_hash = self.admin_hash
            if p1 or p2 or not admin_hash:
                err = check_password(p1, p2)
                if err:
                    self.msg.configure(text=f"{ADMIN_NAME}-Passwort: {err}")
                    return
                admin_hash = hash_password(p1)
            users = [new_user(ADMIN_NAME, admin_hash, admin=True, cats=cats)] + self.users

        items = self.app.items
        backup = [dict(it) for it in items]      # bei Fehlschlag beim Speichern zurücksetzen
        if len(items) > n:
            if not messagebox.askyesno(
                    APP_NAME, f"Es gibt {len(items)} Einträge, erlaubt sind {n}.\n"
                              "Die überzähligen Einträge am Ende werden entfernt. Fortfahren?",
                    parent=self):
                return
            del items[n:]
        imported = fill_from_folder(items, folder, n) if self.import_var.get() else 0
        total = len(list_excel_files(folder)) if self.import_var.get() else 0
        while len(items) < n:
            items.append({"name": f"Auswahl {len(items) + 1}", "path": "", "cat": ""})
        if self.manage_users:                    # umbenannte Bereiche übernehmen, gelöschte leeren
            renamed = {old: new for new, old in self.cat_origin.items() if old}
            for it in items:
                it["cat"] = renamed.get(it["cat"], "")

        old_settings = dict(self.app.settings)
        self.app.settings.update(max_items=n, files_dir=folder, setup_done=True, users=users,
                                 categories=cats)
        if self.first_run:
            self.app.user = users[0]             # nach der Einrichtung als Admin angemeldet
        if not self.app.save():
            if self.first_run:
                self.app.user = None
            if STATE["conflict"]:                # aktueller Stand wurde neu geladen
                self.grab_release()
                self.destroy()
            else:
                self.app.settings.clear()
                self.app.settings.update(old_settings)
                items[:] = backup
            return
        self.grab_release()
        self.destroy()
        if self.first_run:
            self.app.show_main()
        else:
            self.app.build(force=True)
        if self.import_var.get():
            messagebox.showinfo(APP_NAME, f"{imported} Datei(en) übernommen "
                                          f"(im Ordner liegen {total}).", parent=self.app)


# ---------------------------------------------------------------- Update-Fenster
class UpdateDialog(tk.Toplevel):
    """Sucht nach Updates und installiert sie. Netzwerkzugriffe laufen im Hintergrund."""

    def __init__(self, app, found=None):
        super().__init__(app)
        self.app, self.found, self.job, self.progress = app, None, None, (0, 0)
        p = app.pal
        self.title("Update")
        self.minsize(520, 200)                   # Größe richtet sich nach dem Inhalt
        self.configure(bg=p["bg"])
        self.transient(app)
        # Nur modal, wenn es aus den (modalen) Einstellungen kommt; sonst kann man während des
        # Downloads im Hauptfenster weiterarbeiten.
        self.modal = app._alive(app.settings_win)
        if self.modal:
            self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        self.protocol("WM_DELETE_WINDOW", self.close)

        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=1)
        self.head = ttk.Label(f, style="Title.TLabel")
        self.head.grid(row=0, column=0, sticky="w")
        self.info = ttk.Label(f, wraplength=530, justify="left")
        self.info.grid(row=1, column=0, sticky="w", pady=(8, 10))
        self.notes = tk.Text(f, height=8, wrap="word", relief="flat", bd=0, highlightthickness=1,
                             bg=p["panel"], fg=p["fg"], highlightbackground=p["line"],
                             padx=8, pady=6)
        self.notes.grid(row=2, column=0, sticky="nsew")
        self.notes.grid_remove()
        self.bar = ttk.Progressbar(f, length=520)
        self.bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.bar.grid_remove()
        self.btns = ttk.Frame(f)
        self.btns.grid(row=4, column=0, sticky="e", pady=(14, 0))
        pending = app.pending_update
        if pending and os.path.exists(pending["path"]):
            self.show_ready()                    # schon heruntergeladen: nur noch installieren
        elif found:
            self.show_update(found)
        else:
            self.check()

    def run(self, work, done):
        """work() im Hintergrund ausführen, danach done(ergebnis, fehler) im Tk-Thread."""
        box = {}

        def target():
            try:
                box["result"] = work()
            except Exception as e:  # noqa: BLE001 – Fehler werden im Fenster angezeigt
                box["error"] = e

        self.job = threading.Thread(target=target, daemon=True)
        self.job.start()

        def poll():
            if not self.winfo_exists():
                return
            if self.job.is_alive():
                self.after(100, poll)
            else:
                done(box.get("result"), box.get("error"))
        self.after(100, poll)

    def set_buttons(self, *buttons):
        """buttons: (Text, Befehl, hervorgehoben)"""
        for w in self.btns.winfo_children():
            w.destroy()
        for text, cmd, accent in buttons:
            ttk.Button(self.btns, text=text, command=cmd,
                       style="Accent.TButton" if accent else "TButton").pack(side="left", padx=(6, 0))

    def check(self):
        self.head.configure(text="Nach Updates suchen")
        self.info.configure(text="Verbindung zu GitHub …")
        self.notes.grid_remove()
        self.bar.configure(mode="indeterminate")
        self.bar.grid()
        self.bar.start(12)
        self.set_buttons(("Schließen", self.close, False))
        self.run(find_update, self.checked)

    def checked(self, found, error):
        self.bar.stop()
        self.bar.grid_remove()
        if error:
            self.head.configure(text="Keine Verbindung")
            self.info.configure(text=f"Die Suche nach Updates ist fehlgeschlagen:\n{error}")
            self.set_buttons(("Erneut versuchen", self.check, False), ("Schließen", self.close, True))
        elif not found:
            channel = " (Nightly)" if APP_CHANNEL == "nightly" else ""
            self.head.configure(text="DocuBuilder ist aktuell")
            self.info.configure(text=f"Installiert ist die neueste Version {APP_VERSION}{channel}.")
            self.set_buttons(("Schließen", self.close, True))
        else:
            self.show_update(found)

    def show_update(self, found):
        self.found = found
        v = fmt_version(found["version"])
        mode = install_mode()
        is_admin_user = bool(self.app.user and self.app.user["admin"])
        lines = [f"Installiert: {APP_VERSION}   →   neu: {v}" + (" (Nightly)" if found["nightly"] else "")]
        if mode == "source":
            lines.append("DocuBuilder läuft als Python-Skript. Bitte die neue Version über die "
                         "Release-Seite holen.")
        elif not is_admin_user:
            lines.append(f"Installieren darf nur der {ADMIN_NAME}.")
        else:
            lines.append("Das Update wird zuerst heruntergeladen und geprüft. Zum Installieren muss "
                         "DocuBuilder danach geschlossen werden.")
        self.head.configure(text=f"Version {v} ist verfügbar")
        self.info.configure(text="\n\n".join(lines))
        if found["notes"]:
            self.notes.configure(state="normal")
            self.notes.delete("1.0", "end")
            self.notes.insert("1.0", found["notes"])
            self.notes.configure(state="disabled")
            self.notes.grid()
        buttons = [("Release-Seite", lambda: webbrowser.open(found["page"]), False),
                   ("Später", self.close, False)]
        if mode != "source" and is_admin_user:
            buttons.append(("Herunterladen", self.download, True))
        self.set_buttons(*buttons)

    def download(self):
        mode = install_mode()
        name = update_asset(self.found, mode)
        self.head.configure(text="Update wird heruntergeladen")
        self.info.configure(text=name)
        self.notes.grid_remove()
        self.bar.configure(mode="determinate", maximum=100, value=0)
        self.bar.grid()
        self.set_buttons()
        self.progress = (0, 0)

        def progress(done, total):               # läuft im Hintergrund-Thread: nur Werte merken
            self.progress = (done, total)

        tmp = tempfile.mkdtemp(prefix="docubuilder-update-")
        self.run(lambda: download_update(self.found, name, tmp, progress),
                 lambda path, error: self.downloaded(path, error, mode))
        self.tick(name)

    def tick(self, name):
        if not self.winfo_exists() or not self.job.is_alive():
            return
        done, total = self.progress
        if total:
            self.bar.configure(value=done * 100 / total)
            self.info.configure(text=f"{name}\n{done / 1e6:.1f} von {total / 1e6:.1f} MB")
        self.after(150, lambda: self.tick(name))

    def downloaded(self, path, error, mode):
        if error:
            self.bar.grid_remove()
            self.head.configure(text="Download fehlgeschlagen")
            self.info.configure(text=str(error))
            self.set_buttons(("Erneut versuchen", self.download, False), ("Schließen", self.close, True))
            return
        self.app.pending_update = {"version": self.found["version"], "path": path, "mode": mode}
        self.show_ready()

    def show_ready(self):
        """Heruntergeladen und geprüft: Installation erst nach dem Schließen von DocuBuilder."""
        pending = self.app.pending_update
        v = fmt_version(pending["version"])
        self.bar.grid_remove()
        self.notes.grid_remove()
        self.head.configure(text=f"Update {v} ist bereit")
        uac = " Windows fragt dabei nach Administratorrechten." if pending["mode"] == "setup" else ""
        self.info.configure(text=(
            "Das Update wurde heruntergeladen, die Prüfsumme ist korrekt.\n\n"
            f"Zum Installieren muss DocuBuilder geschlossen werden.{uac} Danach startet DocuBuilder "
            "automatisch neu.\n\n"
            "Du kannst auch später installieren – beim Beenden von DocuBuilder wirst du noch einmal "
            "gefragt."))
        self.set_buttons(("Später", self.close, False),
                         ("DocuBuilder schließen und installieren", self.install_now, True))

    def install_now(self):
        v = fmt_version(self.app.pending_update["version"])
        if not messagebox.askokcancel(
                "Update installieren",
                f"DocuBuilder wird jetzt geschlossen, damit Version {v} installiert werden kann.\n\n"
                "Bitte vorher offene Eingaben abschließen. DocuBuilder startet nach dem Update "
                "automatisch neu.", icon="warning", parent=self):
            return
        try:
            self.app.install_update()
        except OSError as e:
            self.head.configure(text="Update nicht installiert")
            self.info.configure(text=install_error_text(e))
            self.set_buttons(("Erneut versuchen", self.install_now, False), ("Schließen", self.close, True))

    def close(self):
        if self.modal:
            self.grab_release()
        self.destroy()
        if self.app._alive(self.app.settings_win):
            self.app.settings_win.grab_set()     # Einstellungen waren vorher modal


# ---------------------------------------------------------------- Hauptfenster
class App(tk.Tk):
    TILE_W = 220

    def __init__(self):
        if WIN:
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
            except Exception:  # noqa: BLE001
                pass
        super().__init__()
        self.set_icon()
        self.title(APP_TITLE)
        self.geometry("720x420")
        self.minsize(360, 240)
        prepare_machine_dir(MACHINE_DIR)
        self.settings, self.items = load_config()
        load_prefs(self.settings)
        self.user = None                         # angemeldeter Benutzer (dict aus settings["users"])
        self.settings_win = None
        self.default_win = None
        self.update_win = None
        self.info_win = None
        self.pending_update = None               # heruntergeladenes, geprüftes Update (wartet aufs Schließen)
        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        self._update_checked = False            # automatische Suche nur einmal je Programmstart
        self._watching = False                  # Netzwerk-Version: Änderungen anderer übernehmen
        self.ui_font = tkfont.nametofont("TkDefaultFont").actual("family")
        self.pal = COLOR_SCHEMES[DEFAULT_SCHEME][1]["light"]
        self.is_dark = False
        self.apply_theme()

        top = ttk.Frame(self, padding=(14, 12))
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, style="Title.TLabel").pack(side="left")
        self.session = ttk.Frame(top)            # nur sichtbar, wenn jemand angemeldet ist
        self.menu_btn = ttk.Button(self.session, text="Menü ▾", command=self.show_menu)
        self.menu_btn.pack(side="right", padx=(6, 0))
        ttk.Button(self.session, text="Einstellungen", command=self.open_settings).pack(side="right")
        ttk.Button(self.session, text="Abmelden", command=self.logout).pack(side="right", padx=6)
        self.who = ttk.Label(self.session, style="Mute.TLabel")
        self.who.pack(side="right", padx=6)

        self.area = ttk.Frame(self, padding=(14, 0, 14, 14))
        self.cols = 0
        self.area.bind("<Configure>", lambda e: self.build())
        self.login = self.build_login()

        for key in DEFAULT_SHORTCUTS:
            self.bind_all(key, lambda e: self.open_default() or "break")
        self.after(200, self.startup)

    def set_icon(self):
        """Fenster- und Taskleisten-Icon, gilt auch für alle Dialoge."""
        try:
            if WIN:
                self.iconbitmap(default=resource("docubuilder.ico"))
            else:
                self._icon = tk.PhotoImage(file=resource("docubuilder.png"))
                self.iconphoto(True, self._icon)
        except (tk.TclError, OSError):
            pass

    @property
    def can_edit(self):
        return bool(self.user and self.user["can_edit"])

    @staticmethod
    def _alive(win):
        return win is not None and win.winfo_exists()

    # ---- Anmeldung
    def build_login(self):
        outer = ttk.Frame(self, padding=14)
        f = ttk.Frame(outer)
        f.place(relx=0.5, rely=0.4, anchor="center")
        ttk.Label(f, text="Anmelden", style="Title.TLabel").grid(row=0, column=0, columnspan=2,
                                                                 sticky="w", pady=(0, 10))
        self.login_name, self.login_pw = tk.StringVar(), tk.StringVar()
        ttk.Label(f, text="Benutzer:").grid(row=1, column=0, sticky="w", pady=4)
        self.login_entry = ttk.Entry(f, textvariable=self.login_name, width=28)
        self.login_entry.grid(row=1, column=1, pady=4, padx=(8, 0))
        ttk.Label(f, text="Passwort:").grid(row=2, column=0, sticky="w", pady=4)
        pw = ttk.Entry(f, textvariable=self.login_pw, show="•", width=28)
        pw.grid(row=2, column=1, pady=4, padx=(8, 0))
        self.login_btn = ttk.Button(f, text="Anmelden", style="Accent.TButton",
                                    command=self.do_login)
        self.login_btn.grid(row=3, column=1, sticky="e", pady=(10, 0))
        self.login_msg = ttk.Label(f, style="Mute.TLabel", wraplength=320, justify="left")
        self.login_msg.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        for w in (self.login_entry, pw):
            w.bind("<Return>", lambda e: self.do_login())
        return outer

    def show_login(self, msg=""):
        for win in (self.settings_win, self.default_win, self.update_win, self.info_win):
            if self._alive(win):
                win.destroy()
        self.user = None
        self.session.pack_forget()
        self.area.pack_forget()
        self.login.pack(fill="both", expand=True)
        self.login_pw.set("")
        self.login_msg.configure(text=msg)
        self.login_entry.focus_set()

    def do_login(self):
        if self.login_btn.instate(["disabled"]):
            return
        # frisch einlesen, damit Änderungen anderer Benutzer gelten
        self.settings, self.items = load_config()
        load_prefs(self.settings)
        if STATE["offline"]:
            return self.login_msg.configure(text=OFFLINE_TEXT)
        if STATE["error"]:
            return self.locked()
        if not has_admin(self.settings):         # Server war beim Start nicht erreichbar
            return self.startup()
        user = find_user(self.settings["users"], self.login_name.get())
        if not (user and verify_password(self.login_pw.get(), user["hash"])):
            self.login_pw.set("")
            self.login_msg.configure(text="Benutzername oder Passwort ist falsch.")
            self.login_btn.state(["disabled"])   # kurze Sperre gegen schnelles Durchprobieren
            self.after(1500, lambda: self.login_btn.state(["!disabled"]))
            return
        self.user = user
        if STATE["legacy"] and self.can_edit:   # ältere Einstellungen dauerhaft übernehmen
            self.save()
        self.show_main()

    def show_main(self):
        self.login.pack_forget()
        self.who.configure(text=f"Angemeldet: {self.user['name']}")
        self.session.pack(side="right")
        self.area.pack(fill="both", expand=True)
        self.apply_theme()
        self.build(force=True)
        if UPDATES_ENABLED and self.settings["update_check"] and not self._update_checked:
            self._update_checked = True
            self.auto_check_updates()

    def auto_check_updates(self):
        """Stille Suche im Hintergrund; meldet sich nur, wenn es ein Update gibt."""
        box = {}

        def target():
            try:
                box["found"] = find_update()
            except Exception:  # noqa: BLE001 – ohne Netz einfach nichts melden
                pass

        job = threading.Thread(target=target, daemon=True)
        job.start()

        def poll():
            if job.is_alive():
                self.after(500, poll)
            elif box.get("found") and self.user and not self._alive(self.update_win):
                self.open_update(box["found"])
        self.after(500, poll)

    def logout(self):
        self.login_name.set("")
        self.show_login()

    def locked(self):
        who = "Der Server-Administrator" if NETWORK["dir"] else "Ein Windows-Administrator"
        messagebox.showerror(
            APP_NAME, "Die Einstellungsdatei ist beschädigt oder wurde außerhalb von DocuBuilder "
                      f"verändert:\n{CONFIG}\n\n{who} muss die Datei löschen. "
                      "Danach startet DocuBuilder mit der Ersteinrichtung.", parent=self)
        self.destroy()

    # ---- Menü (Knopf oben rechts)
    def menu_entries(self):
        """Einträge des Menüs: (Text, Befehl) oder None für eine Trennlinie.
        Befehl None = noch nicht verfügbar (wird ausgegraut angezeigt)."""
        return [
            ("Nach Updates suchen …", self.open_update if UPDATES_ENABLED else None),
            None,
            ("Info zu DocuBuilder", self.open_info),
            None,
            ("Beenden", self.quit_app),
        ]

    def show_menu(self):
        p = self.pal
        menu = tk.Menu(self, tearoff=False, bg=p["panel"], fg=p["fg"], activebackground=p["accent"],
                       activeforeground=p["accent_fg"], disabledforeground=p["mute"])
        for entry in self.menu_entries():
            if entry is None:
                menu.add_separator()
            else:
                text, cmd = entry
                if cmd:
                    menu.add_command(label=text, command=cmd)
                else:
                    # Nicht verfügbar: selbst grau zeichnen. Windows zeichnet gesperrte Einträge
                    # mit hellem Prägeeffekt, der im Dunkelmodus wie hervorgehoben wirkt.
                    menu.add_command(label=text, foreground=p["disabled"],
                                     activeforeground=p["disabled"], activebackground=p["panel"])
        b = self.menu_btn
        menu.update_idletasks()
        x = b.winfo_rootx() + b.winfo_width() - menu.winfo_reqwidth()   # rechtsbündig zum Knopf
        try:
            menu.tk_popup(max(x, 0), b.winfo_rooty() + b.winfo_height())
        finally:
            menu.grab_release()

    # ---- Update installieren / Beenden
    def install_update(self):
        """Heruntergeladenes Update installieren. DocuBuilder wird dafür beendet; das Setup wartet,
        bis der Prozess weg ist. Wirft OSError, wenn Windows den Start verweigert."""
        pending = self.pending_update
        clean_child_env()
        if pending["mode"] == "setup":
            start_setup(pending["path"])
        else:
            replace_program(sys.executable, pending["path"])
            subprocess.Popen([sys.executable])
        self.pending_update = None
        self.destroy()

    def quit_app(self):
        """Beenden; liegt ein heruntergeladenes Update bereit, wird angeboten, es jetzt zu installieren."""
        pending = self.pending_update
        if pending and os.path.exists(pending["path"]):
            answer = messagebox.askyesnocancel(
                APP_NAME, f"Das Update auf Version {fmt_version(pending['version'])} ist heruntergeladen "
                          "und bereit.\n\nJetzt installieren? DocuBuilder wird dafür geschlossen und "
                          "startet danach neu.\n\nJa = installieren · Nein = ohne Update beenden",
                parent=self)
            if answer is None:
                return
            if answer:
                try:
                    self.install_update()
                except OSError as e:
                    messagebox.showerror(APP_NAME, install_error_text(e), parent=self)
                return
        self.destroy()

    # ---- Fenster
    def open_update(self, found=None):
        if not self.user or self._alive(self.default_win):
            return
        if self._alive(self.update_win):
            self.update_win.lift()
            return
        self.update_win = UpdateDialog(self, found)

    def open_settings(self):
        if not self.user or self._alive(self.default_win):
            return
        if not self._alive(self.settings_win):
            self.settings_win = Settings(self)
        self.settings_win.lift()

    def open_info(self):
        if not self.user:
            return
        if not self._alive(self.info_win):
            self.info_win = InfoDialog(self)
        self.info_win.lift()

    def open_default(self, first_run=False):
        if self._alive(self.settings_win):
            return
        if self._alive(self.default_win):
            self.default_win.lift()
            return
        if not first_run:
            if not self.user:
                return
            if not self.can_edit:
                messagebox.showinfo(APP_NAME, NO_RIGHTS, parent=self)
                return
        self.default_win = DefaultDialog(self, first_run=first_run)

    def startup(self):
        cleanup_old_program()                    # Reste eines Updates der einzelnen .exe
        if not self._watching and NETWORK["dir"]:
            self._watching = True
            self.after(30_000, self.watch_config)
        if STATE["offline"]:
            return self.show_login(OFFLINE_TEXT)
        if STATE["error"]:
            return self.locked()
        if has_admin(self.settings):
            self.show_login()
        else:                                    # Ersteinrichtung: Admin-Passwort und Benutzer
            self.show_login("Ersteinrichtung läuft …")
            self.open_default(first_run=True)

    def save(self):
        if not self.can_edit:
            messagebox.showerror(APP_NAME, NO_RIGHTS, parent=self)
            return False
        if not save_config(self.settings, self.items):
            if STATE["conflict"]:
                self.reload_config()
            return False
        if STATE["legacy"]:                      # ältere Datei übernommen: entfernen
            save_prefs(self.settings)
            unprotect_file(STATE["legacy"])
            try:
                os.remove(STATE["legacy"])
            except OSError:
                pass
            STATE["legacy"] = None
        return True

    def reload_config(self):
        """Aktuellen Stand der gemeinsamen Konfiguration übernehmen (z. B. nach Änderungen an einem
        anderen Platz). Wurde der eigene Benutzer gelöscht oder sein Passwort geändert: abmelden."""
        prefs = {k: self.settings[k] for k in ("theme", "scheme", "update_check")}
        settings, items = load_config()
        if STATE["error"] or STATE["offline"]:
            return                               # beim nächsten Anmelden wird es gemeldet
        settings.update(prefs)
        self.settings, self.items = settings, items
        if not self.user:
            return
        user = find_user(settings["users"], self.user["name"])
        if not user or user["hash"] != self.user["hash"]:
            self.login_name.set("")
            self.show_login("Dein Konto wurde geändert. Bitte neu anmelden.")
            return
        self.user = user
        self.build(force=True)
        if self._alive(self.settings_win):
            self.settings_win.refresh()

    def watch_config(self):
        """Alle 30 Sekunden prüfen, ob die gemeinsame Konfiguration geändert wurde, und neue
        Buttons übernehmen, solange kein Fenster zum Bearbeiten offen ist."""
        self.after(30_000, self.watch_config)
        if not self.user or self._alive(self.default_win) or self._alive(self.settings_win):
            return
        try:
            stamp = file_stamp(read_config_raw())
        except OSError:
            return
        if stamp != STATE["stamp"]:
            self.reload_config()

    def apply_theme(self):
        mode = self.settings["theme"]
        effective = detect_system_theme() if mode == "system" else mode
        self.is_dark = effective == "dark"
        scheme = COLOR_SCHEMES.get(self.settings["scheme"], COLOR_SCHEMES[DEFAULT_SCHEME])
        self.pal = scheme[1][effective]
        setup_style(self, self.pal)
        set_dark_titlebar(self, self.is_dark)
        for w in self.winfo_children():
            if isinstance(w, tk.Toplevel):
                w.configure(bg=self.pal["bg"])
                set_dark_titlebar(w, self.is_dark)

    def build(self, force=False):
        width = max(self.area.winfo_width(), 1)
        cols = max(1, width // (self.TILE_W + 10))
        if cols == self.cols and not force:
            return
        self.cols = cols
        if not self.user:
            return
        p = self.pal
        for w in self.area.winfo_children():
            w.destroy()
        items = [it for it in self.items if visible(it, self.user)]
        if not items:
            text = ("Noch keine Einträge. Lege sie unter „Einstellungen“ an." if self.can_edit else
                    "Für dich sind noch keine Einträge freigegeben. Bitte beim Admin melden.")
            ttk.Label(self.area, text=text, style="Mute.TLabel").pack(pady=30)
            return
        for c in range(cols):
            self.area.columnconfigure(c, weight=1, uniform="t")
        for i, it in enumerate(items):
            b = tk.Button(self.area, text=it["name"] or "Ohne Namen", wraplength=self.TILE_W - 24,
                          height=3, relief="flat", bd=0, cursor="hand2",
                          font=(self.ui_font, 11), bg=p["panel"], fg=p["fg"],
                          activebackground=p["accent"], activeforeground=p["accent_fg"],
                          highlightthickness=1, highlightbackground=p["line"],
                          highlightcolor=p["accent"],
                          command=lambda path=it["path"]: open_target(path, self.settings["files_dir"]))
            b.bind("<Enter>", lambda e, w=b: w.configure(bg=p["btn_hover"]))
            b.bind("<Leave>", lambda e, w=b: w.configure(bg=p["panel"]))
            b.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="nsew", ipady=6)


if __name__ == "__main__":
    if "--setup-machine" in sys.argv[1:]:
        # vom Installationsprogramm (mit Adminrechten) aufgerufen: gemeinsamen Ordner einrichten
        prepare_machine_dir(MACHINE_DIR)
    else:
        App().mainloop()
