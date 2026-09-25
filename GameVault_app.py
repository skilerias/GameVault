#!/usr/bin/env python3
"""
Steam Game Value Tracker — desktop app
------------------------------------------
Runs as a native desktop window (via pywebview) instead of a browser tab.
Same UI, same features. It still talks to Steam's real store API over the
internet — search, prices, DLC, everything is live Steam data, not local
mock data — and adds a "Play" button on each game card that launches the
game straight from your installed Steam client. Once it's running, the
same button turns into "Stop", which finds and closes the game's process
(requires psutil — see Requirements below).

Requirements:
    pip install flask requests pywebview psutil icoextract pillow pywin32
    (psutil is optional but required for the "Stop" button — see below.
    icoextract + pillow are optional too — required only to pull each local
    game's own icon off its .exe for its library card image. pywin32 is
    optional too — Windows only — required only so that changing the app
    icon (Themes page) also updates any taskbar/Start pin for this app
    immediately, instead of you having to unpin and repin it. The embedded
    EXE icon is changed only when you run the separate rebuild helper.)

Usage:
    python app.py
This opens a desktop window automatically. If pywebview isn't installed,
it falls back to a normal browser tab at http://127.0.0.1:5000.

Packaging into a real standalone .exe / .app (no Python required to run
it) is covered in README.md, which is included alongside this file.
"""

import base64
import difflib
import io
import json
import math
import os
import sys
import tempfile
import threading
import time
import requests
import re
import webbrowser
import shutil
from html import unescape
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, Response, jsonify, request

try:
    import webview  # pywebview — gives us the native desktop window
    HAS_WEBVIEW = True
except ImportError:
    HAS_WEBVIEW = False

try:
    import psutil  # used to find/stop the running game process for the Stop button
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

try:
    from icoextract import IconExtractor  # pulls the embedded icon out of a .exe
    from PIL import Image                 # converts it to a PNG for the card image
    HAS_ICON_EXTRACT = True
except ImportError:
    HAS_ICON_EXTRACT = False

try:
    import win32com.client  # pywin32 — lets us edit the .lnk icon of any
                             # taskbar/Start shortcut pinned to this app, see
                             # _sync_pinned_taskbar_icon. Windows-only feature.
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

try:
    import PIL  # noqa: F401 -- just probing availability here; each feature
    HAS_PIL = True  # that needs it does its own local `from PIL import Image`.
except ImportError:
    HAS_PIL = False

app = Flask(__name__)

# Runtime/user data is kept in a stable per-user folder instead of next to
# the executable. That means you can move/copy GameVault.exe anywhere
# (Desktop, another folder, USB, etc.) and the same library/categories/cache
# follow you automatically. On the first run after this update, existing
# data next to the old .exe is migrated into the new folder.
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
    _APPDATA_ROOT = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA") or APP_DIR
    DATA_DIR = os.path.join(_APPDATA_ROOT, "GameVault")
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = APP_DIR

os.makedirs(DATA_DIR, exist_ok=True)

def _data_file(name):
    target = os.path.join(DATA_DIR, name)
    # One-time migration from the old "next to the exe" layout. Never
    # overwrite data that already exists in the new persistent location.
    if DATA_DIR != APP_DIR and not os.path.exists(target):
        old = os.path.join(APP_DIR, name)
        if os.path.isfile(old):
            try:
                shutil.copy2(old, target)
            except OSError:
                pass
    return target

LIBRARY_FILE = _data_file("library.json")
DLC_CACHE_FILE = _data_file("dlc_cache.json")
DLC_CACHE_TTL = 6 * 60 * 60
CATEGORIES_FILE = _data_file("categories.json")
LOCAL_FOLDERS_FILE = _data_file("local_folders.json")


# ---------- app icon ----------
# Embedded as base64 so the app stays a single file. This is used for: the
# browser-tab / in-window favicon (works everywhere, incl. the plain
# browser-tab fallback mode), and — best effort — the actual desktop window
# icon via webview.start(icon=...), which pywebview only supports on the
# GTK/Qt (Linux) backends. On Windows and macOS, pywebview says the window
# icon has to be baked into the packaged .exe/.app at build time instead
# (see README.md's packaging section) — icon.ico is written out below
# specifically so it's ready to hand to that packaging step
# (e.g. `pyinstaller --icon=icon.ico ...`) without any extra conversion.
_ICON_PNG_B64 = "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"

_ICON_ICO_B64 = None  # built lazily from the PNG bytes below, no runtime PIL dependency needed

# The user can swap in their own picture as the app icon (Themes page ->
# "App icon"). When present, custom_icon.png overrides the built-in icon
# above everywhere: the in-page favicon, the window icon written for
# packaging, and (best effort, GTK/Qt only) the live desktop window icon.
# "Reset to original" just removes these two files, which falls back to
# the embedded default automatically.
CUSTOM_ICON_FILE = _data_file("custom_icon.png")
CUSTOM_ICON_ICO_FILE = _data_file("custom_icon.ico")


def _default_icon_png_bytes():
    return base64.b64decode(_ICON_PNG_B64)


def _current_icon_png_bytes():
    """The icon that should be shown right now: the user's custom picture if
    they set one, otherwise the built-in default."""
    if os.path.exists(CUSTOM_ICON_FILE):
        try:
            with open(CUSTOM_ICON_FILE, "rb") as f:
                return f.read()
        except OSError:
            pass
    return _default_icon_png_bytes()


def _current_icon_data_uri():
    return "data:image/png;base64," + base64.b64encode(_current_icon_png_bytes()).decode("ascii")


def _write_app_icon_files():
    """Write the current icon to per-user AppData.

    The ICO filename changes when the icon changes. This is intentional:
    Windows Explorer caches shortcut icons aggressively, so a new filename
    plus a recreated shortcut is much more reliable than overwriting one ICO.
    """
    png_path = os.path.join(DATA_DIR, "icon.png")
    try:
        png_bytes = _current_icon_png_bytes()
        with open(png_path, "wb") as f:
            f.write(png_bytes)
        try:
            from PIL import Image
            img = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
            # Stable filename for the current image, unique per image.
            import hashlib
            digest = hashlib.sha256(png_bytes).hexdigest()[:16]
            ico_path = os.path.join(DATA_DIR, "icon_" + digest + ".ico")
            img.save(ico_path, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
            # Keep a conventional current icon.ico too for compatibility.
            try:
                shutil.copy2(ico_path, os.path.join(DATA_DIR, "icon.ico"))
            except OSError:
                pass
            return ico_path
        except ImportError:
            pass
        except Exception:
            pass
    except OSError:
        pass
    fallback = os.path.join(DATA_DIR, "icon.ico")
    return fallback if os.path.exists(fallback) else None


def _windows_shortcut_paths():
    if not sys.platform.startswith("win"):
        return []
    home = os.path.expanduser("~")
    desktop_candidates = [
        os.path.join(home, "Desktop"),
        os.path.join(os.environ.get("OneDrive", ""), "Desktop"),
    ]
    desktop = next((os.path.join(d, "GameVault.lnk") for d in desktop_candidates if d and os.path.isdir(d)), None)
    appdata = os.environ.get("APPDATA", "")
    start = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "GameVault.lnk") if appdata else None
    return [p for p in (desktop, start) if p]


def _ensure_gamevault_shortcuts(ico_path):
    """Create/update normal Windows Desktop + Start Menu shortcuts.

    These shortcuts point to the installed GameVault.exe and use an external
    ICO. This lets the user change the icon without touching the EXE itself.
    """
    if not (sys.platform.startswith("win") and HAS_PYWIN32 and ico_path):
        return
    try:
        target = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
        working = os.path.dirname(target)
        shell = win32com.client.Dispatch("WScript.Shell")
        for shortcut_path in _windows_shortcut_paths():
            try:
                os.makedirs(os.path.dirname(shortcut_path), exist_ok=True)
                # Recreate our shortcut so Explorer notices the new icon path.
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                sc = shell.CreateShortCut(shortcut_path)
                sc.TargetPath = target
                sc.WorkingDirectory = working
                sc.IconLocation = f"{ico_path},0"
                sc.Description = "GameVault - Game Library"
                sc.Save()
            except Exception:
                pass
        _sync_pinned_taskbar_icon(ico_path)
        try:
            import ctypes
            SHCNE_UPDATEDIR = 0x00001000
            SHCNE_ASSOCCHANGED = 0x08000000
            SHCNF_PATHW = 0x0005
            ctypes.windll.shell32.SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATHW, None, None)
            ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, 0, None, None)
        except Exception:
            pass
        try:
            subprocess = __import__("subprocess")
            subprocess.run(["ie4uinit.exe", "-show"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
        except Exception:
            pass
    except Exception:
        pass

def _sync_pinned_taskbar_icon(ico_path):
    """Windows only, best-effort: find any taskbar/Start shortcut (.lnk)
    pinned to this app and repoint its icon at ico_path, then tell Explorer
    to redraw icons. This is what actually fixes the "I changed the icon
    but the pin still shows the old one" problem: Windows records a pinned
    shortcut's icon at pin time and never re-reads it on its own, so a
    later icon change (this app's Themes page, or a fresh install with a
    different default icon) never reaches an existing pin by itself — you'd
    otherwise have to unpin and repin it every time. Requires pywin32; a
    missing pywin32, a shortcut Explorer won't let us touch, or anything
    else going wrong here is silently ignored, since this is a nice-to-have
    on top of the icon change itself, not something that should ever crash
    the app or block it from starting."""
    if not (sys.platform.startswith("win") and HAS_PYWIN32 and ico_path):
        return
    try:
        frozen = getattr(sys, "frozen", False)
        target = os.path.abspath(sys.executable if frozen else __file__)
        appdata = os.environ.get("APPDATA", "")
        pin_dirs = [
            os.path.join(appdata, "Microsoft", "Internet Explorer",
                         "Quick Launch", "User Pinned", "TaskBar"),
            os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs"),
        ]
        shell = win32com.client.Dispatch("WScript.Shell")
        touched = False
        for pin_dir in pin_dirs:
            if not os.path.isdir(pin_dir):
                continue
            for root, _dirs, files in os.walk(pin_dir):
                for name in files:
                    if not name.lower().endswith(".lnk"):
                        continue
                    lnk_path = os.path.join(root, name)
                    try:
                        shortcut = shell.CreateShortcut(lnk_path)
                        shortcut_target = shortcut.TargetPath or ""
                        is_ours = os.path.abspath(shortcut_target) == target
                        if not is_ours and not frozen:
                            # Script mode: the shortcut's real target is
                            # python.exe, so only touch it if our script is
                            # among its arguments -- otherwise we'd be
                            # rewriting the icon of someone else's unrelated
                            # Python shortcut.
                            is_ours = (
                                os.path.basename(shortcut_target).lower().startswith("python")
                                and os.path.abspath(__file__) in (shortcut.Arguments or "")
                            )
                        if is_ours:
                            shortcut.IconLocation = "{},0".format(ico_path)
                            shortcut.save()
                            touched = True
                    except Exception:
                        continue
        if touched:
            import ctypes
            SHCNE_ASSOCCHANGED = 0x08000000
            SHCNF_IDLIST = 0x0000
            ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
    except Exception:
        pass


def _schedule_external_exe_rebuild(ico_path):
    """Deprecated: EXE rebuilding is manual via rebuild_gamevault.bat.

    This deliberately does nothing. A running PyInstaller one-file EXE must
    never rebuild or replace itself.
    """
    return False

def _extract_theme_color_from_bytes(img_bytes):
    """Pick a single accent hex color out of an arbitrary picture, for the
    "Theme from a picture" option. Downsamples, buckets into a small
    palette, then prefers whichever swatch is both fairly saturated and not
    too close to pure black/white -- picking the literal average color
    tends to produce mud, and the single most common color is often just
    the background. Falls back to the plain average if nothing in the
    picture is saturated enough (e.g. a black-and-white photo)."""
    from PIL import Image
    import io
    import colorsys

    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    img.thumbnail((150, 150))

    paletted = img.quantize(colors=12, method=Image.MEDIANCUT)
    palette = paletted.getpalette()
    candidates = []
    for count, idx in paletted.getcolors():
        r, g, b = palette[idx * 3:idx * 3 + 3]
        h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        candidates.append((count, s, l, r, g, b))

    vivid = [c for c in candidates if c[1] >= 0.18 and 0.15 <= c[2] <= 0.82]
    pick_from = vivid if vivid else candidates
    pick_from.sort(key=lambda c: (c[1], c[0]), reverse=True)  # most saturated, then most common
    _, _, _, r, g, b = pick_from[0]
    return "#{:02x}{:02x}{:02x}".format(r, g, b)


def _square_crop_and_resize(img, size=256):
    """Center-crop to square then resize -- so a wide/tall photo picked for
    the app icon doesn't come out squished."""
    w, h = img.size
    side = min(w, h)
    left = (w - side) // 2
    top = (h - side) // 2
    img = img.crop((left, top, left + side, top + side))
    return img.resize((size, size))


def _read_image_bytes_from_request(data):
    """Shared by the icon and theme-from-picture endpoints: pull the source
    image bytes either from a local file path (desktop native picker) or a
    data: URI (browser-tab fallback's <input type=file>). Returns
    (bytes, error_code_or_None)."""
    path = (data or {}).get("path")
    image_data = (data or {}).get("image_data")
    if path:
        if not os.path.isfile(path):
            return None, "not_found"
        try:
            with open(path, "rb") as f:
                return f.read(), None
        except OSError:
            return None, "not_found"
    if image_data:
        try:
            b64 = image_data.split(",", 1)[-1]
            return base64.b64decode(b64), None
        except Exception:
            return None, "bad_image"
    return None, "missing_image"


CC = "iq"    # country/currency code for pricing — "iq" (Iraq) gets Steam's discounted
             # USD_MENA regional pricing instead of full US pricing. Other MENA countries
             # that also qualify: bh, eg, jo, lb, om, ps, tr, ye, dz, ly, ma, tn, sa, ae, qa, kw
LANG = "en"

SEARCH_URL = "https://store.steampowered.com/api/storesearch/"
DETAILS_URL = "https://store.steampowered.com/api/appdetails"

# Steam's store API is picky about requests without a normal browser
# User-Agent, and throttles/blocks callers more aggressively without one.
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
}


# A shared Session (with a bigger connection pool) instead of plain
# requests.get() everywhere: requests.get() opens a brand-new TCP+TLS
# connection to Steam on every single call, which is most of why typing in
# the search box or opening the DLC picker feels slow. A shared Session
# reuses (keeps alive) connections to store.steampowered.com across calls,
# which cuts a lot of that latency, especially for the many back-to-back
# calls the DLC picker makes.
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
_adapter = requests.adapters.HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=1)
SESSION.mount("https://", _adapter)
SESSION.mount("http://", _adapter)


# ---------- storage ----------
#
# All writes below go through _atomic_write_json, which:
#   1. writes the new data to a temp file in the same folder,
#   2. flushes + fsyncs it to disk,
#   3. atomically replaces the real file with os.replace().
# os.replace() on Windows/NTFS either fully succeeds or fully fails - there
# is no in-between state where the target file is half-written. That means
# a power loss can no longer leave library.json truncated/corrupted: the
# old file stays intact until the new one is completely and safely on disk.
# Before replacing, the previous good file is also copied to a rotating
# ".bak" as a second safety net.

def _atomic_write_json(path, data):
    folder = os.path.dirname(path) or "."
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        if os.path.exists(path):
            try:
                shutil.copy2(path, path + ".bak")
            except OSError:
                pass
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def _load_json_with_backup(path, default):
    for candidate in (path, path + ".bak"):
        if os.path.exists(candidate):
            try:
                with open(candidate, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if candidate != path:
                    # Main file was missing/corrupted; restore it from
                    # the backup right away so future saves stay honest.
                    try:
                        shutil.copy2(candidate, path)
                    except OSError:
                        pass
                return data
            except (OSError, ValueError):
                continue
    return default


def load_library():
    return _load_json_with_backup(LIBRARY_FILE, [])


def save_library(games):
    _atomic_write_json(LIBRARY_FILE, games)

def load_custom_categories():
    data = _load_json_with_backup(CATEGORIES_FILE, [])
    return data if isinstance(data, list) else []

def save_custom_categories(categories):
    _atomic_write_json(CATEGORIES_FILE, categories)


def load_local_folders():
    data = _load_json_with_backup(LOCAL_FOLDERS_FILE, [])
    return data if isinstance(data, list) else []

def save_local_folders(folders):
    _atomic_write_json(LOCAL_FOLDERS_FILE, folders)


def _extract_exe_icon_data_uri(exe_path):
    """Best-effort: pull the icon embedded in a .exe out and return it as a
    data: URI PNG, so a local game's card shows the same icon as the .exe
    itself. Needs icoextract + Pillow (optional — see Requirements at the
    top of this file); fails soft to None (blank card image) if either is
    missing, or if this particular .exe has no icon resource to read."""
    if not HAS_ICON_EXTRACT:
        return None
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=".ico")
        os.close(fd)
        IconExtractor(exe_path).export_icon(tmp_path, num=0)
        with Image.open(tmp_path) as img:
            img = img.convert("RGBA")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


# A recursive folder scan for "every .exe" turns up a lot of things that
# aren't the game itself — installers, redistributables, anti-cheat and
# crash-reporter helpers that ship alongside almost every game. Skipping
# filenames that obviously match those keeps a folder scan from flooding
# the library with entries nobody would call a game.
_EXE_EXCLUDE_HINTS = (
    "unins", "setup", "redist", "vcredist", "vc_redist", "directx", "dxsetup",
    "dxwebsetup", "crashpad", "crashreporter", "crash_reporter", "crashhandler",
    "crashreport", "crash_report", "unitycrashhandler", "prereq", "prereqsetup",
    "dotnetfx", "dotnet", "netfx", "eossdk", "eosbootstrapper", "battleye",
    "easyanticheat", "vulkan-1", "ueprereqsetup", "install", "updater",
    "launcher_installer", "oalinst", "openal", "physx", "fmod", "wwise",
    "miles", "nvidia", "amd_ags", "bugsplat", "sentry-native", "crashlogger",
    "overlay", "vc_runtime", "msvcp", "msvcr", "vcruntime", "steamworks_common",
    "steam_api", "webhelper", "subprocess", "cef_process", "elevate",
    "elevation", "helper.exe", "service.exe", "diagnostics", "benchmark",
    "unrealcef", "epicwebhelper", "activation", "licens",
    "chromedriver", "nwjc", "geckodriver", "phantomjs",
)


def _looks_like_game_exe(filename):
    lower = filename.lower()
    return not any(hint in lower for hint in _EXE_EXCLUDE_HINTS)


# ---------- opportunistic Steam matching for local (non-Steam) games ----------
# A local .exe's filename is often close to, but not exactly, its real Steam
# title (build suffixes, underscores, engine noise). These two helpers turn
# the filename into a plausible search term and, if a confident match turns
# up on Steam, hand back that listing's appid so the caller can pull the
# game's real name/artwork/genres/tags instead of guessing from the file.
# A local game that doesn't match anything on Steam (an Epic/GOG/itch.io
# exclusive, for instance) is left exactly as before — this only enriches,
# it never blocks a game from being added.
_EXE_NAME_NOISE = (
    "win64", "win32", "x64", "x86", "shipping", "final", "release",
    "launcher", "client", "game", "app", "bin", "retail", "steam",
    "dx11", "dx12", "vulkan", "demo",
)


def _clean_exe_name_for_match(filename):
    name = os.path.splitext(filename)[0]
    name = re.sub(r"[_\-\.]+", " ", name)
    name = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", name)  # camelCase -> spaced
    tokens = [t for t in name.split() if t.lower() not in _EXE_NAME_NOISE]
    cleaned = " ".join(tokens).strip()
    return cleaned or name.strip()


def _normalize_for_compare(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _match_local_exe_to_steam(filename):
    """Match a local .exe's cleaned filename to a Steam listing, so its
    real name, artwork, price, genres and community tags get pulled in
    automatically — exactly like a game added straight from Steam. Always
    links to the closest-named result among Steam's search matches when
    any exist. This can occasionally pick the wrong listing for a short or
    generic filename (there's no way to be certain from a filename alone)
    — if that happens, open the game's editor page to relink it to the
    right one, or unlink it."""
    term = _clean_exe_name_for_match(filename)
    if not term:
        return None
    try:
        results = steam_search(term)
    except Exception:
        return None
    if not results:
        return None
    target = _normalize_for_compare(term)
    best_item, best_ratio = results[0], -1.0
    for item in results[:10]:
        cand = _normalize_for_compare(item.get("name", ""))
        if not cand:
            continue
        ratio = difflib.SequenceMatcher(None, target, cand).ratio()
        if target == cand:
            ratio = 1.0
        if ratio > best_ratio:
            best_ratio, best_item = ratio, item
    return best_item.get("id") if best_item else None


def _build_local_game_entry(path, watch_folder, existing_ids):
    """Build one library entry for a local .exe. If its filename confidently
    matches a Steam listing, the entry is enriched with Steam's real name,
    header image, genres and community tags (fixing blurry icons and empty
    tags); otherwise it falls back to the bare filename + extracted .exe
    icon, same as before. `existing_ids` is a set of appids already in the
    library — it is updated in place so a caller building several entries in
    a row never collides."""
    filename = os.path.basename(path)
    fallback_name = os.path.splitext(filename)[0]
    norm_path = os.path.normcase(os.path.normpath(path))
    appid = -(abs(hash(norm_path)) % 2_000_000_000 + 1)
    while appid in existing_ids:
        appid -= 1
    existing_ids.add(appid)

    entry = {
        "appid": appid,
        "name": fallback_name,
        "image": _extract_exe_icon_data_uri(path) or "",
        "source": "local",
        "exe_path": path,
        "watch_folder": watch_folder,
        "steam_appid": None,
        "price": None,
        "currency": None,
        "discount": 0,
        "steam_genres": [],
        "steam_tags": [],
        "developers": [],
        "publishers": [],
        "dlcs": [],
        "bundle_plan": {"cost": 0, "bundles": [], "individual": []},
        "rating": None,
        "played": False,
        "notes": "",
        "pros": "",
        "cons": "",
        "category": "",
        "favorite": False,
        "custom_categories": [],
        "playtime_seconds": 0,
    }

    matched_appid = _match_local_exe_to_steam(filename)
    if matched_appid:
        details = steam_details(matched_appid, include_tags=True)
        if details:
            entry["name"] = details["name"]
            entry["image"] = details["image"] or entry["image"]
            entry["price"] = details["price"]
            entry["currency"] = details["currency"]
            entry["discount"] = details["discount"]
            entry["steam_genres"] = details["steam_genres"]
            entry["steam_tags"] = details["steam_tags"]
            entry["category"] = details["category"]
            entry["developers"] = details.get("developers", [])
            entry["publishers"] = details.get("publishers", [])
            entry["steam_appid"] = matched_appid
    return entry


def _scan_folder_for_exes(folder):
    """Every .exe under `folder`, recursively, minus the obvious non-game
    noise from _looks_like_game_exe. Returns normalized absolute paths."""
    found = []
    for root, _dirs, files in os.walk(folder):
        for fname in files:
            if fname.lower().endswith(".exe") and _looks_like_game_exe(fname):
                found.append(os.path.normpath(os.path.join(root, fname)))
    return found


def _apply_folder_scan(folder, games):
    """Reconcile one watched folder's current .exe contents against the
    library in place: adds newly-found exes as local games, removes local
    games under this folder whose exe is no longer there. Returns
    (added_count, removed_count); mutates `games`."""
    found = _scan_folder_for_exes(folder)
    found_by_norm = {os.path.normcase(p): p for p in found}
    found_paths = set(found_by_norm.keys())
    folder_norm = os.path.normcase(os.path.normpath(folder))

    existing_local_paths = {
        os.path.normcase(os.path.normpath(g["exe_path"]))
        for g in games if g.get("source") == "local" and g.get("exe_path")
    }

    removed = 0
    kept = []
    for g in games:
        if (g.get("source") == "local" and g.get("watch_folder")
                and os.path.normcase(os.path.normpath(g["watch_folder"])) == folder_norm
                and os.path.normcase(os.path.normpath(g.get("exe_path", ""))) not in found_paths):
            removed += 1
            continue
        kept.append(g)
    games[:] = kept

    added = 0
    existing_ids = {g["appid"] for g in games}
    for norm_path, real_path in found_by_norm.items():
        if norm_path in existing_local_paths:
            continue
        games.append(_build_local_game_entry(real_path, folder, existing_ids))
        added += 1
    return added, removed


# ---------- lightweight TF-IDF (no numpy/scipy/scikit-learn) ----------
# api_recommendations() below scores candidate games by text similarity to
# your favorites, using the same TF-IDF + cosine-similarity technique as
# content_based.ipynb in IsabelJoseph101/steam_recommendation_system
# (github.com/IsabelJoseph101/steam_recommendation_system). That notebook
# leans on scikit-learn, but scikit-learn drags in numpy + scipy — roughly
# 190+ MB just for this one feature, way out of proportion for a small
# desktop app — so this reimplements the same math (word 1-2 grams,
# term-frequency × inverse-document-frequency, cosine similarity) in plain
# Python with only the standard library. Every "matrix" below is just a
# dict of {term: weight} per document; there's no real matrix in memory.

_STOPWORDS = {
    "a", "an", "the", "and", "or", "but", "of", "to", "in", "on", "for",
    "with", "is", "are", "was", "were", "be", "been", "being", "it", "its",
    "this", "that", "these", "those", "as", "at", "by", "from", "into",
    "your", "you", "their", "his", "her", "our", "will", "can", "has",
    "have", "had", "not", "no", "if", "then", "than", "so", "up", "out",
    "about", "all", "more", "most", "such", "only", "also",
}


def _tokenize(text):
    """Lowercase word tokens with stopwords dropped, plus adjacent-word
    bigrams — mirrors TfidfVectorizer(analyzer='word', ngram_range=(1, 2))."""
    words = [w for w in re.findall(r"[a-z0-9]+", text.lower()) if w not in _STOPWORDS]
    bigrams = [f"{a} {b}" for a, b in zip(words, words[1:])]
    return words + bigrams


def _tfidf_vectors(docs):
    """Given a list of raw text documents, return one {term: tf-idf weight}
    dict per document (unit-normalized so cosine similarity is a plain dot
    product)."""
    tokenized = [_tokenize(d) for d in docs]
    n_docs = len(tokenized)

    doc_freq = Counter()
    for tokens in tokenized:
        doc_freq.update(set(tokens))

    idf = {term: math.log((1 + n_docs) / (1 + df)) + 1 for term, df in doc_freq.items()}

    vectors = []
    for tokens in tokenized:
        tf = Counter(tokens)
        weights = {term: count * idf[term] for term, count in tf.items()}
        norm = math.sqrt(sum(w * w for w in weights.values())) or 1.0
        vectors.append({term: w / norm for term, w in weights.items()})
    return vectors


def _cosine(vec_a, vec_b):
    # Vectors are already unit-normalized, so cosine similarity is just the
    # dot product over whichever side has fewer terms.
    if len(vec_a) > len(vec_b):
        vec_a, vec_b = vec_b, vec_a
    return sum(w * vec_b.get(term, 0.0) for term, w in vec_a.items())


# ---------- Steam calls ----------

_search_cache = {}  # term -> (results, cached_at) — short-lived, in-memory only
_SEARCH_CACHE_TTL = 60  # seconds; just long enough to smooth out retyping/backspacing


def steam_search(term):
    cached = _search_cache.get(term)
    if cached and time.time() - cached[1] < _SEARCH_CACHE_TTL:
        return cached[0]
    try:
        r = SESSION.get(SEARCH_URL, params={"term": term, "cc": CC, "l": LANG}, timeout=10)
        r.raise_for_status()
        items = r.json().get("items", [])
    except requests.RequestException:
        # Steam hiccuped or rate-limited us — just show no results instead of crashing
        return []
    _search_cache[term] = (items, time.time())
    # Keep this cache small; it's only meant to smooth out a single typing burst.
    if len(_search_cache) > 200:
        oldest = min(_search_cache, key=lambda k: _search_cache[k][1])
        _search_cache.pop(oldest, None)
    return items


def steam_details(appid, include_tags=False):
    try:
        params = {
            "appids": appid, "cc": CC, "l": LANG,
            # a changing dummy param busts Steam's CDN cache so "refresh" always
            # pulls a live price instead of a stale cached one
            "_": int(time.time() * 1000),
        }
        r = SESSION.get(DETAILS_URL, params=params, timeout=10)
        r.raise_for_status()
    except requests.RequestException:
        # network hiccup or Steam rate-limited us (common if you refresh a big
        # library fast) — caller keeps the old price instead of crashing
        return None

    data = r.json().get(str(appid))
    if not data or not data.get("success"):
        return None
    d = data["data"]
    price_info = d.get("price_overview")
    if price_info:
        price, currency, discount = price_info["final"] / 100.0, price_info["currency"], price_info.get("discount_percent", 0)
    elif d.get("is_free"):
        price, currency, discount = 0.0, "USD", 0
    else:
        price, currency, discount = None, None, 0
    # Steam provides both "categories" (features such as Single-player,
    # Co-op, Steam Workshop, etc.) and "genres" (Action, RPG, Strategy, ...).
    # Keep both so the UI can automatically group games without manual tagging.
    steam_categories = [x.get("description", "").strip() for x in d.get("categories", []) if x.get("description")]
    steam_genres = [x.get("description", "").strip() for x in d.get("genres", []) if x.get("description")]
    # Steam's own "genres" field above is deliberately narrow (usually just
    # 2-3 broad tags like "Indie, Simulation") — it is NOT the same as the
    # longer, community-voted tag list shown on a game's store page (e.g.
    # "Base Building", "Automation", "Sandbox", ...). That richer list only
    # comes from SteamSpy (see steamspy_top_tags above). include_tags is
    # opt-in and only passed True by the callers that populate the library
    # (add/refresh) — NOT by the DLC-fetching path, which runs several of
    # these concurrently and would otherwise get serialized by SteamSpy's
    # ~1/sec throttle for no benefit (DLC entries don't need genre tags).
    steam_tags = steamspy_top_tags(appid, limit=10) if include_tags else []
    # Developer/publisher — used by the recommendation engine to catch
    # "more games from this studio" matches that pure genre overlap misses.
    developers = [x.strip() for x in (d.get("developers") or []) if x and x.strip()]
    publishers = [x.strip() for x in (d.get("publishers") or []) if x and x.strip()]
    return {
        "appid": appid,
        "name": d.get("name", "Unknown"),
        "image": d.get("header_image", ""),
        "price": price,
        "currency": currency,
        "discount": discount,
        "description": d.get("short_description", "") or "",
        "steam_categories": steam_categories,
        "steam_genres": steam_genres,
        "steam_tags": steam_tags,
        "developers": developers,
        "publishers": publishers,
        # category is now automatic; use the first Steam genre as the
        # primary grouping while preserving all Steam classifications.
        "category": steam_genres[0] if steam_genres else (steam_categories[0] if steam_categories else ""),
    }


# In-memory mirror of DLC_CACHE_FILE. Every steamspy_top_tags /
# steamspy_games_for_tag / DLC-bundle / recommendations lookup used to call
# _load_dlc_cache(), which re-read and re-parsed the *entire* cache file
# from disk every single time — including in tight loops like api_refresh
# iterating a whole library. Since this process is the only writer of that
# file, it's safe to keep one in-memory copy and only touch disk again when
# something actually changes it (_save_dlc_cache below still writes through
# to disk immediately on every mutation, exactly as before, so durability
# is unchanged — this only removes the redundant re-reads).
_DLC_CACHE_LOCK = threading.Lock()
_DLC_CACHE_MEM = None


def _load_dlc_cache():
    global _DLC_CACHE_MEM
    with _DLC_CACHE_LOCK:
        if _DLC_CACHE_MEM is None:
            try:
                if os.path.exists(DLC_CACHE_FILE):
                    with open(DLC_CACHE_FILE, "r", encoding="utf-8") as f:
                        _DLC_CACHE_MEM = json.load(f)
                else:
                    _DLC_CACHE_MEM = {}
            except (OSError, ValueError, TypeError):
                _DLC_CACHE_MEM = {}
        return _DLC_CACHE_MEM


def _save_dlc_cache(cache):
    global _DLC_CACHE_MEM
    with _DLC_CACHE_LOCK:
        _DLC_CACHE_MEM = cache
    try:
        _atomic_write_json(DLC_CACHE_FILE, cache)
    except OSError:
        pass


# ---------- SteamSpy: real crowd-sourced signal for recommendations ----------
#
# Steam itself has no public API for "people who played/reviewed this also
# played/reviewed that" — that collaborative-filtering model is trained on
# purchase/wishlist/playtime data across every Steam account, and it's
# Valve-internal, not something any outside app can query (see the note this
# used to show on the Recommendations page).
#
# SteamSpy (steamspy.com) is the closest thing that's actually public: an
# independent, free, no-key-needed API that aggregates real Steam player
# data — most usefully here, community-voted tags per game (thousands of
# players tagging "Roguelike", "Soulslike", "Metroidvania", etc., with vote
# counts) and, per tag, which games carry it. That's a genuine "what does
# the wider player base associate with this game" signal, which is used
# below in two ways: (1) pulling in a game's top community tags as extra
# text for its TF-IDF document, and (2) using SteamSpy's own by-tag game
# lists to widen the recommendation candidate pool beyond plain Steam-store
# keyword search. It's still not personalized collaborative filtering (no
# access to who-owns-what), but it's real aggregate player-community data,
# not just Steam store metadata.
#
# SteamSpy asks callers to keep individual (appdetails/tag) requests to
# roughly 1/second, so calls go through a small throttle, and results are
# cached to disk (reusing the same cache file as the DLC cache) so a given
# game or tag is only ever fetched live once every so often.

_STEAMSPY_URL = "https://steamspy.com/api.php"
_STEAMSPY_LOCK = threading.Lock()
_STEAMSPY_LAST_CALL = [0.0]
_STEAMSPY_MIN_INTERVAL = 1.1
_STEAMSPY_TAG_TTL = 7 * 24 * 60 * 60       # a game's top tags barely change — cache a week
_STEAMSPY_TAGGAMES_TTL = 24 * 60 * 60      # which games carry a tag shifts daily — cache a day


def _steamspy_throttle():
    with _STEAMSPY_LOCK:
        wait = _STEAMSPY_LAST_CALL[0] + _STEAMSPY_MIN_INTERVAL - time.time()
        if wait > 0:
            time.sleep(wait)
        _STEAMSPY_LAST_CALL[0] = time.time()


def steamspy_top_tags(appid, limit=6):
    """This game's top community-voted tags on SteamSpy, most-voted first.
    Fails soft to [] — SteamSpy is a small independent project, not an
    official Valve service, and is sometimes slow or briefly down."""
    cache = _load_dlc_cache()
    key = f"steamspy_tags:{appid}"
    entry = cache.get(key)
    if entry and time.time() - entry.get("_cached_at", 0) < _STEAMSPY_TAG_TTL:
        return entry.get("tags", [])
    _steamspy_throttle()
    top = []
    try:
        r = SESSION.get(_STEAMSPY_URL, params={"request": "appdetails", "appid": appid}, timeout=8)
        data = r.json()
        tags = data.get("tags") or {}
        if isinstance(tags, dict):
            # raw SteamSpy shape: {"Roguelike": 4213, "Indie": 3980, ...}
            top = [t for t, _votes in sorted(tags.items(), key=lambda kv: -kv[1])[:limit]]
        elif isinstance(tags, list):
            top = list(tags)[:limit]
    except (requests.RequestException, ValueError, AttributeError, TypeError):
        pass
    cache[key] = {"tags": top, "_cached_at": time.time()}
    _save_dlc_cache(cache)
    return top


def steamspy_games_for_tag(tag, limit=15):
    """Appids SteamSpy lists under this community tag (already ranked by
    owner count), for widening the recommendation candidate pool with games
    the wider player base tags the same way as your favorites — not just
    ones that happen to share a Steam store genre string."""
    cache = _load_dlc_cache()
    key = f"steamspy_tag_games:{tag.lower()}"
    entry = cache.get(key)
    if entry and time.time() - entry.get("_cached_at", 0) < _STEAMSPY_TAGGAMES_TTL:
        return entry.get("appids", [])
    _steamspy_throttle()
    appids = []
    try:
        r = SESSION.get(_STEAMSPY_URL, params={"request": "tag", "tag": tag}, timeout=10)
        data = r.json()
        appids = [int(a) for a in list(data.keys())[:limit]]
    except (requests.RequestException, ValueError, AttributeError, TypeError):
        pass
    cache[key] = {"appids": appids, "_cached_at": time.time()}
    _save_dlc_cache(cache)
    return appids


def _steam_price_batch(appids):
    """Fetch prices for many AppIDs in one Steam request (up to 100 at a time)."""
    prices = {}
    ids = [str(int(x)) for x in appids if x is not None]
    for start in range(0, len(ids), 100):
        batch = ids[start:start + 100]
        try:
            r = SESSION.get(
                DETAILS_URL,
                params={
                    "appids": ",".join(batch),
                    "cc": CC,
                    "l": LANG,
                    "filters": "price_overview",
                },
                timeout=15,
            )
            r.raise_for_status()
            payload = r.json()
            for appid in batch:
                item = payload.get(str(appid), {})
                data = item.get("data", {}) if item.get("success") else {}
                p = data.get("price_overview")
                if p:
                    prices[str(appid)] = {
                        "price": p.get("final", 0) / 100.0,
                        "currency": p.get("currency"),
                        "discount": p.get("discount_percent", 0),
                    }
                elif data.get("is_free"):
                    prices[str(appid)] = {"price": 0.0, "currency": "USD", "discount": 0}
                else:
                    prices[str(appid)] = {"price": None, "currency": None, "discount": 0}
        except (requests.RequestException, ValueError):
            continue
    return prices


def _steam_resolve_bundles(bundle_ids):
    """Resolve Steam bundle IDs to their included AppIDs and current prices."""
    ids = [str(int(x)) for x in bundle_ids if x is not None]
    if not ids:
        return []
    try:
        r = SESSION.get(
            "https://store.steampowered.com/actions/ajaxresolvebundles",
            params={"bundleids": ",".join(ids), "cc": CC, "l": LANG},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        return data if isinstance(data, list) else []
    except (requests.RequestException, ValueError, TypeError):
        return []


def _steam_package_details(package_ids):
    """Fetch Steam package/sub details (these are the 'Edition' offers shown on store pages)."""
    ids = [str(int(x)) for x in package_ids if x is not None]
    if not ids:
        return {}
    out = {}
    for start in range(0, len(ids), 100):
        batch = ids[start:start + 100]
        try:
            r = SESSION.get(
                "https://store.steampowered.com/api/packagedetails",
                params={"packageids": ",".join(batch), "cc": CC, "l": LANG}, timeout=15,
            )
            r.raise_for_status()
            payload = r.json()
            for pid in batch:
                row = payload.get(str(pid), {})
                if row.get("success") and isinstance(row.get("data"), dict):
                    out[int(pid)] = row["data"]
        except (requests.RequestException, ValueError, TypeError):
            continue
    return out


def _find_dlc_bundles(appid, dlc_ids, base_currency=None):
    """Find both Steam dynamic bundles and Edition/package offers that can acquire DLCs.

    Steam has two different mechanisms on store pages:
      * dynamic bundles (data-ds-bundleid)
      * packages/editions (data-ds-packageid), e.g. "Shadow of the Erdtree Edition".

    The latter is important because many games, including Elden Ring, expose their
    base+expansion deal as a package rather than a dynamic bundle.
    """
    target = {int(x) for x in dlc_ids}
    if not target:
        return []
    base_id = int(appid)
    try:
        r = SESSION.get(
            f"https://store.steampowered.com/app/{base_id}/",
            params={"cc": CC, "l": LANG}, timeout=15,
        )
        r.raise_for_status()
        html_text = r.text
    except requests.RequestException:
        return []

    # HTML attributes are HTML-escaped (e.g. &quot;), so unescape before JSON parsing.
    bundle_candidates = []
    bundle_pat = re.compile(
        r'data-ds-bundleid=["\'](\d+)["\'][^>]*data-ds-bundle-data=["\']([^"\']+)["\']', re.I
    )
    bundle_pat_rev = re.compile(
        r'data-ds-bundle-data=["\']([^"\']+)["\'][^>]*data-ds-bundleid=["\'](\d+)["\']', re.I
    )
    seen_b = set()
    for pat, rev in ((bundle_pat, False), (bundle_pat_rev, True)):
        for m in pat.finditer(html_text):
            bid, enc = (m.group(2), m.group(1)) if rev else (m.group(1), m.group(2))
            try:
                bd = json.loads(unescape(enc))
            except (ValueError, TypeError):
                continue
            if int(bid) not in seen_b:
                seen_b.add(int(bid))
                bundle_candidates.append((int(bid), bd))

    # Packages/editions are often the real "bundle" shown in the purchase boxes.
    package_ids = set(int(x) for x in re.findall(r'data-ds-packageid=["\'](\d+)["\']', html_text, re.I))
    package_data = _steam_package_details(sorted(package_ids))

    resolved = {
        int(x.get("bundleid")): x
        for x in _steam_resolve_bundles([x[0] for x in bundle_candidates])
        if isinstance(x, dict) and x.get("bundleid") is not None
    }

    out = []

    # Dynamic bundles: their component prices already include the bundle discount.
    for bid, bd in bundle_candidates:
        if bd.get("m_bMustPurchaseAsSet"):
            continue
        covered, included = set(), set()
        price = 0
        valid = True
        for item in bd.get("m_rgItems") or []:
            try:
                apps = {int(x) for x in (item.get("m_rgIncludedAppIDs") or [])}
            except (ValueError, TypeError):
                continue
            included |= apps
            hit = apps & target
            if hit:
                covered |= hit
                p = item.get("m_nFinalPriceWithBundleDiscount")
                if p is None:
                    valid = False
                    break
                price += int(p)
        if not valid or not covered or not included.issubset(target | {base_id}):
            continue
        rv = resolved.get(bid, {})
        if price <= 0 and rv.get("final_price") is not None:
            price = int(rv["final_price"])
        out.append({
            "id": bid,
            "kind": "bundle",
            "name": str(rv.get("name") or "Steam DLC Bundle"),
            "price": price / 100.0,
            "currency": base_currency or "USD",
            "includes_base": base_id in included,
            "covered_dlc": sorted(covered),
            "appids": sorted(included),
        })

    # Packages/editions: use the package's actual purchase price. For this app's
    # value model, if a package contains the base game plus DLC, that package price
    # is the acquisition cost of the covered DLC because the base game's standalone
    # value is counted separately in the library total.
    for pid, pd in package_data.items():
        apps = []
        for item in pd.get("apps") or []:
            try:
                aid = int(item.get("id"))
            except (TypeError, ValueError):
                continue
            apps.append(aid)
        included = set(apps)
        covered = included & target
        if not covered:
            continue
        # Do not use editions that bring unrelated games/DLCs into the value plan.
        if not included.issubset(target | {base_id}):
            continue
        price_info = pd.get("price") or {}
        final_price = price_info.get("final")
        if final_price is None:
            continue
        out.append({
            "id": pid,
            "kind": "package",
            "name": str(pd.get("name") or "Steam Edition"),
            "price": float(final_price) / 100.0,
            "currency": price_info.get("currency") or base_currency or "USD",
            "includes_base": base_id in included,
            "covered_dlc": sorted(covered),
            "appids": sorted(included),
        })

    # Keep only the cheapest acquisition option for an identical DLC coverage set.
    best = {}
    for b in out:
        key = tuple(b["covered_dlc"])
        if key not in best or b["price"] < best[key]["price"]:
            best[key] = b
    return list(best.values())

def _best_dlc_plan(selected,bundles):
    chosen={int(d["appid"]):d for d in selected}; ids=list(chosen)
    if not ids:return {"cost":0,"bundles":[],"individual":[]}
    usable=[]
    for b in bundles or []:
        cov=set(map(int,b.get("covered_dlc",[])))&set(chosen)
        if cov: usable.append((cov,float(b.get("price") or 0),b))
    if len(ids)>20:
        # Greedy for unusually huge DLC lists.
        rem=set(ids); picks=[]
        while rem:
            best=None; saving=0
            for cov,cost,b in usable:
                hit=cov&rem; sv=sum(float(chosen[x].get("price") or 0) for x in hit)-cost
                if hit and sv>saving: saving=sv; best=(hit,b)
            if not best: break
            picks.append(best[1]); rem-=best[0]
        ind=sorted(rem); cost=sum(float(b["price"]) for b in picks)+sum(float(chosen[x].get("price") or 0) for x in ind)
        return {"cost":cost,"bundles":picks,"individual":ind}
    n=len(ids); idx={x:i for i,x in enumerate(ids)}; full=(1<<n)-1
    dp={0:(0,[],[])}
    for mask in range(full+1):
        if mask not in dp: continue
        cost,bs,inds=dp[mask]
        for i,x in enumerate(ids):
            if mask&(1<<i): continue
            nm=mask|(1<<i); nc=cost+float(chosen[x].get("price") or 0)
            if nm not in dp or nc<dp[nm][0]: dp[nm]=(nc,bs,inds+[x])
        for cov,bc,b in usable:
            bm=0
            for x in cov:
                if x in idx: bm|=1<<idx[x]
            nm=mask|bm
            if nm!=mask:
                nc=cost+bc
                if nm not in dp or nc<dp[nm][0]: dp[nm]=(nc,bs+[b],inds)
    cost,bs,inds=dp[full]
    return {"cost":cost,"bundles":bs,"individual":inds}

def steam_dlc_options(appid, force_refresh=False):
    """Get DLCs with caching/batching and detect an exact complete-DLC bundle."""
    cache = _load_dlc_cache()
    now = time.time()

    try:
        r = SESSION.get(
            DETAILS_URL,
            params={"appids": appid, "cc": CC, "l": LANG}, timeout=15,
        )
        r.raise_for_status()
        data = r.json().get(str(appid))
        if not data or not data.get("success"):
            return {"dlcs": [], "bundle": None}
        game_data = data["data"]
        dlc_ids = [int(x) for x in (game_data.get("dlc", []) or [])]
        base_currency = None
        if game_data.get("price_overview"):
            base_currency = game_data["price_overview"].get("currency")
        elif game_data.get("is_free"):
            base_currency = "USD"
    except (requests.RequestException, ValueError, TypeError):
        return {"dlcs": [], "bundle": None}

    if not dlc_ids:
        return {"dlcs": [], "bundle": None}

    result_by_id = {}
    missing = []
    for dlc_id in dlc_ids:
        cached = cache.get(str(dlc_id))
        if (not force_refresh and cached and
                now - float(cached.get("_cached_at", 0)) < DLC_CACHE_TTL):
            result_by_id[dlc_id] = {k: v for k, v in cached.items() if k != "_cached_at"}
        else:
            missing.append(dlc_id)

    # Fetching each DLC's metadata (a batch price call, plus one appdetails
    # call per not-yet-cached DLC) and finding bundles/editions (a full
    # store-page scrape plus follow-up calls) are two independent, slow
    # network round trips. Running them side by side instead of one after
    # another is roughly twice as fast the first time you open a game with
    # DLC that isn't cached yet.
    with ThreadPoolExecutor(max_workers=max(2, min(13, len(missing) + 1))) as pool:
        bundle_future = pool.submit(_find_dlc_bundles, appid, dlc_ids, base_currency)

        if missing:
            batch_prices = _steam_price_batch(missing)

            def fetch_metadata(dlc_id):
                d = steam_details(dlc_id)
                if not d:
                    return None
                p = batch_prices.get(str(dlc_id))
                return {
                    "appid": dlc_id,
                    "name": d.get("name", "Unknown"),
                    "image": d.get("image", ""),
                    "price": p["price"] if p else d.get("price"),
                    "currency": p["currency"] if p else d.get("currency"),
                    "discount": p["discount"] if p else d.get("discount", 0),
                }

            futures = [pool.submit(fetch_metadata, did) for did in missing]
            for future in as_completed(futures):
                item = future.result()
                if item:
                    result_by_id[item["appid"]] = item
                    cache[str(item["appid"])] = {**item, "_cached_at": now}
            _save_dlc_cache(cache)

        bundles = bundle_future.result()

    dlcs = [result_by_id[x] for x in dlc_ids if x in result_by_id]
    return {"dlcs": dlcs, "bundles": bundles}


# ---------- routes ----------

@app.route("/")
def index():
    # Substituted per-request (not baked in once at import time) so a
    # custom icon set via the Themes page shows up on the very next page
    # load, with no restart needed.
    #
    # PAGE is plain HTML/CSS/JS with a single literal placeholder — it has
    # no Jinja {{ }} / {% %} syntax at all, so running it through
    # render_template_string() on every request just meant Flask compiling
    # and walking a 100KB+ Jinja template for nothing. A plain Response
    # with the same bytes is byte-for-byte identical output, just without
    # that wasted work on every page load.
    html = PAGE.replace("__APP_ICON_DATA_URI__", _current_icon_data_uri())
    return Response(html, mimetype="text/html")


@app.route("/api/games")
def api_games():
    games = load_library()
    # Add in-progress session time (accumulated but not yet flushed to disk)
    # so the number on screen keeps ticking up live while a game is running.
    with _PLAYTIME_LOCK:
        pending = dict(_PENDING_PLAYTIME)
    if pending:
        for g in games:
            extra = pending.get(int(g["appid"]))
            if extra:
                g["playtime_seconds"] = float(g.get("playtime_seconds") or 0) + extra
    return jsonify(games)


@app.route("/api/categories")
def api_categories():
    return jsonify(load_custom_categories())

@app.route("/api/category/create", methods=["POST"])
def api_category_create():
    name = str((request.json or {}).get("name", "")).strip()
    if not name:
        return jsonify({"error": "empty_name"}), 400
    categories = load_custom_categories()
    if any(c.lower() == name.lower() for c in categories):
        return jsonify({"error": "already_exists"}), 400
    categories.append(name)
    save_custom_categories(categories)
    return jsonify(categories)

@app.route("/api/category/delete", methods=["POST"])
def api_category_delete():
    name = str((request.json or {}).get("name", "")).strip()
    categories = [c for c in load_custom_categories() if c != name]
    save_custom_categories(categories)
    games = load_library()
    for g in games:
        g["custom_categories"] = [c for c in g.get("custom_categories", []) if c != name]
    save_library(games)
    return jsonify({"categories": categories, "games": games})

@app.route("/api/dlc-options/<int:appid>")
def api_dlc_options(appid):
    force = request.args.get("refresh") == "1"
    base = steam_details(appid)
    result = steam_dlc_options(appid, force_refresh=force)
    return jsonify({"base": base, "dlcs": result.get("dlcs", []), "bundles": result.get("bundles", [])})

@app.route("/api/search")
def api_search():
    term = request.args.get("term", "").strip()
    if not term:
        return jsonify([])
    try:
        limit = int(request.args.get("limit", 8))
    except (TypeError, ValueError):
        limit = 8
    limit = max(1, min(limit, 25))
    items = steam_search(term)
    return jsonify([
        {"appid": i["id"], "name": i["name"], "thumb": i.get("tiny_image", "")}
        for i in items[:limit]
    ])


@app.route("/api/recommendations")
def api_recommendations():
    """Suggest games not already in the library, based on the games you
    actually seem to like — not just which genres happen to be present.

    Important honesty note (also shown in the UI): this is NOT Steam's own
    recommendation engine. Steam's is a proprietary machine-learning system
    trained on purchases, wishlists and playtime across all Steam users —
    that model and its data aren't public, so no app outside Valve can
    reproduce it exactly.

    What this does instead: it picks out your best-liked games as "seeds"
    — weighted by rating, playtime, and whether you favorited them, so a
    game you rated 9/10 and put 40 hours into counts for a lot more than
    one you just own — then pulls live candidate games from Steam (via
    search on your top genres/studios) and scores each candidate by content
    similarity to your seeds.

    That scoring step is TF-IDF + cosine similarity — the content-based
    technique from content_based.ipynb in IsabelJoseph101's
    steam_recommendation_system project (github.com/IsabelJoseph101/
    steam_recommendation_system): each game's name, genres, categories,
    developer/publisher, and description are turned into one bag-of-words
    "document"; TF-IDF over 1-2 word n-grams weights the distinctive words
    (a rare exact studio-name match ends up carrying a lot of weight on its
    own); cosine similarity between a candidate's vector and each seed's
    vector, averaged and weighted by how much you liked that seed, gives
    the final match score. The notebook does this with scikit-learn, but
    that pulls in numpy/scipy too (~190+ MB) for one small feature, so it's
    reimplemented here in plain Python (see the TF-IDF helper functions
    near the top of this file) — no new dependency at all. If your library
    has no rated/played/favorited games yet to build a taste profile from,
    this falls back to simple genre/studio set-overlap matching instead.

    Note: that repo also has a collaborative-filtering model (KNN/SVD over
    many users' playtime data). True collaborative filtering needs ratings
    from lots of *other* users to find people with similar taste to compare
    you against, and this app only ever sees your own local library, with
    no access to who-else-owns/reviewed-what — Steam doesn't expose that
    (see the module-level comment above steamspy_top_tags(), a bit further
    up this file, for what public data *is* available and how it's used
    below as the closest practical stand-in: real community-voted tags and
    by-tag game lists from SteamSpy, folded into the seed documents and the
    candidate pool respectively).
    """
    force = request.args.get("refresh") == "1"
    try:
        offset = max(0, int(request.args.get("offset", 0)))
    except (TypeError, ValueError):
        offset = 0
    try:
        limit = max(1, min(48, int(request.args.get("limit", 12))))
    except (TypeError, ValueError):
        limit = 12
    games = load_library()
    owned = {int(g["appid"]) for g in games}
    if not games:
        return jsonify({"recommendations": [], "based_on": [], "has_more": False})

    def affinity(g):
        """How much you seem to like this game — higher rating, more
        playtime, and favoriting all push a game further toward being a
        seed for recommendations."""
        score = 0.0
        if g.get("rating") is not None:
            try:
                score += (float(g["rating"]) / 10.0) * 2.0
            except (TypeError, ValueError):
                pass
        playtime_hours = float(g.get("playtime_seconds") or 0) / 3600.0
        if playtime_hours > 0:
            # log-scaled so a handful of very-long-played games don't
            # completely drown out everything else
            score += min(1.5, math.log10(playtime_hours + 1) * 0.75)
        if g.get("favorite"):
            score += 1.0
        return score

    seeds = sorted(((g, affinity(g)) for g in games), key=lambda pair: -pair[1])
    seeds = [(g, w) for g, w in seeds if w > 0][:6]

    genre_weights = Counter()
    studio_weights = Counter()
    seed_names = []
    for g, w in seeds:
        seed_names.append(g.get("name"))
        for genre in (g.get("steam_genres") or [])[:3]:
            genre_weights[genre] += w
        for dev in (g.get("developers") or []):
            studio_weights[dev] += w
        for pub in (g.get("publishers") or []):
            studio_weights[pub] += w * 0.5

    if not genre_weights:
        # Nothing rated, played, or favorited yet to build a real taste
        # profile from — fall back to a flat genre tally across the whole
        # library rather than showing nothing.
        for g in games:
            for genre in (g.get("steam_genres") or [])[:3]:
                genre_weights[genre] += 1

    # Same idea, but for community tags, so the cold-start fallback below
    # (used whenever you have no rated/played/favorited seeds) can match on
    # tags too, not just genres. Uses each library game's already-stored
    # steam_tags — no extra network calls here.
    tag_weights = Counter()
    for g, w in seeds:
        for tag in (g.get("steam_tags") or [])[:5]:
            tag_weights[tag] += w
    if not tag_weights:
        for g in games:
            for tag in (g.get("steam_tags") or [])[:5]:
                tag_weights[tag] += 1
    top_tags = [name for name, _ in tag_weights.most_common(8)]

    if not genre_weights:
        return jsonify({"recommendations": [], "based_on": [], "has_more": False})

    top_genres = [name for name, _ in genre_weights.most_common(6)]
    top_studios = [name for name, _ in studio_weights.most_common(5)]
    based_on = seed_names if seed_names else top_genres

    cache = _load_dlc_cache()  # reuse the on-disk cache file store for this too
    # "v2" bumps the key on purpose: an older cached run (from before the
    # candidate pool was widened) would otherwise keep being served as-is
    # even though it only has a dozen or so games in it, since the genre/
    # studio/tag lists that make up the rest of the key can end up
    # identical to before.
    cache_key = "v2|" + ",".join(sorted(top_genres)) + "|" + ",".join(sorted(top_studios)) + "|" + ",".join(sorted(top_tags))
    rec_cache = cache.get("_recommendations", {})
    if (not force and rec_cache.get("key") == cache_key and
            time.time() - float(rec_cache.get("_cached_at", 0)) < DLC_CACHE_TTL):
        all_games = rec_cache.get("games", [])
        page = all_games[offset:offset + limit]
        return jsonify({
            "recommendations": page,
            "based_on": based_on,
            "has_more": offset + limit < len(all_games),
        })

    candidate_ids = []
    seen = set(owned)
    for term in top_genres + top_studios:
        for item in steam_search(term):
            aid = item.get("id")
            if aid is None or aid in seen:
                continue
            seen.add(aid)
            candidate_ids.append(aid)
        time.sleep(0.15)

    # Widen the pool with SteamSpy's own by-tag game lists — games the wider
    # player *community* tags the same way as your liked games, which is
    # closer to "people into similar games also have this" than a plain
    # Steam-store genre-string search is. Every seed already has its own
    # community tags stored (steam_tags, fetched when it was added/
    # refreshed into the library), so this pulls from all of them at no
    # extra network cost — only the resulting steamspy_games_for_tag()
    # lookups are throttled network calls, one per distinct tag across your
    # seeds (capped below; each one costs ~1.1s through SteamSpy's
    # throttle, so this is kept well under what candidate scoring itself
    # costs now that scoring no longer hits SteamSpy per-candidate).
    widen_tags = []
    seen_tags = set()
    for g, _w in seeds:
        for tag in (g.get("steam_tags") or [])[:3]:
            key = tag.lower()
            if key not in seen_tags:
                seen_tags.add(key)
                widen_tags.append(tag)
    for tag in widen_tags[:6]:
        for aid in steamspy_games_for_tag(tag):
            if aid in seen:
                continue
            seen.add(aid)
            candidate_ids.append(aid)

    candidate_ids = candidate_ids[:70]

    candidates = []
    with ThreadPoolExecutor(max_workers=8) as pool:
        # include_tags=False here — fetching each candidate's own SteamSpy
        # tags used to funnel through the single ~1/sec SteamSpy throttle
        # regardless of thread count, so scoring 30-40 candidates could take
        # 30-40+ seconds the first time a given taste profile was scored.
        # Candidates are scored on Steam's own genres/categories/developer/
        # publisher/description instead, which is fetched concurrently
        # (8 workers) and isn't throttled, so this step now takes a couple
        # of seconds instead of tens of seconds. Community tags are still
        # used on the *seed* side (your own library games already have
        # steam_tags cached from when they were added) and to widen the
        # candidate pool above, just not fetched fresh per candidate here.
        futures = {pool.submit(steam_details, aid, False): aid for aid in candidate_ids}
        for fut in as_completed(futures):
            d = fut.result()
            if d and d.get("name"):
                candidates.append(d)

    def game_doc(d, include_steamspy=False):
        """Bag-of-words text for one game, mirroring the tag-building step
        in content_based.ipynb: name + genres + categories + short
        description, lowercased. Developers/publishers are folded in too —
        TF-IDF gives a rare, exact-matching studio name a lot of weight on
        its own (high IDF), which reproduces that notebook's "same studio"
        signal without a hand-coded rule. include_steamspy also folds in
        this game's SteamSpy community tags where available — real
        player-voted tags from across all of Steam, not just this app's own
        genre/category metadata. For seeds (your own library games) these
        are already stored as steam_tags (fetched when the game was added/
        refreshed), so no extra network call is needed here. Candidates are
        NOT fetched with their own SteamSpy tags (see the ThreadPoolExecutor
        above — that used to cost up to ~40 throttled SteamSpy calls, one
        per candidate, the first time a taste profile was scored), so this
        field is empty for the candidate side; scoring falls back to Steam's
        own genres/categories/developer/publisher/description for
        candidates, which is enough signal and keeps this fast."""
        parts = (
            [d.get("name", "")]
            + (d.get("steam_genres") or [])
            + (d.get("steam_categories") or [])
            + (d.get("developers") or [])
            + (d.get("publishers") or [])
            + [d.get("description", "") or ""]
        )
        if include_steamspy:
            parts += (d.get("steam_tags") or [])
        return " ".join(str(p) for p in parts).lower()

    weighted_sims = None
    if seeds and candidates:
        # Content-based scoring adapted from content_based.ipynb: TF-IDF
        # over each game's text (word 1-2 grams, stopwords stripped — see
        # the pure-Python TF-IDF helper near the top of this file), then
        # cosine similarity between each candidate and each seed, weighted
        # by how much you liked that seed (see affinity() above) and
        # averaged.
        seed_docs = [game_doc(g, include_steamspy=True) for g, _w in seeds]
        seed_weights = [w for _g, w in seeds]
        cand_docs = [game_doc(d, include_steamspy=True) for d in candidates]
        total_weight = sum(seed_weights) or 1.0
        vectors = _tfidf_vectors(seed_docs + cand_docs)
        seed_vecs, cand_vecs = vectors[:len(seed_docs)], vectors[len(seed_docs):]
        weighted_sims = [
            sum(_cosine(cv, sv) * w for sv, w in zip(seed_vecs, seed_weights)) / total_weight
            for cv in cand_vecs
        ]

    scored = []
    for i, d in enumerate(candidates):
        if weighted_sims is not None:
            score = weighted_sims[i]
            if score < 0.05:
                continue
        else:
            # No rated/played/favorited seeds to build a text profile from
            # (fresh library) — fall back to plain genre/tag/studio
            # set-overlap matching using the flat tallies computed above.
            genre_match = set(d.get("steam_genres") or []) & set(top_genres)
            tag_match = set(d.get("steam_tags") or []) & set(top_tags)
            studio_match = set((d.get("developers") or []) + (d.get("publishers") or [])) & set(top_studios)
            if not genre_match and not tag_match and not studio_match:
                continue
            # Genre + community-tag overlap, plus a strong bonus for a
            # same-studio hit — "another game from the people who made the
            # game you loved" is usually a better signal than tag overlap
            # alone.
            score = len(genre_match) + len(tag_match) + (2.0 if studio_match else 0.0)
        scored.append({
            "appid": d["appid"],
            "name": d["name"],
            "image": d.get("image", ""),
            "price": d.get("price"),
            "currency": d.get("currency"),
            "discount": d.get("discount", 0),
            "steam_genres": d.get("steam_genres", []),
            "match_score": score,
        })

    scored.sort(key=lambda x: (-x["match_score"], x["name"]))

    cache["_recommendations"] = {"key": cache_key, "games": scored, "_cached_at": time.time()}
    _save_dlc_cache(cache)

    page = scored[offset:offset + limit]
    return jsonify({
        "recommendations": page,
        "based_on": based_on,
        "has_more": offset + limit < len(scored),
    })


@app.route("/api/add", methods=["POST"])
def api_add():
    appid = request.json.get("appid")
    games = load_library()
    if any(g["appid"] == appid for g in games):
        return jsonify({"error": "already_added"}), 400
    details = steam_details(appid, include_tags=True)
    if not details:
        return jsonify({"error": "not_found"}), 404
    selected_dlcs = request.json.get("dlcs", []) or []
    selected_plan = request.json.get("bundle_plan") or {"cost": 0, "bundles": [], "individual": []}
    # Keep only the DLC data the user selected, plus the exact cheapest acquisition
    # plan chosen by the UI (bundles/editions + uncovered individual DLCs).
    details["dlcs"] = [{
        "appid": int(d.get("appid")),
        "name": str(d.get("name", "Unknown")),
        "image": str(d.get("image", "")),
        "price": d.get("price"),
        "currency": d.get("currency"),
        "discount": d.get("discount", 0),
    } for d in selected_dlcs if d.get("appid") is not None]
    details["bundle_plan"] = selected_plan
    details.update({
        "rating": None,
        "played": False,
        "notes": "",
        "pros": "",
        "cons": "",
        "category": "",
        "favorite": False,
        "custom_categories": [],
        "playtime_seconds": 0,
    })
    games.append(details)
    save_library(games)
    return jsonify(games)


@app.route("/api/add_local", methods=["POST"])
def api_add_local():
    """Add a single local .exe (Epic, Rockstar, GOG, itch.io, a standalone
    game — anything not launched through Steam) as its own library entry,
    with the same Play/Stop, rating, favorite, playtime and category
    features as the Steam games. If the .exe's filename confidently matches
    a Steam listing, the entry uses that game's real name/artwork/tags;
    otherwise it falls back to the .exe's own filename. Each entry gets a
    negative synthetic appid so it never collides with a real Steam appid,
    but otherwise flows through the exact same /api/update, /api/remove,
    /api/game_status etc. routes."""
    data = request.json or {}
    path = (data.get("path") or "").strip()
    if not path:
        return jsonify({"error": "missing_path"}), 400
    if not os.path.isfile(path):
        return jsonify({"error": "not_found"}), 404
    if not path.lower().endswith(".exe"):
        return jsonify({"error": "not_exe"}), 400

    games = load_library()
    norm_path = os.path.normcase(os.path.normpath(path))
    if any(g.get("source") == "local" and os.path.normcase(os.path.normpath(g.get("exe_path", ""))) == norm_path
           for g in games):
        return jsonify({"error": "already_added"}), 400

    existing_ids = {g["appid"] for g in games}
    entry = _build_local_game_entry(path, None, existing_ids)
    games.append(entry)
    save_library(games)
    return jsonify(games)


@app.route("/api/attach_local_steam", methods=["POST"])
def api_attach_local_steam():
    """Manually link a local (non-Steam) game to a Steam listing the user
    picks themselves — pulls in that listing's real name, artwork, price,
    genres and community tags. This is the manual counterpart to the
    automatic filename-matching in _build_local_game_entry, for the local
    games it couldn't confidently guess (or guessed wrong) on its own."""
    data = request.json or {}
    appid = data.get("appid")
    steam_appid = data.get("steam_appid")
    if appid is None or steam_appid is None:
        return jsonify({"error": "missing_params"}), 400

    games = load_library()
    game = next((g for g in games if g.get("appid") == appid and g.get("source") == "local"), None)
    if not game:
        return jsonify({"error": "not_found"}), 404

    details = steam_details(steam_appid, include_tags=True)
    if not details:
        return jsonify({"error": "steam_not_found"}), 404

    game["name"] = details["name"]
    game["image"] = details["image"] or game.get("image", "")
    game["price"] = details["price"]
    game["currency"] = details["currency"]
    game["discount"] = details["discount"]
    game["steam_genres"] = details["steam_genres"]
    game["steam_tags"] = details["steam_tags"]
    game["category"] = details["category"]
    game["developers"] = details.get("developers", [])
    game["publishers"] = details.get("publishers", [])
    game["steam_appid"] = steam_appid
    save_library(games)
    return jsonify(games)


@app.route("/api/detach_local_steam", methods=["POST"])
def api_detach_local_steam():
    """Undo an attach (manual or automatic): drop a local game back to its
    bare .exe filename and extracted icon, clearing everything that was
    pulled in from Steam."""
    data = request.json or {}
    appid = data.get("appid")
    games = load_library()
    game = next((g for g in games if g.get("appid") == appid and g.get("source") == "local"), None)
    if not game:
        return jsonify({"error": "not_found"}), 404

    exe_path = game.get("exe_path", "")
    game["name"] = os.path.splitext(os.path.basename(exe_path))[0] if exe_path else game["name"]
    game["image"] = _extract_exe_icon_data_uri(exe_path) or "" if exe_path else ""
    game["price"] = None
    game["currency"] = None
    game["discount"] = 0
    game["steam_genres"] = []
    game["steam_tags"] = []
    game["category"] = ""
    game["developers"] = []
    game["publishers"] = []
    game["steam_appid"] = None
    save_library(games)
    return jsonify(games)


@app.route("/api/local_folders", methods=["GET"])
def api_local_folders():
    return jsonify(load_local_folders())


@app.route("/api/add_local_folder", methods=["POST"])
def api_add_local_folder():
    """Start watching a folder: every .exe found in it (recursively) is
    added as a local game right away, and future /api/scan_local_folders
    calls keep it in sync — new .exes get added, ones that disappeared
    (uninstalled) get removed."""
    path = ((request.json or {}).get("path") or "").strip()
    if not path:
        return jsonify({"error": "missing_path"}), 400
    if not os.path.isdir(path):
        return jsonify({"error": "not_found"}), 404
    path = os.path.normpath(path)

    folders = load_local_folders()
    if any(os.path.normcase(os.path.normpath(f)) == os.path.normcase(path) for f in folders):
        return jsonify({"error": "already_added"}), 400
    folders.append(path)
    save_local_folders(folders)

    games = load_library()
    added, removed = _apply_folder_scan(path, games)
    save_library(games)
    return jsonify({"folders": folders, "games": games, "added": added, "removed": removed})


@app.route("/api/remove_local_folder", methods=["POST"])
def api_remove_local_folder():
    """Stop watching a folder, and remove the local games that came from
    it (games added by hand via /api/add_local, or found under a
    different watched folder, are left alone)."""
    path = ((request.json or {}).get("path") or "").strip()
    if not path:
        return jsonify({"error": "missing_path"}), 400
    path_norm = os.path.normcase(os.path.normpath(path))

    folders = [f for f in load_local_folders() if os.path.normcase(os.path.normpath(f)) != path_norm]
    save_local_folders(folders)

    games = [g for g in load_library()
             if not (g.get("source") == "local" and g.get("watch_folder")
                     and os.path.normcase(os.path.normpath(g["watch_folder"])) == path_norm)]
    save_library(games)
    return jsonify({"folders": folders, "games": games})


@app.route("/api/scan_local_folders", methods=["POST"])
def api_scan_local_folders():
    """Re-scan every watched folder: add any new .exe found, remove any
    local game whose .exe is no longer there. This is the "Refresh" button
    for Local Games — point it at a folder once and keep it in sync with a
    click instead of adding each game by hand."""
    folders = load_local_folders()
    games = load_library()
    total_added = total_removed = 0
    for folder in folders:
        if not os.path.isdir(folder):
            continue  # drive unplugged / folder moved — skip rather than wipe its games
        added, removed = _apply_folder_scan(folder, games)
        total_added += added
        total_removed += removed
    save_library(games)
    return jsonify({"games": games, "added": total_added, "removed": total_removed})


def _scan_installed_steam_apps():
    """Every Steam game actually installed on this machine right now, found
    by reading each library folder's appmanifest_<appid>.acf files -- the
    same manifests the real Steam client writes and keeps up to date.
    Returns {appid: {"name": ..., "installdir": ...}}, or {} if Steam isn't
    installed / no libraries can be found. Best-effort, like the rest of
    the Steam on-disk reading in this file."""
    steam_root = _find_steam_root()
    if not steam_root:
        return {}
    found = {}
    for lib in _library_folders(steam_root):
        steamapps_dir = os.path.join(lib, "steamapps")
        try:
            entries = os.listdir(steamapps_dir)
        except OSError:
            continue
        for fname in entries:
            if not (fname.startswith("appmanifest_") and fname.endswith(".acf")):
                continue
            try:
                with open(os.path.join(steamapps_dir, fname), "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()
            except OSError:
                continue
            m_id = re.search(r'"appid"\s*"(\d+)"', text)
            if not m_id:
                continue
            m_name = re.search(r'"name"\s*"([^"]+)"', text)
            m_dir = re.search(r'"installdir"\s*"([^"]+)"', text)
            appid = int(m_id.group(1))
            found[appid] = {
                "name": m_name.group(1) if m_name else None,
                "installdir": os.path.join(steamapps_dir, "common", m_dir.group(1)) if m_dir else None,
            }
    return found


@app.route("/api/detect_steam_games", methods=["POST"])
def api_detect_steam_games():
    """Scan the local Steam client's own library folders for installed
    games and bring the library in line with what's actually on disk:
    any installed game not already tracked gets added (pulling its real
    name/artwork/price from Steam's store, same as adding by hand), and
    every existing Steam entry's `installed` flag is updated to match.
    Steam entries are never removed here even if uninstalled -- GameVault
    also tracks games you own but don't currently have installed."""
    installed = _scan_installed_steam_apps()
    games = load_library()
    existing_appids = {g["appid"] for g in games if g.get("source") != "local"}

    added = []
    for i, appid in enumerate(sorted(installed)):
        if appid in existing_appids:
            continue
        if added:
            time.sleep(0.6)  # be gentle with Steam's store API across several new games
        info = installed[appid]
        details = steam_details(appid, include_tags=True)
        if not details:
            # Not a real store listing (delisted, a tool/runtime with its own
            # manifest, etc.) -- still track it, using the manifest's own name.
            details = {
                "appid": appid, "name": info.get("name") or f"App {appid}", "image": "",
                "price": None, "currency": None, "discount": 0, "description": "",
                "steam_categories": [], "steam_genres": [], "steam_tags": [],
                "developers": [], "publishers": [], "category": "",
            }
        details.update({
            "rating": None, "played": False, "notes": "", "pros": "", "cons": "",
            "favorite": False, "custom_categories": [], "playtime_seconds": 0,
            "dlcs": [], "bundle_plan": None,
            "installed": True, "install_dir": info.get("installdir"),
        })
        games.append(details)
        added.append(appid)

    for g in games:
        if g.get("source") == "local":
            continue
        info = installed.get(g["appid"])
        g["installed"] = info is not None
        g["install_dir"] = info.get("installdir") if info else None

    save_library(games)
    return jsonify({"games": games, "added": added})


def _find_local_uninstaller(exe_path):
    """Best-effort: an uninstaller sitting next to a local game's .exe --
    the common patterns installers leave behind (Inno Setup's unins000.exe,
    a generic uninstall.exe/uninst.exe). Only ever looks; never invents or
    deletes anything itself."""
    game_dir = os.path.dirname(exe_path)
    patterns = ("unins000.exe", "unins001.exe", "uninstall.exe", "uninst.exe", "uninstaller.exe")
    for d in (game_dir, os.path.dirname(game_dir)):
        if not d or not os.path.isdir(d):
            continue
        try:
            names = {n.lower(): n for n in os.listdir(d)}
        except OSError:
            continue
        for pat in patterns:
            if pat in names:
                return os.path.join(d, names[pat])
    return None


# ---------- manual cleanup for local games with no real uninstaller ----------
#
# When a local game ships no unins000.exe/uninstall.exe of its own, GameVault
# removes it by hand instead of just complaining: delete everything sitting
# next to the .exe (its whole install folder), then also look for and remove
# any leftover per-game folders it left in AppData/ProgramData/Program Files.
# This never touches anything outside those specific folders, and never wipes
# a folder that looks shared/shallow rather than owned by a single game.

_PROTECTED_LEAF_NAMES = {
    "windows", "system32", "program files", "program files (x86)", "programdata",
    "users", "desktop", "downloads", "documents", "appdata", "roaming", "local",
    "locallow", "steam", "steamapps", "common", "games", "public", "onedrive",
}


def _is_safe_to_wipe(path):
    """Refuse to recursively delete anything that isn't clearly a single
    game's own folder -- a drive root, the user's home folder, or any of the
    big shared Windows/Steam folders, however deep the exe happened to sit."""
    if not path:
        return False
    norm = os.path.normpath(os.path.abspath(path))
    if not os.path.isdir(norm):
        return False
    drive, tail = os.path.splitdrive(norm)
    parts = [p for p in tail.split(os.sep) if p]
    if len(parts) < 2:
        return False  # e.g. C:\Games -- too shallow to be one game's folder
    try:
        home = os.path.normpath(os.path.expanduser("~"))
    except Exception:
        home = None
    if home and os.path.normcase(norm) == os.path.normcase(home):
        return False
    if parts[-1].lower() in _PROTECTED_LEAF_NAMES:
        return False
    return True


def _candidate_leftover_roots():
    """Places installers commonly stash per-game data outside the install
    folder itself."""
    roots = []
    appdata = os.environ.get("APPDATA")
    local_appdata = os.environ.get("LOCALAPPDATA")
    if appdata:
        roots.append(appdata)
    if local_appdata:
        roots.append(local_appdata)
        roots.append(os.path.join(os.path.dirname(local_appdata), "LocalLow"))
    programdata = os.environ.get("PROGRAMDATA")
    if programdata:
        roots.append(programdata)
    for env in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        p = os.environ.get(env)
        if p:
            roots.append(p)
    return [r for r in roots if r and os.path.isdir(r)]


def _game_name_tokens(game, exe_path):
    """Normalized (alnum-only, lowercase) names worth matching leftover
    folders against: the library name, the .exe's own filename, and the
    folder the .exe lives in."""
    raw = {game.get("name"), os.path.splitext(os.path.basename(exe_path))[0],
           os.path.basename(os.path.dirname(exe_path))}
    tokens = {_normalize_for_compare(n) for n in raw if n}
    return {t for t in tokens if len(t) >= 3}


def _find_leftover_dirs(game, exe_path):
    """Best-effort scan of AppData/ProgramData/Program Files for folders
    that look like they belong to this game, by name only -- one level
    deep, never recursing into every subfolder on the system."""
    tokens = _game_name_tokens(game, exe_path)
    if not tokens:
        return []
    found = []
    for root in _candidate_leftover_roots():
        try:
            entries = os.listdir(root)
        except OSError:
            continue
        for entry in entries:
            full = os.path.join(root, entry)
            if not os.path.isdir(full):
                continue
            norm_entry = _normalize_for_compare(entry)
            if not norm_entry:
                continue
            if any(tok in norm_entry or norm_entry in tok for tok in tokens):
                found.append(full)
    return found


def _manual_local_cleanup(game, exe_path):
    """Delete everything next to a local game's .exe (its whole install
    folder), plus any leftover per-game folders found in AppData/ProgramData/
    Program Files. Returns the list of paths actually removed. Used only
    when the game ships no uninstaller of its own."""
    removed = []
    game_dir = os.path.dirname(exe_path)
    if _is_safe_to_wipe(game_dir):
        try:
            shutil.rmtree(game_dir, ignore_errors=True)
            removed.append(game_dir)
        except OSError:
            pass
    else:
        # Folder looked shared/shallow -- just take the .exe itself rather
        # than risk wiping something that isn't this game's alone.
        try:
            os.remove(exe_path)
            removed.append(exe_path)
        except OSError:
            pass

    for leftover in _find_leftover_dirs(game, exe_path):
        if not _is_safe_to_wipe(leftover):
            continue
        try:
            shutil.rmtree(leftover, ignore_errors=True)
            removed.append(leftover)
        except OSError:
            pass
    return removed


@app.route("/api/uninstall", methods=["POST"])
def api_uninstall():
    """Uninstall a game. Steam games are handed off to Steam's own
    steam://uninstall/<appid> flow -- Steam shows its normal confirmation
    dialog and does the actual removal; GameVault never deletes Steam game
    files itself. Local games are uninstalled via whatever uninstaller can
    be found next to their .exe; if none exists, GameVault removes it by
    hand instead -- the whole install folder next to the .exe, plus any
    leftover per-game folders it can find in AppData/Program Files."""
    appid = (request.json or {}).get("appid")
    if appid is None:
        return jsonify({"error": "missing_appid"}), 400
    appid = int(appid)

    game = next((g for g in load_library() if g["appid"] == appid), None)
    if not game:
        return jsonify({"error": "not_found"}), 404

    if game.get("source") == "local":
        exe_path = game.get("exe_path")
        if not exe_path or not os.path.isfile(exe_path):
            return jsonify({"error": "exe_not_found",
                             "detail": "The saved .exe path no longer exists."}), 404
        uninstaller = _find_local_uninstaller(exe_path)
        if uninstaller:
            try:
                import subprocess
                subprocess.Popen([uninstaller], cwd=os.path.dirname(uninstaller))
                return jsonify({"ok": True, "method": "uninstaller"})
            except OSError as e:
                return jsonify({"error": "launch_failed", "detail": str(e)}), 500
        removed = _manual_local_cleanup(game, exe_path)
        return jsonify({"ok": True, "method": "manual", "removed_paths": removed})

    if not game.get("installed"):
        return jsonify({"error": "not_installed",
                         "detail": "This game isn't currently installed."}), 400

    url = f"steam://uninstall/{appid}"
    try:
        if sys.platform.startswith("win"):
            os.startfile(url)  # noqa: S606 — intentional steam:// handoff
        elif sys.platform == "darwin":
            os.system(f'open "{url}"')
        else:
            os.system(f'xdg-open "{url}"')
        return jsonify({"ok": True})
    except OSError as e:
        return jsonify({"error": "launch_failed", "detail": str(e)}), 500


@app.route("/api/launch", methods=["POST"])
def api_launch():
    """Launch a game. Local games (Epic/Rockstar/GOG/anything added via
    "Local Games") are run directly from their saved .exe path; everything
    else is handed off to the local Steam client via the steam:// protocol
    (requires Steam to be installed and the game owned/installed there).
    """
    appid = (request.json or {}).get("appid")
    if not appid:
        return jsonify({"error": "missing_appid"}), 400
    appid = int(appid)

    game = next((g for g in load_library() if g["appid"] == appid), None)
    if game and game.get("source") == "local":
        exe_path = game.get("exe_path")
        if not exe_path or not os.path.isfile(exe_path):
            return jsonify({"error": "exe_not_found",
                             "detail": "The saved .exe path no longer exists."}), 404
        try:
            import subprocess
            subprocess.Popen([exe_path], cwd=os.path.dirname(exe_path))
            return jsonify({"ok": True})
        except OSError as e:
            return jsonify({"error": "launch_failed", "detail": str(e)}), 500

    url = f"steam://run/{appid}"
    try:
        if sys.platform.startswith("win"):
            os.startfile(url)  # noqa: S606 — intentional steam:// handoff
        elif sys.platform == "darwin":
            os.system(f'open "{url}"')
        else:
            os.system(f'xdg-open "{url}"')
        return jsonify({"ok": True})
    except OSError as e:
        return jsonify({"error": "launch_failed", "detail": str(e)}), 500


# ---------- finding/stopping a running game ----------
#
# We hand launching off entirely to Steam (steam://run/<appid>), so we never
# get a PID back the way we would from subprocess.Popen. To support a Stop
# button we have to go find the game's process ourselves: locate Steam's
# install folder, read the game's manifest to get its install directory,
# then look for a running process whose executable lives under that folder.
# This is best-effort — it depends on Steam's on-disk layout, which is
# stable but not a documented public API — so every step fails soft.

_STEAM_ROOT_CACHE = None


def _find_steam_root():
    """Best-effort path to the Steam install folder (contains steamapps/)."""
    global _STEAM_ROOT_CACHE
    if _STEAM_ROOT_CACHE and os.path.isdir(_STEAM_ROOT_CACHE):
        return _STEAM_ROOT_CACHE

    candidates = []
    if sys.platform.startswith("win"):
        try:
            import winreg
            for hive, subkey in (
                (winreg.HKEY_CURRENT_USER, r"Software\Valve\Steam"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Valve\Steam"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Valve\Steam"),
            ):
                try:
                    with winreg.OpenKey(hive, subkey) as key:
                        val, _ = winreg.QueryValueEx(key, "SteamPath" if hive == winreg.HKEY_CURRENT_USER else "InstallPath")
                        candidates.append(val)
                except OSError:
                    pass
        except ImportError:
            pass
        candidates += [r"C:\Program Files (x86)\Steam", r"C:\Program Files\Steam"]
    elif sys.platform == "darwin":
        candidates.append(os.path.expanduser("~/Library/Application Support/Steam"))
    else:
        candidates += [os.path.expanduser("~/.steam/steam"), os.path.expanduser("~/.local/share/Steam")]

    for c in candidates:
        if c and os.path.isdir(c):
            _STEAM_ROOT_CACHE = os.path.normpath(c)
            return _STEAM_ROOT_CACHE
    return None


def _library_folders(steam_root):
    """Every Steam library folder (the default one plus any added on other drives)."""
    libs = [steam_root]
    vdf_path = os.path.join(steam_root, "steamapps", "libraryfolders.vdf")
    try:
        with open(vdf_path, "r", encoding="utf-8") as f:
            text = f.read()
        # Minimal VDF scraping: pull every quoted "path" value out of the file
        # rather than pulling in a full VDF parser for one field.
        for m in re.finditer(r'"path"\s*"([^"]+)"', text):
            p = m.group(1).replace("\\\\", "\\")
            if os.path.isdir(p):
                libs.append(p)
    except OSError:
        pass
    return libs


def _game_install_dir(appid):
    """Read appmanifest_<appid>.acf to find the game's install folder name."""
    steam_root = _find_steam_root()
    if not steam_root:
        return None
    for lib in _library_folders(steam_root):
        manifest = os.path.join(lib, "steamapps", f"appmanifest_{appid}.acf")
        if os.path.exists(manifest):
            try:
                with open(manifest, "r", encoding="utf-8") as f:
                    text = f.read()
                m = re.search(r'"installdir"\s*"([^"]+)"', text)
                if m:
                    return os.path.join(lib, "steamapps", "common", m.group(1))
            except OSError:
                pass
    return None


def _running_game_processes(appid, game=None):
    """psutil.Process objects for this game.

    Steam games: any process whose executable lives inside the game's Steam
    install folder (found via appid, as before). Local games (source
    "local" — Epic/Rockstar/GOG/anything added via "Local Games"): the
    single process matching the exact saved .exe path, since there's no
    Steam manifest to read an install folder from.
    """
    if not HAS_PSUTIL:
        return []

    if game is None:
        game = next((g for g in load_library() if g["appid"] == appid), None)

    if game and game.get("source") == "local":
        exe_path = game.get("exe_path")
        if not exe_path:
            return []
        exe_path = os.path.normcase(os.path.normpath(exe_path))
        matches = []
        for proc in psutil.process_iter(["exe"]):
            try:
                exe = proc.info.get("exe")
                if exe and os.path.normcase(os.path.normpath(exe)) == exe_path:
                    matches.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return matches

    install_dir = _game_install_dir(appid)
    if not install_dir:
        return []
    install_dir = os.path.normcase(os.path.normpath(install_dir))
    matches = []
    for proc in psutil.process_iter(["exe"]):
        try:
            exe = proc.info.get("exe")
            if exe and os.path.normcase(os.path.normpath(exe)).startswith(install_dir + os.sep):
                matches.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return matches


@app.route("/api/game_status", methods=["POST"])
def api_game_status():
    """Is a game (previously launched from here) currently running?

    Used to keep the Play/Stop button honest even if the player quits the
    game from inside Steam itself instead of using our Stop button.
    """
    appid = (request.json or {}).get("appid")
    if not appid:
        return jsonify({"error": "missing_appid"}), 400
    if not HAS_PSUTIL:
        return jsonify({"running": None, "reason": "psutil_not_installed"})
    running = bool(_running_game_processes(int(appid)))
    return jsonify({"running": running})


# ---------- background playtime tracker ----------
#
# Polls every library game's install folder for a running process, on its
# own, independent of whether the Play button in this app was ever clicked.
# This is what lets the app notice a game that was started from the Steam
# client directly (or any other way) and still show a Stop button for it,
# and it's also what accumulates playtime in the background: every tick a
# game is seen running, elapsed seconds since the previous tick are added
# to that game's stored total, and every few ticks the totals are flushed
# to library.json so a crash doesn't lose more than a few seconds of it.

_PLAYTIME_LOCK = threading.Lock()
_RUNNING_STATE = {}     # appid(int) -> True/False/None(unknown), last known status
_PENDING_PLAYTIME = {}  # appid(int) -> extra seconds accumulated since last save
_PLAYTIME_POLL_SECONDS = 10


def _flush_pending_playtime(games=None):
    with _PLAYTIME_LOCK:
        pending = dict(_PENDING_PLAYTIME)
        _PENDING_PLAYTIME.clear()
    if not pending:
        return
    # The tracker loop below already has a freshly-loaded copy of the
    # library in hand (it just scanned every game's process for this same
    # tick) — reuse it instead of reading and re-parsing library.json from
    # disk a second time. Callers outside the loop (none currently) still
    # get the old load-it-here behavior by omitting the argument.
    if games is None:
        games = load_library()
    changed = False
    for g in games:
        extra = pending.get(int(g["appid"]))
        if extra:
            g["playtime_seconds"] = float(g.get("playtime_seconds") or 0) + extra
            changed = True
    if changed:
        save_library(games)


def _playtime_tracker_loop():
    if not HAS_PSUTIL:
        return
    last_tick = time.time()
    while True:
        time.sleep(_PLAYTIME_POLL_SECONDS)
        now = time.time()
        elapsed = now - last_tick
        last_tick = now
        try:
            games = load_library()
        except (OSError, ValueError):
            continue
        with _PLAYTIME_LOCK:
            for g in games:
                appid = int(g["appid"])
                running = bool(_running_game_processes(appid, g))
                _RUNNING_STATE[appid] = running
                if running:
                    _PENDING_PLAYTIME[appid] = _PENDING_PLAYTIME.get(appid, 0) + elapsed
        _flush_pending_playtime(games)


def _start_playtime_tracker():
    if HAS_PSUTIL:
        threading.Thread(target=_playtime_tracker_loop, daemon=True).start()


@app.route("/api/running_status")
def api_running_status():
    """Which library games are currently running right now, whether they
    were launched from this app, from Steam itself, or any other way.

    Powered by the same background tracker that accumulates playtime, so
    this is instant (no per-request process scan) and reflects games
    started outside the app too.
    """
    if not HAS_PSUTIL:
        return jsonify({"running": {}, "supported": False})
    with _PLAYTIME_LOCK:
        running = {str(appid): bool(v) for appid, v in _RUNNING_STATE.items() if v}
    return jsonify({"running": running, "supported": True})


@app.route("/api/stop", methods=["POST"])
def api_stop():
    """Stop a running game by terminating its process(es).

    Best-effort: requires psutil and the ability to locate the game's
    install folder under a detected Steam library. If either isn't
    available, we say so instead of silently doing nothing.
    """
    appid = (request.json or {}).get("appid")
    if not appid:
        return jsonify({"error": "missing_appid"}), 400
    if not HAS_PSUTIL:
        return jsonify({"error": "psutil_not_installed",
                         "detail": "Run: pip install psutil"}), 500

    procs = _running_game_processes(int(appid))
    if not procs:
        return jsonify({"ok": True, "stopped": 0, "detail": "No running process found for this game."})

    stopped = 0
    for proc in procs:
        try:
            proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    gone, alive = psutil.wait_procs(procs, timeout=3)
    stopped += len(gone)
    for proc in alive:
        try:
            proc.kill()  # didn't die politely — force it
            stopped += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return jsonify({"ok": True, "stopped": stopped})


@app.route("/api/remove", methods=["POST"])
def api_remove():
    appid = request.json.get("appid")
    games = [g for g in load_library() if g["appid"] != appid]
    save_library(games)
    return jsonify(games)


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    games = load_library()
    updated = []
    for i, g in enumerate(games):
        if g.get("source") == "local":
            if not g.get("steam_appid"):
                # Not linked to a Steam listing — nothing to refresh against.
                updated.append(g)
                continue
            if i > 0:
                time.sleep(0.6)
            details = steam_details(g["steam_appid"], include_tags=True)
            if details:
                g = {
                    **g,
                    "name": details["name"],
                    "image": details["image"] or g.get("image", ""),
                    "price": details["price"],
                    "currency": details["currency"],
                    "discount": details["discount"],
                    "steam_genres": details["steam_genres"],
                    "steam_tags": details["steam_tags"],
                    "category": details["category"],
                    "developers": details.get("developers", []),
                    "publishers": details.get("publishers", []),
                }
            updated.append(g)
            continue
        if i > 0:
            time.sleep(0.6)
        details = steam_details(g["appid"], include_tags=True)
        merged = {**g, **details} if details else g

        dlcs = g.get("dlcs", []) or []
        if dlcs:
            # Re-resolve the complete DLC set on refresh so a bundle's live price
            # is used instead of adding the individual DLC prices.
            fresh = steam_dlc_options(g["appid"], force_refresh=True)
            refreshed_by_id = {int(d["appid"]): d for d in fresh.get("dlcs", [])}
            merged["dlcs"] = [refreshed_by_id.get(int(d.get("appid")), d) for d in dlcs]
            plan = _best_dlc_plan(merged["dlcs"], fresh.get("bundles", []))
            merged["bundle_plan"] = plan if plan.get("bundles") else None
        else:
            merged["dlcs"] = []
            merged["bundle_plan"] = None

        updated.append(merged)

    save_library(updated)
    return jsonify(updated)


@app.route("/api/update", methods=["POST"])
def api_update():
    data = request.json or {}
    appid = data.get("appid")
    games = load_library()
    found = False
    for g in games:
        if g["appid"] == appid:
            found = True
            if "rating" in data:
                rating = data["rating"]
                if rating is not None:
                    try:
                        rating = max(0, min(10, round(float(rating), 1)))
                    except (TypeError, ValueError):
                        rating = None
                g["rating"] = rating
            if "notes" in data:
                g["notes"] = str(data["notes"] or "")
            if "pros" in data:
                g["pros"] = str(data["pros"] or "")
            if "cons" in data:
                g["cons"] = str(data["cons"] or "")
            # Category is intentionally not accepted from the client.
            # It is derived from Steam's own genres/categories.
            if "played" in data:
                g["played"] = bool(data["played"])
            if "favorite" in data:
                g["favorite"] = bool(data["favorite"])
            if "custom_categories" in data:
                allowed = set(load_custom_categories())
                g["custom_categories"] = [str(c) for c in (data["custom_categories"] or []) if str(c) in allowed]
            break
    if not found:
        return jsonify({"error": "not_found"}), 404
    save_library(games)
    return jsonify(games)


@app.route("/api/set_app_icon", methods=["POST"])
def api_set_app_icon():
    """Set a custom app icon.

    The custom icon is stored persistently and the existing taskbar/Start
    shortcut is updated immediately. The embedded icon of the EXE itself is
    intentionally NOT rebuilt from inside the running one-file EXE: doing
    that can corrupt PyInstaller's bundled archive. Use rebuild_gamevault.bat
    after the app closes to embed the current icon into a fresh EXE.
    """
    data = request.get_json(silent=True) or {}
    raw_bytes, err = _read_image_bytes_from_request(data)
    if err:
        return jsonify({"error": err}), 400
    try:
        if HAS_PIL:
            from PIL import Image
            import io
            img = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
            img = _square_crop_and_resize(img, 256)
            img.save(CUSTOM_ICON_FILE, format="PNG")
        else:
            with open(CUSTOM_ICON_FILE, "wb") as f:
                f.write(raw_bytes)
    except Exception:
        return jsonify({"error": "bad_image"}), 400
    ico_path = _write_app_icon_files()
    _ensure_gamevault_shortcuts(ico_path)
    # Do NOT rebuild or replace the running one-file EXE here. PyInstaller
    # one-file executables must be rebuilt as a separate process after the
    # app is fully closed. The custom icon is already persistent and the
    # taskbar/Start shortcut is updated immediately.
    return jsonify({"ok": True, "data_uri": _current_icon_data_uri()})


@app.route("/api/reset_app_icon", methods=["POST"])
def api_reset_app_icon():
    """Remove the custom icon and restore the built-in icon.

    The persistent custom icon and any existing pin are updated immediately.
    The actual EXE resource is changed only by a fresh PyInstaller build.
    """
    for path in (CUSTOM_ICON_FILE, CUSTOM_ICON_ICO_FILE):
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            pass
    ico_path = _write_app_icon_files()
    _ensure_gamevault_shortcuts(ico_path)
    # Rebuilding the embedded EXE icon is intentionally manual; see
    # rebuild_gamevault.bat. Never replace a running PyInstaller one-file EXE.
    return jsonify({"ok": True, "data_uri": _current_icon_data_uri()})


@app.route("/api/theme_from_image", methods=["POST"])
def api_theme_from_image():
    """Pick an accent color out of a picture ({"path": ...} or
    {"image_data": ...}) and return it as a hex string for the client to
    hand to applyTheme(), same as clicking a theme swatch."""
    if not HAS_PIL:
        return jsonify({"error": "pillow_missing"}), 400
    data = request.get_json(silent=True) or {}
    raw_bytes, err = _read_image_bytes_from_request(data)
    if err:
        return jsonify({"error": err}), 400
    try:
        hex_color = _extract_theme_color_from_bytes(raw_bytes)
    except Exception:
        return jsonify({"error": "bad_image"}), 400
    return jsonify({"ok": True, "hex": hex_color})


# ---------- self-update (from GitHub Releases) ----------
# Where updates are published. GameVault.iss's AppId never changes, so a
# newer GameVault_Setup.exe run on top of an existing install upgrades it
# in place (same folder, same shortcuts, library/settings untouched) --
# see GameVault.iss for that half of the flow. This just finds and runs it.
GITHUB_UPDATE_REPO = "skilerias/GameVault"
GITHUB_API_RELEASES = f"https://api.github.com/repos/{GITHUB_UPDATE_REPO}/releases"


# Must match AppId in GameVault.iss (Inno writes DisplayVersion under it).
_INSTALLER_APP_ID = "{8F4F7E4D-8B1A-4D9B-9F2E-7A7D4A6C9D31}"


def _version_file_candidates():
    """Every place the VERSION file could plausibly be: PyInstaller's data
    folder (_MEIPASS; that's <install>\\_internal on PyInstaller 6), the
    install folder itself, its _internal subfolder, and next to this script
    when run from source."""
    bases = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(meipass)
    bases += [
        APP_DIR,
        os.path.join(APP_DIR, "_internal"),
        os.path.dirname(os.path.abspath(__file__)),
    ]
    seen, out = set(), []
    for b in bases:
        path = os.path.join(b, "VERSION")
        if path not in seen:
            seen.add(path)
            out.append(path)
    return out


def _installed_version_from_registry():
    """Version the installer recorded for this install (Inno Setup writes
    DisplayVersion to the uninstall key). Safety net for when VERSION didn't
    make it into the bundle. Windows only; returns '' if not found."""
    try:
        import winreg
    except ImportError:
        return ""
    subkey = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\%s_is1" % _INSTALLER_APP_ID
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ | view) as k:
                    value, _ = winreg.QueryValueEx(k, "DisplayVersion")
                value = str(value).strip()
                if value:
                    return value
            except OSError:
                continue
    return ""


def _resolve_app_version():
    """(version, where_it_came_from). Never raises."""
    for path in _version_file_candidates():
        try:
            # utf-8-sig so a BOM from Windows editors can't corrupt the number
            with open(path, "r", encoding="utf-8-sig") as f:
                v = f.read().strip()
            if v:
                return v, path
        except (OSError, UnicodeDecodeError):
            continue
    if getattr(sys, "frozen", False):
        try:
            v = _installed_version_from_registry()
        except Exception:
            v = ""
        if v:
            return v, "installer registry entry"
    # '0.0.0' so an update always looks newer, rather than the check
    # silently never firing, if the version genuinely can't be found.
    return "0.0.0", "not found"


def get_app_version():
    return _resolve_app_version()[0]


APP_VERSION, APP_VERSION_SOURCE = _resolve_app_version()


def _version_tuple(v):
    """'v1.2.10' / '1.2.10' -> (1, 2, 10). Non-numeric parts are dropped
    so an odd tag (e.g. '1.2.0-beta') still compares reasonably instead
    of blowing up."""
    v = (v or "").strip()
    if v[:1].lower() == "v":
        v = v[1:]
    parts = []
    for chunk in v.split("."):
        digits = "".join(c for c in chunk if c.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) or (0,)


def _version_is_newer(latest, current):
    a, b = _version_tuple(latest), _version_tuple(current)
    length = max(len(a), len(b))
    a = a + (0,) * (length - len(a))
    b = b + (0,) * (length - len(b))
    return a > b


def _tag_has_version(tag):
    """True if the tag has at least one digit we can compare on ('1.0.1',
    'v1.0.1', 'release-1.0.1' all pass). A tag like 'feat' or 'latest' has
    no digits, so _version_tuple would silently read it as 0.0.0 and it
    would never be offered as an update -- this lets the caller flag that
    instead of dropping the release without explanation."""
    v = (tag or "").strip()
    if v[:1].lower() == "v":
        v = v[1:]
    return any(c.isdigit() for c in v)


# Background download/install state, polled by the frontend (added in a
# later step). stage is one of: idle, downloading, launching, done, error.
_update_install_state = {"stage": "idle", "percent": 0, "error": None}


def _version_sort_key(v):
    t = _version_tuple(v)
    return t + (0,) * (8 - len(t))


def _installer_asset_of(release):
    """The installer asset -- whatever .exe the release has attached."""
    for a in (release.get("assets") or []):
        if str(a.get("name", "")).lower().endswith(".exe"):
            return a
    return None


def check_for_update(timeout=8):
    """Look through GITHUB_UPDATE_REPO's releases and compare the newest
    usable one to the running version. Returns a plain dict and never
    raises -- network problems, a repo with no releases yet, etc. all come
    back as a normal {"ok": False, ...} result the caller can show in the UI.

    Looks at the whole release list (not just GitHub's "latest" pointer, which
    depends on creation date and ignores pre-releases) and picks the highest
    version that has an installer .exe attached. Newer releases that had to
    be skipped are reported in "skipped_newer" with the reason, instead of
    silently saying "you're on the latest version".
    """
    base = {"current_version": APP_VERSION}
    try:
        resp = requests.get(
            GITHUB_API_RELEASES,
            params={"per_page": 30},
            headers={"Accept": "application/vnd.github+json"},
            timeout=timeout,
        )
        if resp.status_code == 404:
            return {"ok": False, "error": "no_releases", **base}
        resp.raise_for_status()
        releases = [r for r in (resp.json() or []) if not r.get("draft")]
        if not releases:
            return {"ok": False, "error": "no_releases", **base}

        usable, skipped = [], []
        for r in releases:
            tag = r.get("tag_name") or ""
            asset = _installer_asset_of(r)
            if r.get("prerelease"):
                reason = "is marked as a pre-release"
            elif not asset:
                reason = "has no installer (.exe) attached"
            elif not _tag_has_version(tag):
                reason = "isn't tagged with a version number"
            else:
                reason = None
            if reason is None:
                usable.append((r, asset))
            elif reason == "isn't tagged with a version number" or _version_is_newer(tag, APP_VERSION):
                # An unversioned tag can't be compared to APP_VERSION, so it's
                # always surfaced rather than silently dropped like the other
                # two reasons (which only matter if the release IS newer).
                skipped.append({"tag": tag or "(untagged)", "reason": reason})

        result = {
            **base,
            "ok": True,
            "latest_version": "",
            "update_available": False,
            "notes": "",
            "download_url": None,
            "asset_name": None,
            "release_url": None,
            "skipped_newer": skipped,
        }
        if usable:
            best, asset = max(usable, key=lambda ra: _version_sort_key(ra[0].get("tag_name")))
            latest_version = best.get("tag_name") or ""
            result.update({
                "latest_version": latest_version,
                "update_available": _version_is_newer(latest_version, APP_VERSION),
                "notes": best.get("body") or "",
                "download_url": asset.get("browser_download_url"),
                "asset_name": asset.get("name"),
                "release_url": best.get("html_url"),
            })
        return result
    except requests.RequestException as e:
        return {"ok": False, "error": "network", "detail": str(e), **base}
    except Exception as e:
        return {"ok": False, "error": "unexpected", "detail": str(e), **base}


class _ElevationError(Exception):
    """Raised when the installer couldn't be launched with admin rights."""


# ShellExecuteW return codes <= 32 mean failure (this is a WinAPI quirk:
# success returns an HINSTANCE-shaped value, always > 32). These are the
# ones actually worth telling the user apart.
_SHELLEXEC_ERRORS = {
    0: "not_enough_memory",
    2: "installer_missing",
    5: "elevation_denied",     # UAC blocked it (e.g. non-admin account, no consent rights)
    8: "not_enough_memory",
    26: "sharing_violation",
    1223: "elevation_cancelled",  # user clicked "No" on the UAC prompt
}


def _launch_installer_elevated(installer_path):
    """Launch installer_path asking Windows for admin rights up front.

    GameVault.iss's installer requires admin (PrivilegesRequired=admin).
    subprocess.Popen() launches via CreateProcess, which does NOT show a
    UAC prompt for a program whose manifest demands elevation -- it just
    fails outright (WinError 740, ERROR_ELEVATION_REQUIRED) with no dialog
    at all. From the user's side that looked like GameVault silently dying
    with the update never actually happening. ShellExecuteW with the
    "runas" verb is the API that actually triggers the UAC prompt.
    """
    import ctypes
    result = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", installer_path, None, None, 1  # SW_SHOWNORMAL
    )
    if result <= 32:
        raise _ElevationError(_SHELLEXEC_ERRORS.get(result, "launch_failed_%d" % result))


def _run_update_download_and_launch(download_url, asset_name):
    """Runs in a background thread: download the installer to a temp
    folder, then launch it. GameVault.iss's installer already knows how to
    close a running GameVault and upgrade it in place (see PrepareToInstall
    there), so once it's launched this app's job is done."""
    global _update_install_state
    import subprocess
    try:
        _update_install_state = {"stage": "downloading", "percent": 0, "error": None}
        tmp_dir = tempfile.mkdtemp(prefix="gamevault_update_")
        installer_path = os.path.join(tmp_dir, asset_name or "GameVault_Setup.exe")
        with requests.get(download_url, stream=True, timeout=30) as r:
            r.raise_for_status()
            total = int(r.headers.get("Content-Length") or 0)
            written = 0
            with open(installer_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=262144):
                    if not chunk:
                        continue
                    f.write(chunk)
                    written += len(chunk)
                    if total:
                        _update_install_state["percent"] = min(99, int(written * 100 / total))
        _update_install_state = {"stage": "launching", "percent": 100, "error": None}
        if sys.platform.startswith("win"):
            _launch_installer_elevated(installer_path)
        else:
            # Detached so the installer survives this app closing/being
            # killed a moment later (the installer force-closes GameVault
            # itself too).
            subprocess.Popen([installer_path], close_fds=True)
        _update_install_state = {"stage": "done", "percent": 100, "error": None}
    except _ElevationError as e:
        _update_install_state = {"stage": "error", "percent": 0, "error": str(e)}
    except Exception as e:
        _update_install_state = {"stage": "error", "percent": 0, "error": str(e)}


@app.route("/api/app_update/check")
def api_app_update_check():
    result = check_for_update()
    result["version_source"] = APP_VERSION_SOURCE
    return jsonify(result)


@app.route("/api/app_update/install", methods=["POST"])
def api_app_update_install():
    data = request.json or {}
    download_url = data.get("download_url")
    asset_name = data.get("asset_name")
    if not download_url:
        return jsonify({"error": "missing_download_url"}), 400
    if _update_install_state["stage"] in ("downloading", "launching"):
        return jsonify({"error": "already_running"}), 409
    threading.Thread(target=_run_update_download_and_launch,
                      args=(download_url, asset_name), daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/app_update/install_status")
def api_app_update_install_status():
    return jsonify(_update_install_state)


PAGE = """
<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Game Vault</title>
<link rel="icon" type="image/png" href="__APP_ICON_DATA_URI__">
<style>
  :root{
    --bg:#14181f; --bg-elev:#1c222c; --bg-elev-2:#232a36;
    --line:#2c3442; --ink:#e7ebf0; --ink-dim:#8b96a5;
    /* --accent-rgb drives the whole accent color (nav highlights, borders,
       focus states, etc). --teal stays as the readable alias everything
       else in this stylesheet already refers to; changing --accent-rgb at
       runtime (see applyTheme() below) re-themes every rule that uses
       var(--teal) or rgba(var(--accent-rgb),...) without touching this file. */
    --accent-rgb:63,143,138; --teal:rgb(var(--accent-rgb));
    --gold:#d8a84e; --danger:#c1614f;
  }
  *{box-sizing:border-box}
  html{background:var(--bg)}
  body{margin:0;background:var(--bg);color:var(--ink);font-family:"Segoe UI",Tahoma,Geneva,sans-serif}
  .app-shell{min-height:100vh;display:grid;grid-template-columns:210px minmax(0,1fr);transition:grid-template-columns .22s ease}
  .app-shell.nav-hidden{grid-template-columns:0 minmax(0,1fr)}
  .side-nav{background:#10141a;border-right:1px solid var(--line);padding:18px 10px;overflow-x:hidden;overflow-y:auto;transition:opacity .18s ease,padding .22s ease,border .22s ease;position:sticky;top:0;align-self:start;height:100vh;display:flex;flex-direction:column;z-index:30}
  .app-shell.nav-hidden .side-nav{opacity:0;padding-left:0;padding-right:0;border-right:0}
  .side-nav-inner{width:100%}
  .side-nav-bottom{width:100%;margin-top:auto;padding-top:14px;border-top:1px solid var(--line)}
  .side-nav-top{display:flex;justify-content:flex-end;margin-bottom:22px}
  /* fixed (not absolute): these stay pinned to the viewport instead of
     scrolling away with the page when you scroll down */
  .nav-toggle{position:fixed;top:18px;left:0;z-index:1000;width:30px;height:42px;background:var(--bg-elev);border:1px solid var(--line);border-left:0;color:var(--ink-dim);border-radius:0 9px 9px 0;cursor:pointer;font-size:1rem;box-shadow:0 3px 10px rgba(0,0,0,.2)}
  .nav-toggle:hover{border-color:var(--teal);color:var(--ink)}
  .fullscreen-toggle{position:fixed;top:18px;right:18px;z-index:1000;width:38px;height:38px;display:inline-flex;align-items:center;justify-content:center;background:var(--bg-elev);border:1px solid var(--line);color:var(--ink-dim);border-radius:9px;cursor:pointer;font-size:1.05rem;box-shadow:0 3px 10px rgba(0,0,0,.2)}
  .fullscreen-toggle:hover{border-color:var(--teal);color:var(--ink)}
  .nav-item{width:100%;display:flex;align-items:center;gap:10px;border:1px solid rgba(var(--accent-rgb),.45);background:rgba(var(--accent-rgb),.16);color:var(--ink);border-radius:9px;padding:10px;font:600 .86rem inherit;white-space:nowrap}
  .nav-item .nav-icon{color:var(--gold);font-size:1.05rem;min-width:17px;text-align:center}
  .nav-item .nav-text{overflow:hidden}
  .category-nav{margin-top:12px;display:grid;gap:5px;max-height:calc(100vh - 115px);overflow-y:auto;padding-right:2px}
  .category-nav-title{color:var(--ink-dim);font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;padding:8px 4px 3px}
  .category-btn{width:100%;display:flex;align-items:center;justify-content:space-between;gap:8px;border:1px solid transparent;background:transparent;color:var(--ink-dim);border-radius:8px;padding:8px 9px;font:600 .8rem inherit;text-align:left;cursor:pointer}
  .category-btn:hover{background:var(--bg-elev);color:var(--ink)}
  .category-btn.active{background:rgba(var(--accent-rgb),.18);border-color:rgba(var(--accent-rgb),.45);color:#fff}
  .category-count{font-size:.7rem;opacity:.7}
  .nav-button{cursor:pointer;text-align:left;font-family:inherit;margin-bottom:8px}
  .nav-button.active{background:rgba(var(--accent-rgb),.24);border-color:rgba(var(--accent-rgb),.65)}
  .main-menu-page{display:none;min-height:70vh;align-items:center;justify-content:center}
  .main-menu-page.active{display:flex}
  .main-menu-grid{display:grid;grid-template-columns:repeat(2,minmax(170px,1fr));gap:16px;width:100%;max-width:560px}
  .main-menu-btn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:30px 16px;background:var(--bg-elev);border:1px solid var(--line);border-radius:14px;color:var(--ink);font:600 .92rem inherit;font-family:inherit;cursor:pointer;text-align:center;transition:border-color .15s ease,transform .15s ease,background .15s ease}
  .main-menu-btn:hover{border-color:rgba(var(--accent-rgb),.65);background:var(--bg-elev-2);transform:translateY(-2px)}
  .main-menu-icon{font-size:1.7rem;color:var(--gold)}
  @media(max-width:520px){.main-menu-grid{grid-template-columns:1fr}}
  .library-toolbar{display:flex;align-items:center;justify-content:space-between;margin:22px 0 12px;gap:12px}
  .local-add-row{display:flex;align-items:center;gap:10px;margin-top:16px;flex-wrap:wrap}
  .local-add-row #localExePath{flex:2;min-width:220px}
  .local-add-row #localFolderPath{flex:2;min-width:220px}
  .toolbar-label{font-size:.92rem;font-weight:700;color:var(--ink)}
  .filter-search{flex:1;min-width:140px;max-width:340px;background:var(--bg-elev);border:1px solid var(--line);color:var(--ink);border-radius:8px;padding:9px 12px;font-size:.85rem;font-family:inherit}
  .filter-search::placeholder{color:var(--ink-dim)}
  .sort-wrap{display:flex;align-items:center;gap:8px}
  .sort-wrap label{font-size:.78rem;color:var(--ink-dim)}
  .sort-select{background:var(--bg-elev);border:1px solid var(--line);color:var(--ink);border-radius:8px;padding:8px 10px;font:inherit;font-size:.82rem;cursor:pointer}
  .sort-select:focus{outline:none;border-color:var(--teal)}
  .installed-toggle{display:flex;align-items:center;gap:7px;font-size:.82rem;color:var(--ink-dim);cursor:pointer;white-space:nowrap;user-select:none}
  .installed-toggle input{width:auto;cursor:pointer;accent-color:var(--teal)}
  .categories-page{display:none}
  .categories-page.active{display:block}
  .theme-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(74px,1fr));gap:14px;margin-top:18px}
  .theme-swatch{position:relative;width:100%;aspect-ratio:1;border-radius:12px;border:2px solid transparent;cursor:pointer;padding:0;box-shadow:0 2px 6px rgba(0,0,0,.35)}
  .theme-swatch:hover{transform:translateY(-2px)}
  .theme-swatch.active{border-color:#fff;box-shadow:0 0 0 3px rgba(255,255,255,.18),0 2px 6px rgba(0,0,0,.35)}
  .theme-swatch.active::after{content:'✓';position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#fff;font-size:1.2rem;font-weight:800;text-shadow:0 1px 3px rgba(0,0,0,.6)}
  .theme-swatch-label{margin-top:6px;font-size:.68rem;color:var(--ink-dim);text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .theme-swatch-cell{display:flex;flex-direction:column;align-items:center}
  .appearance-block{margin-top:26px;padding-top:22px;border-top:1px solid var(--line)}
  .appearance-block-title{margin:0 0 4px;font-size:1rem}
  .appearance-row{display:flex;align-items:center;gap:12px;margin-top:12px;flex-wrap:wrap}
  .icon-preview{width:48px;height:48px;border-radius:10px;object-fit:cover;border:1px solid var(--line);background:var(--bg-elev)}
  .appearance-status{font-size:.8rem;color:var(--ink-dim)}
  .update-notes{margin-top:10px;font-size:.8rem;color:var(--ink-dim);white-space:pre-wrap;max-height:120px;overflow-y:auto;background:var(--bg-elev);border:1px solid var(--line);border-radius:8px;padding:10px}
  .update-progress-wrap{margin-top:12px}
  .update-progress-bar{width:100%;max-width:360px;height:8px;border-radius:5px;background:var(--bg-elev-2);overflow:hidden}
  .update-progress-fill{height:100%;width:0%;background:var(--teal);transition:width .2s ease}
  .update-progress-label{margin-top:6px}
  .category-box-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:14px;margin-top:24px}
  .category-box{min-height:125px;background:linear-gradient(145deg,var(--bg-elev),var(--bg-elev-2));border:1px solid var(--line);border-radius:13px;padding:18px;cursor:pointer;text-align:left;transition:transform .15s,border-color .15s,background .15s}
  .category-box:hover{transform:translateY(-2px);border-color:rgba(var(--accent-rgb),.75);background:var(--bg-elev-2)}
  .category-box.all-box{border-color:rgba(216,168,78,.4)}
  .category-box-title{font-size:1.05rem;font-weight:700;margin-bottom:10px}
  .category-box-count{font-size:.8rem;color:var(--ink-dim)}
  .category-box-icon{font-size:1.5rem;margin-bottom:10px;color:var(--gold)}
  .single-edit{max-width:950px;margin:0 auto}
  .category-detail{margin-top:8px}
  .category-detail-head{display:flex;align-items:flex-end;justify-content:space-between;gap:18px;margin-bottom:20px}
  .category-detail-head h2{font-size:1.35rem;margin:12px 0 4px}
  .back-category{background:transparent;border:0;color:var(--ink-dim);padding:0;cursor:pointer;font:inherit;font-size:.82rem}
  .back-category:hover{color:var(--ink)}
  .favorite-btn{position:absolute;right:9px;top:9px;width:30px;height:30px;border-radius:50%;border:1px solid rgba(255,255,255,.12);background:rgba(10,12,16,.68);color:var(--ink-dim);cursor:pointer;font-size:1rem;z-index:3}
  .favorite-btn:hover{color:var(--gold);border-color:var(--gold)}
  .favorite-btn.active{color:var(--gold);border-color:rgba(216,168,78,.6);background:rgba(216,168,78,.14)}
  .card-media{position:relative}
  @media(max-width:620px){.library-toolbar,.category-detail-head{align-items:flex-start;flex-direction:column}.sort-wrap{width:100%}.sort-select{flex:1}.filter-search{max-width:none;width:100%}}
  .wrap{max-width:1050px;width:100%;margin:0 auto;padding:28px 24px 60px}
  header{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:20px;gap:12px}
  h1{font-size:1.5rem;margin:0;font-weight:700}
  h1 span{color:var(--gold)}
  h2{font-size:1rem;margin:0 0 12px}
  .count{color:var(--ink-dim);font-size:.9rem;white-space:nowrap}
  .total-card{background:linear-gradient(135deg,var(--bg-elev),var(--bg-elev-2));border:1px solid var(--line);border-radius:14px;padding:22px 20px;margin-bottom:22px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}
  .total-label{color:var(--ink-dim);font-size:.9rem;margin-bottom:4px}
  .total-value{font-size:2.2rem;font-weight:800;color:var(--gold);direction:ltr}
  .refresh-btn,.btn-save{background:var(--teal);color:#fff;border:none;border-radius:10px;padding:10px 18px;font-weight:600;cursor:pointer;font-family:inherit}
  .refresh-btn:disabled{opacity:.6;cursor:not-allowed}
  .search-box{position:relative;margin-bottom:10px}
  input.search-input{width:100%;background:var(--bg-elev);border:1px solid var(--line);color:var(--ink);border-radius:10px;padding:13px 14px;font-size:1rem;font-family:inherit}
  input.search-input:focus,.edit-field input:focus,.edit-field textarea:focus{outline:none;border-color:var(--teal)}
  #results{position:absolute;top:calc(100% + 6px);right:0;left:0;z-index:10;background:var(--bg-elev);border:1px solid var(--line);border-radius:10px;max-height:340px;overflow-y:auto;display:none}
  #localAttachResults{position:absolute;top:calc(100% + 6px);right:0;left:0;z-index:10;background:var(--bg-elev);border:1px solid var(--line);border-radius:10px;max-height:280px;overflow-y:auto;display:none}
  .result-row{display:flex;align-items:center;gap:10px;padding:8px 12px;cursor:pointer}
  .result-row:hover{background:var(--bg-elev-2)}
  .result-row img{width:46px;height:22px;object-fit:cover;border-radius:3px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:14px;margin-top:24px;align-items:start}
  .card{background:var(--bg-elev);border:1px solid var(--line);border-radius:10px;overflow:visible;position:relative;z-index:1}
  .card:hover,.card:focus-within{z-index:20}
  .card-media{position:relative;overflow:hidden;border-radius:10px 10px 0 0}
  .card-media img{width:100%;display:block;aspect-ratio:460/215;object-fit:cover}
  .quick-badges{position:absolute;bottom:6px;left:6px;display:flex;gap:5px;z-index:1}
  .qb{display:inline-flex;align-items:center;justify-content:center;background:rgba(0,0,0,.7);border-radius:6px;padding:2px 7px;font-size:.72rem;font-weight:700}
  .qb.rating{color:#fff;display:flex;gap:6px}
  .rating-dot{width:8px;height:8px;border-radius:50%;display:inline-block;box-shadow:0 0 6px rgba(255,255,255,.12)}
  .qb.playtime{color:var(--gold)}
  .info{padding:10px 12px;display:grid;grid-template-columns:minmax(0,1fr) auto auto;gap:4px 8px;align-items:center;background:var(--bg-elev);border-radius:0 0 10px 10px}
  .name{grid-column:1/-1;font-weight:600;font-size:.92rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .price{color:var(--gold);font-weight:700;direction:ltr;text-align:left;min-width:0;line-height:1.25;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  /* Fixed width (instead of auto/content-based) on purpose: every state
     this button cycles through ("▶ Play", "■ Stop", "Launching…",
     "Stopping…") renders at the exact same box size, so swapping the label
     never changes this column's width in the card's grid — which is what
     was causing the whole row of cards to visibly reflow/shift for a
     moment while a game launched. */
  .play-btn{grid-column:2;grid-row:2;display:inline-flex;align-items:center;justify-content:center;gap:4px;background:rgba(var(--accent-rgb),.16);color:var(--teal);border:1px solid rgba(var(--accent-rgb),.55);border-radius:6px;height:25px;width:88px;padding:0;font:700 .72rem inherit;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .play-btn:hover{background:var(--teal);color:#fff}
  .play-btn:disabled{opacity:.5;cursor:default}
  .play-btn.is-playing{background:rgba(224,90,90,.16);color:#e05a5a;border-color:rgba(224,90,90,.55)}
  .play-btn.is-playing:hover{background:#e05a5a;color:#fff}
  .edit-btn{grid-column:3;grid-row:2;display:inline-flex;align-items:center;justify-content:center;background:transparent;color:var(--ink-dim);border:1px solid var(--line);border-radius:6px;width:27px;height:25px;padding:0;cursor:pointer}
  .edit-btn:hover{color:var(--gold);border-color:var(--gold)}
  .edit-btn svg{width:13px;height:13px;pointer-events:none}
  /* Floats over the cards below instead of pushing them down: absolutely
     positioned off the card's own box, so the card's height (and the grid
     row it sits in) never changes on hover. */
  .hover-panel{position:absolute;top:100%;left:-1px;right:-1px;background:var(--bg-elev-2);border:1px solid var(--line);border-top:1px solid var(--line);border-radius:0 0 10px 10px;box-shadow:0 16px 28px rgba(0,0,0,.45);opacity:0;transform:translateY(-6px);transition:opacity .16s ease,transform .16s ease;pointer-events:none;z-index:5}
  .card:hover .hover-panel,.card:focus-within .hover-panel{opacity:1;transform:translateY(0);pointer-events:auto}
  .hover-panel-inner{display:flex;flex-direction:column;gap:4px;padding:10px 12px}
  .hover-row{font-size:.76rem;line-height:1.3;color:var(--ink);display:flex;gap:5px}
  .hover-row b{color:var(--ink-dim);font-weight:600;flex-shrink:0}
  .hover-row span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .empty{text-align:center;color:var(--ink-dim);padding:40px 10px;border:1px dashed var(--line);border-radius:12px;margin-top:20px}
  .section-note{color:var(--ink-dim);font-size:.86rem;margin:-8px 0 20px}
  .edit-page{display:none}
  .edit-page.active{display:block}
  .edit-top{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:24px}
  .edit-top h1{font-size:1.5rem}
  .edit-layout{display:grid;grid-template-columns:280px minmax(0,1fr);gap:20px;align-items:start}
  .edit-list-box,.edit-form-box{background:var(--bg-elev);border:1px solid var(--line);border-radius:14px;padding:16px}
  .edit-list{display:grid;gap:6px;margin-top:12px;max-height:calc(100vh - 210px);overflow-y:auto}
  .edit-list-item{display:flex;align-items:center;gap:9px;width:100%;text-align:left;background:transparent;border:1px solid transparent;color:var(--ink);border-radius:8px;padding:7px;cursor:pointer;font:inherit}
  .edit-list-item:hover{background:var(--bg-elev-2)}
  .edit-list-item.active{background:rgba(var(--accent-rgb),.18);border-color:rgba(var(--accent-rgb),.45)}
  .edit-list-item span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .edit-search{width:100%;background:var(--bg-elev-2);border:1px solid var(--line);color:var(--ink);border-radius:9px;padding:10px 12px;font:inherit}
  .edit-form{display:none}
  .edit-form.visible{display:block}
  .edit-hero{display:flex;gap:18px;align-items:flex-start;margin-bottom:24px}
  .edit-hero img{width:230px;max-width:45%;aspect-ratio:460/215;object-fit:cover;border-radius:10px;border:1px solid var(--line)}
  .edit-hero-copy{min-width:0}
  .edit-hero-copy h2{font-size:1.3rem;margin:4px 0 8px}
  .steam-tags{display:flex;flex-wrap:wrap;gap:6px}
  .steam-tag{font-size:.72rem;padding:4px 8px;border-radius:999px;background:rgba(var(--accent-rgb),.14);border:1px solid rgba(var(--accent-rgb),.35);color:var(--teal)}
  .steam-description{color:var(--ink-dim);font-size:.85rem;line-height:1.5;margin-top:12px}
  .edit-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
  .edit-field{display:block;font-size:.88rem;color:var(--ink-dim);margin-bottom:8px;font-weight:600}
  .edit-field.full{grid-column:1/-1}
  .edit-field input,.edit-field textarea{width:100%;margin-top:8px;background:var(--bg-elev-2);border:1px solid var(--line);color:var(--ink);border-radius:10px;padding:11px 13px;font-family:inherit;font-size:.95rem;resize:vertical}
  .edit-field input[type=number]{direction:ltr;text-align:left}
  .checkbox-field{display:flex;align-items:center;gap:10px;padding-top:30px}
  .checkbox-field input{width:auto;margin-top:0}
  .edit-playtime{margin-top:8px;background:var(--bg-elev-2);border:1px solid var(--line);color:var(--ink);border-radius:10px;padding:11px 13px;font-size:.95rem;font-weight:700}
  .edit-actions{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-top:10px;border-top:1px solid var(--line);padding-top:20px}
  .btn-cancel,.btn-delete{border:none;border-radius:10px;padding:11px 18px;font-weight:600;cursor:pointer;font-family:inherit;font-size:.9rem}
  .btn-cancel{background:var(--bg-elev-2);color:var(--ink-dim)}
  .btn-delete{background:transparent;border:1px solid var(--line);color:var(--danger)}
  .btn-delete:hover{background:rgba(193,97,79,.12);border-color:var(--danger)}
  .edit-actions-right{display:flex;gap:10px}
  .rating-preview{display:flex;align-items:center;gap:9px;margin-top:8px;font-size:.82rem;color:var(--ink-dim)}
  .rating-preview-dot{width:11px;height:11px;border-radius:50%;display:inline-block}
  .category-desc{color:var(--ink-dim);font-size:.78rem;margin-top:7px}
  @media(max-width:760px){
    .app-shell,.app-shell.nav-hidden{grid-template-columns:1fr}
    .side-nav{border-right:0;border-bottom:1px solid var(--line);padding:9px 14px;display:flex;flex-direction:row;align-items:center;gap:12px;height:auto;overflow:visible}
    .side-nav-inner{display:flex;align-items:center;gap:10px;width:auto;flex:1;overflow-x:auto}
    .side-nav-bottom{width:auto;margin-top:0;margin-left:auto;padding-top:0;padding-left:12px;border-top:0;border-left:1px solid var(--line)}
    .side-nav-top{margin:0}
    .nav-item{width:auto}
    .category-nav{position:absolute;left:0;right:0;top:66px;background:#10141a;border-bottom:1px solid var(--line);padding:8px 14px;z-index:20;display:flex;overflow-x:auto}
    .category-nav-title{display:none}
    .category-btn{width:auto;white-space:nowrap}
    .app-shell.nav-hidden .side-nav{display:none}
    .wrap{padding:22px 16px 50px}
    .edit-layout{grid-template-columns:1fr}
      .edit-grid{grid-template-columns:1fr}
    .edit-field.full{grid-column:auto}
    .checkbox-field{padding-top:0}
  }
  .dlc-modal{position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:2000;display:none;align-items:center;justify-content:center;padding:20px}
  .dlc-modal.open{display:flex}
  .dlc-dialog{width:min(720px,100%);max-height:85vh;background:var(--bg-elev);border:1px solid var(--line);border-radius:15px;box-shadow:0 20px 60px rgba(0,0,0,.45);display:flex;flex-direction:column}
  .dlc-head{padding:18px 20px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;gap:15px}
  .dlc-head h2{margin:0;font-size:1.15rem}.dlc-head p{margin:5px 0 0;color:var(--ink-dim);font-size:.8rem}
  .dlc-actions{display:flex;align-items:center;gap:8px}
  .small-btn{background:var(--bg-elev-2);color:var(--ink);border:1px solid var(--line);border-radius:7px;padding:7px 10px;cursor:pointer;font:inherit;font-size:.78rem}
  .small-btn:hover{border-color:var(--teal)}
  .dlc-list{overflow:auto;padding:10px 14px;display:grid;gap:6px}
  .dlc-row{display:grid;grid-template-columns:auto 46px minmax(0,1fr) auto;gap:10px;align-items:center;padding:8px;border:1px solid transparent;border-radius:9px}
  .dlc-row:hover{background:var(--bg-elev-2)}
  .dlc-row img{width:46px;height:22px;object-fit:cover;border-radius:3px}
  .dlc-row-name{font-size:.84rem;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .dlc-row-price{font-size:.82rem;color:var(--gold);font-weight:700;direction:ltr}
  .dlc-footer{padding:14px 18px;border-top:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;gap:12px}
  .dlc-total{font-size:.9rem;color:var(--ink-dim)}.dlc-total strong{display:block;color:var(--gold);font-size:1.2rem}
  .custom-cat-tools{display:flex;gap:8px;margin:0 0 16px}.custom-cat-tools input{flex:1;background:var(--bg-elev);border:1px solid var(--line);color:var(--ink);border-radius:9px;padding:10px 12px;font:inherit}
  .custom-cat-tools button{background:var(--teal);color:#fff;border:0;border-radius:9px;padding:10px 14px;font-weight:600;cursor:pointer}
  .custom-cat-list{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:18px}
  .custom-cat-chip{display:flex;align-items:center;gap:8px;background:var(--bg-elev);border:1px solid var(--line);border-radius:999px;padding:7px 10px;font-size:.8rem}
  .custom-cat-chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:340px}
  .custom-cat-chip button{background:none;border:0;color:var(--danger);cursor:pointer;font-size:1rem}
  .category-box.custom-box{border-color:rgba(var(--accent-rgb),.4)}
  .edit-categories{grid-column:1/-1}
  .category-checks{display:flex;flex-wrap:wrap;gap:7px;margin-top:8px}
  .category-check{display:flex;align-items:center;gap:6px;background:var(--bg-elev-2);border:1px solid var(--line);border-radius:8px;padding:7px 9px;font-size:.8rem;color:var(--ink)}
  .category-check input{margin:0}
  .total-sub{font-size:.75rem;color:var(--ink-dim);margin-top:4px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
</style>
</head>
<body>
<button class="nav-toggle" id="navToggle" title="Hide navigation" aria-label="Hide navigation">☰</button>
<button class="fullscreen-toggle" id="fullscreenToggle" title="Toggle fullscreen (F11)" aria-label="Toggle fullscreen">⛶</button>

<div class="app-shell" id="appShell">
  <aside class="side-nav">
    <div class="side-nav-inner">
      <div class="side-nav-top"></div>
      <button class="nav-item nav-button" id="navMainMenu"><span class="nav-icon">⌂</span><span class="nav-text">Main Menu</span></button>
      <button class="nav-item nav-button active" id="navLibrary"><span class="nav-icon">＋</span><span class="nav-text">Add / remove games</span></button>
      <button class="nav-item nav-button" id="navLocal"><span class="nav-icon">📁</span><span class="nav-text">Local Games</span></button>
      <button class="nav-item nav-button" id="navCategories"><span class="nav-icon">▦</span><span class="nav-text">Categories</span></button>
      <button class="nav-item nav-button" id="navRecommendations"><span class="nav-icon">✦</span><span class="nav-text">Recommendations</span></button>
      <button class="nav-item nav-button" id="navThemes"><span class="nav-icon">🎨</span><span class="nav-text">Themes</span></button>
    </div>
    <div class="side-nav-bottom">
      <button class="nav-item nav-button" id="navSettings"><span class="nav-icon">⚙</span><span class="nav-text">Settings</span></button>
    </div>
  </aside>

  <main class="wrap">
    <section class="main-menu-page" id="mainMenuPage">
      <div class="main-menu-grid">
        <button class="main-menu-btn" id="mmLibrary"><span class="main-menu-icon">＋</span>Add / remove games</button>
        <button class="main-menu-btn" id="mmLocal"><span class="main-menu-icon">📁</span>Local Games</button>
        <button class="main-menu-btn" id="mmCategories"><span class="main-menu-icon">▦</span>Categories</button>
        <button class="main-menu-btn" id="mmRecommendations"><span class="main-menu-icon">✦</span>Recommendations</button>
        <button class="main-menu-btn" id="mmThemes"><span class="main-menu-icon">🎨</span>Themes</button>
        <button class="main-menu-btn" id="mmSettings"><span class="main-menu-icon">⚙</span>Settings</button>
      </div>
    </section>

    <section id="libraryPage">
      <header>
        <h1>Game <span>Vault</span></h1>
        <div class="count" id="countLabel"></div>
      </header>

      <div class="total-card">
        <div>
          <div class="total-label">Total Value</div>
          <div class="total-value" id="totalValue">$0.00</div>
        </div>
        <button class="refresh-btn" id="refreshBtn">Refresh Prices &amp; Steam Data</button>
      </div>

      <div class="search-box">
        <input class="search-input" id="searchInput" type="text" placeholder="Search for a game to add...">
        <div id="results"></div>
      </div>

      <div class="library-toolbar">
        <div class="toolbar-label">Library</div>
        <input class="filter-search" id="librarySearchInput" type="text" placeholder="Search your library...">
        <label class="installed-toggle" title="Show only games that are currently installed">
          <input type="checkbox" id="installedOnlyCheckbox"> Installed only
        </label>
        <div class="sort-wrap">
          <label for="librarySort">Sort by</label>
          <select id="librarySort" class="sort-select">
            <option value="name-asc">Name (A–Z)</option>
            <option value="name-desc">Name (Z–A)</option>
            <option value="rating-desc">Rating (high → low)</option>
            <option value="rating-asc">Rating (low → high)</option>
            <option value="favorites">Favorites first</option>
            <option value="played">Played first</option>
            <option value="price-desc">Price (high → low)</option>
            <option value="price-asc">Price (low → high)</option>
          </select>
        </div>
      </div>
      <div class="grid" id="gamesGrid"></div>
      <div class="empty" id="emptyState" style="display:none;"></div>
    </section>


    <section class="categories-page" id="localPage">
      <div class="edit-top">
        <div>
          <h1>Local Games</h1>
        </div>
        <button class="refresh-btn" id="localScanBtn" type="button">Refresh watched folders</button>
      </div>

      <div class="local-add-row">
        <input class="filter-search" id="localFolderPath" type="text" placeholder="Folder to watch (e.g. your Epic Games install drive)...">
        <button class="refresh-btn" id="localFolderBrowseBtn" type="button">Browse folder...</button>
        <button class="refresh-btn" id="localFolderAddBtn" type="button">+ Watch folder</button>
      </div>
      <div class="custom-cat-list" id="localFoldersList"></div>

      <div class="local-add-row" style="margin-top:22px">
        <input class="filter-search" id="localExePath" type="text" placeholder="Or add one game: path to its .exe...">
        <button class="refresh-btn" id="localBrowseBtn" type="button">Browse...</button>
        <button class="refresh-btn" id="localAddBtn" type="button">+ Add game</button>
      </div>
      <div class="section-note" id="localAddError" style="display:none;color:#e5786c;margin:6px 0 0"></div>

      <div class="edit-top" style="margin-top:28px">
        <div>
          <h2 style="margin:0 0 4px">Your local games</h2>
          <div class="section-note" id="localGamesCount"></div>
        </div>
      </div>
      <div class="grid" id="localGamesGrid"></div>
      <div class="empty" id="localGamesEmpty" style="display:none;"></div>
    </section>

    <section class="categories-page" id="categoriesPage">
      <div class="edit-top">
        <div>
          <h1>Categories</h1>
          <div class="section-note">Steam genres plus your own custom categories. Add a category, then assign games to it from the game editor.</div>
          <div class="custom-cat-tools"><input id="newCategoryInput" placeholder="New custom category name..."><button id="addCategoryBtn">+ Add category</button></div>
          <div class="custom-cat-list" id="customCatList"></div>
        </div>
      </div>
      <div class="library-toolbar" style="margin-top:8px">
        <div class="toolbar-label">Browse</div>
        <input class="filter-search" id="categoryBoxSearch" type="text" placeholder="Search categories...">
      </div>
      <div class="category-box-grid" id="categoryBoxes"></div>
      <div class="category-detail" id="categoryDetail" style="display:none">
        <div class="category-detail-head">
          <div>
            <button class="back-category" id="backCategories">← All categories</button>
            <h2 id="categoryDetailTitle"></h2>
            <div class="section-note" id="categoryDetailCount"></div>
          </div>
          <input class="filter-search" id="categoryDetailSearch" type="text" placeholder="Search this category...">
          <div class="sort-wrap">
            <label for="categorySort">Sort by</label>
            <select id="categorySort" class="sort-select">
              <option value="name-asc">Name (A–Z)</option>
              <option value="name-desc">Name (Z–A)</option>
              <option value="rating-desc">Rating (high → low)</option>
              <option value="rating-asc">Rating (low → high)</option>
              <option value="favorites">Favorites first</option>
              <option value="played">Played first</option>
              <option value="price-desc">Price (high → low)</option>
              <option value="price-asc">Price (low → high)</option>
            </select>
          </div>
        </div>
        <div class="grid" id="categoryGamesGrid"></div>
      </div>
    </section>

    <section class="categories-page" id="recommendationsPage">
      <div class="edit-top">
        <div>
          <h1>Recommendations</h1>
          <div class="section-note" id="recBasedOn"></div>
        </div>
        <button class="refresh-btn" id="recRefreshBtn">Refresh</button>
      </div>
      <div class="grid" id="recommendationsGrid"></div>
      <div class="empty" id="recEmptyState" style="display:none;"></div>
    </section>

    <section class="categories-page" id="themesPage">
      <div class="edit-top">
        <div>
          <h1>Themes</h1>
          <div class="section-note">Pick an accent color for the whole app. Saved on this device.</div>
        </div>
      </div>
      <div class="theme-grid" id="themeGrid"></div>

      <div class="appearance-block">
        <h2 class="appearance-block-title">Theme from a picture</h2>
        <div class="section-note">Pick a picture and the app will lift an accent color out of it.</div>
        <div class="appearance-row">
          <button type="button" class="small-btn" id="themeImageBtn">Choose picture...</button>
          <input type="file" accept="image/*" id="themeImageInput" style="display:none">
          <span class="appearance-status" id="themeImageStatus"></span>
        </div>
      </div>

      <div class="appearance-block">
        <h2 class="appearance-block-title">App icon</h2>
        <div class="section-note">Use your own picture as the app's icon.</div>
        <div class="appearance-row">
          <img class="icon-preview" id="iconPreview" alt="App icon">
          <button type="button" class="small-btn" id="iconChooseBtn">Choose picture...</button>
          <input type="file" accept="image/*" id="iconFileInput" style="display:none">
          <button type="button" class="small-btn btn-cancel" id="iconResetBtn">Reset to original</button>
        </div>
        <div class="appearance-status" id="iconStatus"></div>
      </div>
    </section>

    <section class="categories-page" id="settingsPage">
      <div class="edit-top">
        <div>
          <h1>Settings</h1>
          <div class="section-note">App-level options.</div>
        </div>
      </div>

      <div class="appearance-block">
        <h2 class="appearance-block-title">App updates</h2>
        <div class="section-note">Current version: <strong id="updateCurrentVersion">—</strong></div>
        <div class="appearance-row">
          <button type="button" class="small-btn" id="checkUpdateBtn">Check for updates</button>
          <button type="button" class="small-btn btn-save" id="installUpdateBtn" style="display:none">Update now</button>
          <span class="appearance-status" id="updateStatus"></span>
        </div>
        <div class="update-notes" id="updateNotes" style="display:none"></div>
        <div class="update-progress-wrap" id="updateProgressWrap" style="display:none">
          <div class="update-progress-bar"><div class="update-progress-fill" id="updateProgressFill"></div></div>
          <div class="appearance-status" id="updateProgressLabel"></div>
        </div>
      </div>
    </section>

    <section class="edit-page" id="editPage">
      <div class="edit-top">
        <div>
          <h1>Edit game</h1>
          <div class="section-note">Edit the game you selected from your library.</div>
        </div>
        <button class="btn-cancel" id="editBack">← Back to library</button>
      </div>

      <div class="edit-form-box single-edit">
        <div class="empty" id="editorEmpty" style="margin:0">No game selected.</div>
        <div class="edit-form" id="editForm">
          <div class="edit-hero">
            <img id="editGameImage" src="" alt="">
            <div class="edit-hero-copy">
              <h2 id="editGameName"></h2>
              <div class="steam-tags" id="steamTags"></div>
              <div class="steam-description" id="steamDescription"></div>
            </div>
          </div>

          <div class="edit-field full local-attach-box" id="localAttachBox" style="display:none">
            Link to a Steam game <span style="opacity:.7;font-weight:400">(fills in its price, artwork, genres &amp; tags)</span>
            <div class="search-box" style="max-width:460px;margin-top:8px">
              <input class="search-input" id="localAttachSearch" type="text" placeholder="Search Steam...">
              <div id="localAttachResults"></div>
            </div>
            <div class="section-note" id="localAttachedNote" style="display:none;margin-top:8px;font-weight:400"></div>
          </div>

          <div class="edit-grid">
            <label class="edit-field">
              Rating (out of 10)
              <input id="editRating" type="number" min="0" max="10" step="0.1" placeholder="e.g. 8.5">
              <div class="rating-preview"><span class="rating-preview-dot" id="ratingPreviewDot"></span><span id="ratingPreviewText">No rating</span></div>
            </label>

            <div class="edit-field">
              Playtime
              <div class="edit-playtime" id="editPlaytime">0m</div>
            </div>

            <label class="edit-field checkbox-field">
              <input id="editFavorite" type="checkbox"> Favorite
            </label>

            <div class="edit-field full edit-categories">Custom categories
              <div class="category-checks" id="editCategoryChecks"></div>
            </div>

            <label class="edit-field full">Pros
              <textarea id="editPros" rows="5" placeholder="What do you like about it?"></textarea>
            </label>

            <label class="edit-field full">Cons
              <textarea id="editCons" rows="5" placeholder="What don't you like?"></textarea>
            </label>

            <label class="edit-field full">Notes
              <textarea id="editNotes" rows="6" placeholder="Your notes..."></textarea>
            </label>
          </div>

          <div class="edit-actions">
            <div>
              <button class="btn-danger" id="editDelete">Remove from library</button>
              <button class="btn-danger" id="editUninstall" style="display:none;margin-left:8px">Uninstall</button>
            </div>
            <div>
              <button class="btn-cancel" id="editCancel">Cancel</button>
              <button class="btn-save" id="editSave">Save changes</button>
            </div>
          </div>
        </div>
      </div>
    </section>
  </main>
</div>

<div class="dlc-modal" id="dlcModal">
  <div class="dlc-dialog">
    <div class="dlc-head">
      <div><h2 id="dlcTitle">Choose DLCs</h2><p>Select the DLCs you own. The base game is included automatically.</p></div>
      <div class="dlc-actions"><button class="small-btn" id="dlcSelectAll">Select all</button><button class="small-btn" id="dlcSelectNone">None</button></div>
    </div>
    <div class="dlc-list" id="dlcList"></div>
    <div class="dlc-footer">
      <div class="dlc-total">Selected DLCs<strong id="dlcTotal">$0.00</strong></div>
      <div><button class="btn-cancel" id="dlcCancel">Cancel</button> <button class="btn-save" id="dlcConfirm">Add game</button></div>
    </div>
  </div>
</div>
<script>
const $ = id => document.getElementById(id);
const searchInput=$('searchInput'), resultsBox=$('results'), grid=$('gamesGrid');
const totalValue=$('totalValue'), countLabel=$('countLabel'), emptyState=$('emptyState');
const categoriesPage=$('categoriesPage'), categoryBoxes=$('categoryBoxes'), categoryDetail=$('categoryDetail');
const recommendationsPage=$('recommendationsPage'), recommendationsGrid=$('recommendationsGrid'), recBasedOn=$('recBasedOn'), recEmptyState=$('recEmptyState');
const navLibrary=$('navLibrary'), navCategories=$('navCategories'), navRecommendations=$('navRecommendations'), navLocal=$('navLocal'), navThemes=$('navThemes'), navSettings=$('navSettings'), navMainMenu=$('navMainMenu'), editPage=$('editPage'), libraryPage=$('libraryPage'), mainMenuPage=$('mainMenuPage');
const themesPage=$('themesPage'), themeGrid=$('themeGrid');
const settingsPage=$('settingsPage');
const themeImageBtn=$('themeImageBtn'), themeImageInput=$('themeImageInput'), themeImageStatus=$('themeImageStatus');
const iconPreview=$('iconPreview'), iconChooseBtn=$('iconChooseBtn'), iconFileInput=$('iconFileInput'), iconResetBtn=$('iconResetBtn'), iconStatus=$('iconStatus');
const updateCurrentVersion=$('updateCurrentVersion'), checkUpdateBtn=$('checkUpdateBtn'), installUpdateBtn=$('installUpdateBtn'), updateStatus=$('updateStatus'), updateNotes=$('updateNotes');
const updateProgressWrap=$('updateProgressWrap'), updateProgressFill=$('updateProgressFill'), updateProgressLabel=$('updateProgressLabel');
const localPage=$('localPage');
const localExePath=$('localExePath'), localBrowseBtn=$('localBrowseBtn'), localAddBtn=$('localAddBtn'), localAddError=$('localAddError');
const localFolderPath=$('localFolderPath'), localFolderBrowseBtn=$('localFolderBrowseBtn'), localFolderAddBtn=$('localFolderAddBtn'), localFoldersList=$('localFoldersList'), localScanBtn=$('localScanBtn');
const localGamesGrid=$('localGamesGrid'), localGamesEmpty=$('localGamesEmpty'), localGamesCount=$('localGamesCount');
const localAttachBox=$('localAttachBox'), localAttachSearch=$('localAttachSearch'), localAttachResults=$('localAttachResults'), localAttachedNote=$('localAttachedNote');
let localAttachSearchTimer, localAttachAbortCtrl;
const editorEmpty=$('editorEmpty'), editForm=$('editForm');
const editGameName=$('editGameName'), editGameImage=$('editGameImage'), steamTags=$('steamTags'), steamDescription=$('steamDescription');
const editRating=$('editRating'), editPlaytime=$('editPlaytime'), editFavorite=$('editFavorite'), editPros=$('editPros'), editCons=$('editCons'), editNotes=$('editNotes');
const librarySort=$('librarySort'), categorySort=$('categorySort');
const ratingPreviewDot=$('ratingPreviewDot'), ratingPreviewText=$('ratingPreviewText');
const dlcModal=$('dlcModal'), dlcList=$('dlcList'), dlcTitle=$('dlcTitle'), dlcTotal=$('dlcTotal');
const editCategoryChecks=$('editCategoryChecks');
const categoryDetailTitle=$('categoryDetailTitle'), categoryDetailCount=$('categoryDetailCount'), categoryGamesGrid=$('categoryGamesGrid');
const librarySearchInput=$('librarySearchInput'), categoryBoxSearch=$('categoryBoxSearch'), categoryDetailSearch=$('categoryDetailSearch');
const installedOnlyCheckbox=$('installedOnlyCheckbox');
const INSTALLED_ONLY_KEY='gamevault_installed_only';
function isGameInstalled(g){ return g.source==='local' || !!g.installed; }
let currentCategory=null, currentCategoryIsCustom=false;
let searchTimer, allGames=[], customCategories=[], currentFilter='All', editingAppid=null, pendingAdd=null, pendingDlc=[], pendingBase=null, pendingBundles=[];
let editReturnTo='library', editReturnScrollY=0;  // where to go back to when the edit page closes
let playingAppids=new Set();  // appids we've launched from here and believe are still running
let pendingAction=new Map();  // appid -> 'launching'|'stopping', so re-renders don't wipe a mid-flight action
let psutilSupported=null;     // whether the server can actually detect running games (null = not known yet)

function fmt(n,cur){
  if(n===null||n===undefined)return'N/A';
  return n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})+' '+(cur||'');
}
// Same convention Steam itself uses: hours+minutes once you're over an hour,
// just minutes below that, and "Not played" at zero — instead of "0h 0m".
function fmtPlaytime(seconds){
  const s=Number(seconds)||0;
  if(s<30)return'Not played yet';
  const totalMinutes=Math.round(s/60);
  const h=Math.floor(totalMinutes/60), m=totalMinutes%60;
  if(h>0)return`${h}h ${m}m`;
  return`${m}m`;
}
function gameTotal(g){
  const base=(g.price===null||g.price===undefined)?0:Number(g.price)||0;
  if(g.bundle_plan && Number.isFinite(Number(g.bundle_plan.cost))){
    return base + Number(g.bundle_plan.cost);
  }
  const dlcs=(g.dlcs||[]).reduce((sum,d)=>sum+(Number(d.price)||0),0);
  return base+dlcs;
}
function dlcCurrency(g){
  return g.currency || ((g.dlcs||[]).find(d=>d.currency)?.currency) || '';
}

/* 0 = dark purple, then purple -> blue -> orange -> yellow -> green at 10. */
function ratingColor(value){
  const n=Number(value);
  if(!Number.isFinite(n)) return '#6b7280';
  const stops=[
    [0,[62,20,92]],[1,[88,31,128]],[2,[119,45,145]],[3,[74,78,170]],
    [4,[45,112,177]],[5,[230,126,34]],[6,[242,158,40]],[7,[242,194,48]],
    [8,[176,196,52]],[9,[86,178,74]],[10,[35,170,82]]
  ];
  const x=Math.max(0,Math.min(10,n));
  const lo=Math.floor(x), hi=Math.min(10,lo+1), t=x-lo;
  const a=stops[lo][1], b=stops[hi][1];
  const rgb=a.map((v,i)=>Math.round(v+(b[i]-v)*t));
  return `rgb(${rgb.join(',')})`;
}

function ratingDot(value,cls='rating-dot'){
  return `<span class="${cls}" style="background:${ratingColor(value)}"></span>`;
}

function escapeHtml(v){
  return String(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// The Categories page shows every genre Steam lists for a game (steam_genres),
// not just a fixed whitelist of "major" ones — a niche genre like "Free to Play"
// or "Early Access" now gets its own browsable category same as "Action" or "RPG".
// Community tags (steam_tags, from SteamSpy) are intentionally NOT part of this —
// they're much noisier/more numerous and would clutter a browsable category list.
// Steam's separate feature/category flags (Single-player, Controller Support,
// Steam Achievements, etc., in steam_categories) are also intentionally not
// shown here — those aren't genres either.
function steamCategories(g){
  return g.steam_genres||[];
}

// Used only next to the game description in the editor: the full picture of
// a game's classification, combining Steam's own genres with SteamSpy's
// community-voted tags (deduped case-insensitively). Not used anywhere in
// the Categories browsing page — see steamCategories() above for that.
function allGameTags(g){
  const combined=[...(g.steam_genres||[]),...(g.steam_tags||[])];
  const seen=new Set();
  const out=[];
  combined.forEach(c=>{
    const key=c.toLowerCase();
    if(!seen.has(key)){seen.add(key);out.push(c);}
  });
  return out;
}

function sortGames(list,mode){
  const arr=[...list];
  const rating=g=>g.rating===null||g.rating===undefined?-1:Number(g.rating);
  const price=g=>g.price===null||g.price===undefined?-1:Number(g.price);
  switch(mode){
    case 'name-desc': return arr.sort((a,b)=>b.name.localeCompare(a.name,'en'));
    case 'rating-desc': return arr.sort((a,b)=>rating(b)-rating(a)||a.name.localeCompare(b.name,'en'));
    case 'rating-asc': return arr.sort((a,b)=>rating(a)-rating(b)||a.name.localeCompare(b.name,'en'));
    case 'favorites': return arr.sort((a,b)=>Number(!!b.favorite)-Number(!!a.favorite)||a.name.localeCompare(b.name,'en'));
    case 'played': return arr.sort((a,b)=>Number(!!b.played)-Number(!!a.played)||a.name.localeCompare(b.name,'en'));
    case 'price-desc': return arr.sort((a,b)=>price(b)-price(a)||a.name.localeCompare(b.name,'en'));
    case 'price-asc': return arr.sort((a,b)=>price(a)-price(b)||a.name.localeCompare(b.name,'en'));
    default: return arr.sort((a,b)=>a.name.localeCompare(b.name,'en'));
  }
}

function buildCategoryBoxes(){
  const boxes=$('categoryBoxes');
  boxes.innerHTML='';
  const q=categoryBoxSearch.value.trim().toLowerCase();

  if(!q||'all'.includes(q)){
    const all=document.createElement('button');
    all.className='category-box all-box';
    all.innerHTML=`<div class="category-box-icon">▦</div><div class="category-box-title">All</div><div class="category-box-count">${allGames.length} ${allGames.length===1?'game':'games'}</div>`;
    all.onclick=()=>openCategory('All');
    boxes.append(all);
  }

  // Every distinct genre present across the whole library gets a box now —
  // not just a fixed "major genres" whitelist — so a niche genre like
  // "Free to Play" shows up here exactly like "Action" or "RPG" does.
  const allGenres=new Set();
  allGames.forEach(g=>steamCategories(g).forEach(cat=>allGenres.add(cat)));
  [...allGenres].sort((a,b)=>a.localeCompare(b,'en')).forEach(cat=>{
    if(q&&!cat.toLowerCase().includes(q))return;
    const games=allGames.filter(g=>steamCategories(g).includes(cat));
    if(!games.length)return;
    const box=document.createElement('button');
    box.className='category-box';
    box.innerHTML=`<div class="category-box-icon">◈</div><div class="category-box-title">${cat}</div><div class="category-box-count">${games.length} ${games.length===1?'game':'games'}</div>`;
    box.onclick=()=>openCategory(cat);
    boxes.append(box);
  });
  customCategories.forEach(cat=>{
    if(q&&!cat.toLowerCase().includes(q))return;
    const games=allGames.filter(g=>(g.custom_categories||[]).includes(cat));
    const box=document.createElement('button');
    box.className='category-box custom-box';
    box.innerHTML=`<div class="category-box-icon">◆</div><div class="category-box-title">${escapeHtml(cat)}</div><div class="category-box-count">${games.length} ${games.length===1?'game':'games'}</div>`;
    box.onclick=()=>openCategory(cat,true);
    boxes.append(box);
  });

  if(!boxes.children.length){
    boxes.innerHTML='<div class="empty">No categories match your search.</div>';
  }
}

function openCategory(category,isCustom=false){
  currentCategory=category;
  currentCategoryIsCustom=isCustom;
  categoryDetailSearch.value='';
  $('categoryBoxes').style.display='none';
  $('categoryDetail').style.display='block';
  $('categoryDetailTitle').textContent=category;
  renderCategoryDetail();
}

function renderCategoryDetail(){
  if(currentCategory===null)return;
  const base=currentCategory==='All'?allGames:(currentCategoryIsCustom?allGames.filter(g=>(g.custom_categories||[]).includes(currentCategory)):allGames.filter(g=>steamCategories(g).includes(currentCategory)));
  const q=categoryDetailSearch.value.trim().toLowerCase();
  const games=base.filter(g=>gameMatchesQuery(g,q));
  const qLabel=categoryDetailSearch.value.trim();
  $('categoryDetailCount').textContent=`${games.length} ${games.length===1?'game':'games'}`+(qLabel?` matching "${qLabel}"`:'');
  renderGameCards($('categoryGamesGrid'),games,$('categorySort').value);
}

function closeCategory(){
  currentCategory=null;
  $('categoryDetail').style.display='none';
  $('categoryBoxes').style.display='grid';
}

function renderGameCards(target,games,sortMode){
  target.innerHTML='';
  const sorted=sortGames(games,sortMode);

  if(!sorted.length){
    target.innerHTML='<div class="empty">No games in this category.</div>';
    return;
  }

  sorted.forEach(g=>{
    const hasRating=g.rating!==null&&g.rating!==undefined;
    const cats=steamCategories(g);
    const dlcCount=(g.dlcs||[]).length;
    const card=document.createElement('div');
    card.className='card';
    card.innerHTML=`
      <div class="card-media" tabindex="0">
        <img src="${g.image||''}" alt="">
        <button class="favorite-btn ${g.favorite?'active':''}" title="${g.favorite?'Remove favorite':'Add to favorites'}" aria-label="${g.favorite?'Remove favorite':'Add to favorites'}">${g.favorite?'★':'☆'}</button>
        <div class="quick-badges">
          ${hasRating?`<span class="qb rating">${ratingDot(g.rating)}★ ${g.rating}</span>`:''}
          ${(g.playtime_seconds||0)>=30?`<span class="qb playtime">⏱ ${fmtPlaytime(g.playtime_seconds)}</span>`:''}
        </div>
      </div>
      <div class="info">
        <div class="name">${g.name}</div>
        <div class="price">${(g.source==='local'&&!g.steam_appid)?'Owned':fmt(gameTotal(g),dlcCurrency(g))}<div class="total-sub"${dlcCount?'':' style="visibility:hidden"'} title="Base game + ${dlcCount||1} DLC included in this price">+${dlcCount||1} DLC</div></div>
        ${playBtnMarkup(g.appid,g.name)}
        <button class="edit-btn" title="Edit ${g.name}" aria-label="Edit ${g.name}">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4Z"/>
          </svg>
        </button>
      </div>
      <div class="hover-panel">
        <div class="hover-panel-inner">
          <div class="hover-row"><b>Rating</b><span>${hasRating?ratingDot(g.rating)+' '+g.rating+'/10':'Not rated'}</span></div>
          <div class="hover-row"><b>Playtime</b><span>${fmtPlaytime(g.playtime_seconds)}</span></div>
          <div class="hover-row"><b>Steam</b><span>${cats.length?cats.join(', '):'No main genre returned'}</span></div>
        </div>
      </div>`;
    card.querySelector('.favorite-btn').onclick=()=>toggleFavorite(g.appid,g.favorite);
    card.querySelector('.edit-btn').onclick=()=>openEditPage(g);
    card.querySelector('.play-btn').onclick=()=>handlePlayClick(g.appid,g.name);
    target.append(card);
  });
}

function renderLocalGamesGrid(){
  if(!localGamesGrid)return;
  const games=allGames.filter(g=>g.source==='local');
  localGamesCount.textContent=games.length+(games.length===1?' game found':' games found');
  if(!games.length){
    localGamesGrid.innerHTML='';
    localGamesEmpty.style.display='block';
    localGamesEmpty.textContent='No local games yet — watch a folder or add one above.';
    return;
  }
  localGamesEmpty.style.display='none';
  renderGameCards(localGamesGrid,games,'name-asc');
}

function gameMatchesQuery(g,q){
  if(!q)return true;
  if(g.name.toLowerCase().includes(q))return true;
  if(steamCategories(g).some(c=>c.toLowerCase().includes(q)))return true;
  if((g.custom_categories||[]).some(c=>c.toLowerCase().includes(q)))return true;
  return false;
}

function render(){
  const q=librarySearchInput.value.trim().toLowerCase();
  const installedOnly=installedOnlyCheckbox&&installedOnlyCheckbox.checked;
  const games=allGames.filter(g=>gameMatchesQuery(g,q)&&(!installedOnly||isGameInstalled(g)));
  grid.innerHTML='';

  if(!allGames.length){
    emptyState.style.display='block';
    emptyState.textContent='No games added yet — search Steam above to add your first game.';
  }else if(!games.length){
    emptyState.style.display='block';
    emptyState.textContent='No matching games.';
  }else emptyState.style.display='none';

  countLabel.textContent=games.length+(games.length===1?' game':' games');
  let totalsByCurrency={};
  allGames.forEach(g=>{
    const cur=dlcCurrency(g);
    if(cur)totalsByCurrency[cur]=(totalsByCurrency[cur]||0)+gameTotal(g);
  });
  renderGameCards(grid,games,librarySort.value);
  totalValue.textContent=Object.entries(totalsByCurrency).map(([cur,val])=>fmt(val,cur)).join(' + ')||'$0.00';
}

function renderEditList(){
  // The editor now opens directly for the game selected from the library.
}

function updateRatingPreview(){
  const raw=editRating.value.trim();
  if(raw===''){
    ratingPreviewDot.style.background='#6b7280';
    ratingPreviewText.textContent='No rating';
    return;
  }
  const n=Math.max(0,Math.min(10,Number(raw)));
  ratingPreviewDot.style.background=ratingColor(n);
  ratingPreviewText.textContent=`${n}/10`;
}

function selectEditGame(g){
  editingAppid=g.appid;
  editGameName.textContent=g.name;
  editGameImage.src=g.image||'';
  const tags=allGameTags(g);
  steamTags.innerHTML=tags.length?tags.map(c=>`<span class="steam-tag">${c}</span>`).join(''):'<span class="steam-tag">No Steam categories returned</span>';
  if(g.source==='local'){
    steamDescription.style.display='none';
  }else{
    steamDescription.style.display='';
    steamDescription.textContent=g.description||'Steam did not provide a short description for this game.';
  }
  editRating.value=g.rating??'';
  editPlaytime.textContent=fmtPlaytime(g.playtime_seconds);
  editFavorite.checked=!!g.favorite;
  editPros.value=g.pros||'';
  editCons.value=g.cons||'';
  editNotes.value=g.notes||'';
  renderEditCategoryChecks(g);
  renderLocalAttachBox(g);
  const editUninstallBtn=$('editUninstall');
  if(editUninstallBtn)editUninstallBtn.style.display=(g.source==='local'||g.installed)?'inline-block':'none';
  editorEmpty.style.display='none';
  editForm.classList.add('visible');
  updateRatingPreview();
}

function renderLocalAttachBox(g){
  if(!localAttachBox)return;
  if(g.source!=='local'){
    localAttachBox.style.display='none';
    return;
  }
  localAttachBox.style.display='block';
  localAttachSearch.value='';
  localAttachResults.innerHTML='';
  localAttachResults.style.display='none';
  if(g.steam_appid){
    localAttachedNote.style.display='block';
    localAttachedNote.innerHTML=`Linked to Steam: <b>${g.name}</b> <button class="btn-cancel" id="localUnlinkBtn" style="margin-left:8px;padding:2px 10px;font-weight:600">Unlink</button>`;
    $('localUnlinkBtn').onclick=()=>detachLocalSteam(g.appid);
  }else{
    localAttachedNote.style.display='none';
    localAttachedNote.innerHTML='';
  }
}

async function attachLocalSteam(appid,steamAppid){
  const res=await fetch('/api/attach_local_steam',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid,steam_appid:steamAppid})});
  const data=await res.json();
  if(Array.isArray(data)){
    setGames(data);
    const updated=data.find(x=>x.appid===appid);
    if(updated)selectEditGame(updated);
  }
}

async function detachLocalSteam(appid){
  const res=await fetch('/api/detach_local_steam',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid})});
  const data=await res.json();
  if(Array.isArray(data)){
    setGames(data);
    const updated=data.find(x=>x.appid===appid);
    if(updated)selectEditGame(updated);
  }
}

localAttachSearch.addEventListener('input',()=>{
  clearTimeout(localAttachSearchTimer);
  const term=localAttachSearch.value.trim();
  if(!term){localAttachResults.style.display='none';if(localAttachAbortCtrl)localAttachAbortCtrl.abort();return}
  localAttachResults.innerHTML='<div class="result-row" style="cursor:default;opacity:.65">Searching…</div>';
  localAttachResults.style.display='block';
  localAttachSearchTimer=setTimeout(async()=>{
    if(localAttachAbortCtrl)localAttachAbortCtrl.abort();
    const ctrl=new AbortController();
    localAttachAbortCtrl=ctrl;
    let items;
    try{
      const res=await fetch('/api/search?term='+encodeURIComponent(term)+'&limit=20',{signal:ctrl.signal});
      items=await res.json();
    }catch(e){
      if(e.name==='AbortError')return;
      items=[];
    }
    localAttachResults.innerHTML='';
    items.forEach(it=>{
      const row=document.createElement('div');
      row.className='result-row';
      row.innerHTML=`<img src="${it.thumb||''}" alt=""><span>${it.name}</span>`;
      row.onclick=()=>{
        localAttachResults.style.display='none';
        localAttachSearch.value='';
        attachLocalSteam(editingAppid,it.appid);
      };
      localAttachResults.append(row);
    });
    localAttachResults.style.display=items.length?'block':'none';
  },250);
});

document.addEventListener('click',e=>{
  if(!e.target.closest('#localAttachBox'))localAttachResults.style.display='none';
});

function renderEditCategoryChecks(g){
  if(!editCategoryChecks)return;
  if(!customCategories.length){
    editCategoryChecks.innerHTML='<span class="category-desc">No custom categories yet. Create them from the Categories page.</span>';
    return;
  }
  editCategoryChecks.innerHTML=customCategories.map(c=>`<label class="category-check"><input type="checkbox" value="${escapeHtml(c)}" ${(g.custom_categories||[]).includes(c)?'checked':''}>${escapeHtml(c)}</label>`).join('');
}

function openEditPage(g){
  // Remember exactly where we're editing from (library grid or a
  // category's game list) and how far scrolled, so Save/Cancel/Back can
  // return to that same spot instead of always dumping you back at the
  // top of the library.
  editReturnTo=categoriesPage.classList.contains('active')?'categories':(localPage.classList.contains('active')?'local':'library');
  editReturnScrollY=window.scrollY;
  libraryPage.style.display='none';
  categoriesPage.classList.remove('active');
  localPage.classList.remove('active');
  editPage.classList.add('active');
  navLibrary.classList.add('active');
  navCategories.classList.remove('active');
  selectEditGame(g);
  window.scrollTo({top:0,behavior:'smooth'});
}

function closeEditPage(){
  editPage.classList.remove('active');
  editingAppid=null;
  editForm.classList.remove('visible');
  editorEmpty.style.display='block';
  if(editReturnTo==='categories'){
    categoriesPage.classList.add('active');
    libraryPage.style.display='none';
    localPage.classList.remove('active');
    navLibrary.classList.remove('active');
    navLocal.classList.remove('active');
    navCategories.classList.add('active');
  }else if(editReturnTo==='local'){
    localPage.classList.add('active');
    libraryPage.style.display='none';
    categoriesPage.classList.remove('active');
    navLibrary.classList.remove('active');
    navLocal.classList.add('active');
    navCategories.classList.remove('active');
  }else{
    libraryPage.style.display='block';
    categoriesPage.classList.remove('active');
    localPage.classList.remove('active');
    navLibrary.classList.add('active');
    navLocal.classList.remove('active');
    navCategories.classList.remove('active');
  }
  window.scrollTo({top:editReturnScrollY,left:0});
}


function showLibrary(){
  libraryPage.style.display='block';
  categoriesPage.classList.remove('active');
  recommendationsPage.classList.remove('active');
  localPage.classList.remove('active');
  themesPage.classList.remove('active');
  settingsPage.classList.remove('active');
  mainMenuPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  navLibrary.classList.add('active');
  navLocal.classList.remove('active');
  navCategories.classList.remove('active');
  navRecommendations.classList.remove('active');
  navThemes.classList.remove('active');
  navSettings.classList.remove('active');
  navMainMenu.classList.remove('active');
  closeCategory();
  window.scrollTo({top:0,behavior:'smooth'});
}

function showLocal(){
  libraryPage.style.display='none';
  categoriesPage.classList.remove('active');
  recommendationsPage.classList.remove('active');
  themesPage.classList.remove('active');
  settingsPage.classList.remove('active');
  mainMenuPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  localPage.classList.add('active');
  navLibrary.classList.remove('active');
  navLocal.classList.add('active');
  navCategories.classList.remove('active');
  navRecommendations.classList.remove('active');
  navThemes.classList.remove('active');
  navSettings.classList.remove('active');
  navMainMenu.classList.remove('active');
  closeCategory();
  renderLocalGamesGrid();
  window.scrollTo({top:0,behavior:'smooth'});
  loadLocalFolders();
}

function showCategories(){
  libraryPage.style.display='none';
  recommendationsPage.classList.remove('active');
  localPage.classList.remove('active');
  themesPage.classList.remove('active');
  settingsPage.classList.remove('active');
  mainMenuPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  categoriesPage.classList.add('active');
  navLibrary.classList.remove('active');
  navLocal.classList.remove('active');
  navCategories.classList.add('active');
  navRecommendations.classList.remove('active');
  navThemes.classList.remove('active');
  navSettings.classList.remove('active');
  navMainMenu.classList.remove('active');
  closeCategory();
  window.scrollTo({top:0,behavior:'smooth'});
}

let recOffset=0, recHasMore=false, recLoading=false;
const REC_PAGE_SIZE=12;

function showRecommendations(){
  libraryPage.style.display='none';
  categoriesPage.classList.remove('active');
  localPage.classList.remove('active');
  themesPage.classList.remove('active');
  settingsPage.classList.remove('active');
  mainMenuPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  recommendationsPage.classList.add('active');
  navLibrary.classList.remove('active');
  navLocal.classList.remove('active');
  navCategories.classList.remove('active');
  navRecommendations.classList.add('active');
  navThemes.classList.remove('active');
  navSettings.classList.remove('active');
  navMainMenu.classList.remove('active');
  closeCategory();
  window.scrollTo({top:0,behavior:'smooth'});
  loadRecommendations();
}

function showThemes(){
  libraryPage.style.display='none';
  categoriesPage.classList.remove('active');
  localPage.classList.remove('active');
  recommendationsPage.classList.remove('active');
  settingsPage.classList.remove('active');
  mainMenuPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  themesPage.classList.add('active');
  navLibrary.classList.remove('active');
  navLocal.classList.remove('active');
  navCategories.classList.remove('active');
  navRecommendations.classList.remove('active');
  navThemes.classList.add('active');
  navSettings.classList.remove('active');
  navMainMenu.classList.remove('active');
  closeCategory();
  window.scrollTo({top:0,behavior:'smooth'});
}

function showSettings(){
  libraryPage.style.display='none';
  categoriesPage.classList.remove('active');
  localPage.classList.remove('active');
  recommendationsPage.classList.remove('active');
  themesPage.classList.remove('active');
  mainMenuPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  settingsPage.classList.add('active');
  navLibrary.classList.remove('active');
  navLocal.classList.remove('active');
  navCategories.classList.remove('active');
  navRecommendations.classList.remove('active');
  navThemes.classList.remove('active');
  navSettings.classList.add('active');
  navMainMenu.classList.remove('active');
  closeCategory();
  window.scrollTo({top:0,behavior:'smooth'});
  refreshUpdateInfo(false);
}

function showMainMenu(){
  libraryPage.style.display='none';
  categoriesPage.classList.remove('active');
  localPage.classList.remove('active');
  recommendationsPage.classList.remove('active');
  themesPage.classList.remove('active');
  settingsPage.classList.remove('active');
  editPage.classList.remove('active');
  editorEmpty.style.display='block';
  editForm.classList.remove('visible');
  mainMenuPage.classList.add('active');
  navLibrary.classList.remove('active');
  navLocal.classList.remove('active');
  navCategories.classList.remove('active');
  navRecommendations.classList.remove('active');
  navThemes.classList.remove('active');
  navSettings.classList.remove('active');
  navMainMenu.classList.add('active');
  closeCategory();
  window.scrollTo({top:0,behavior:'smooth'});
}

// ---------- Themes ----------
// Just an accent-color swap for now (no separate light/dark palettes) —
// every rule in the stylesheet above reads its accent from var(--teal) or
// rgba(var(--accent-rgb),alpha), so picking a theme only ever needs to set
// --accent-rgb on the root element; nothing else in the page has to change.
const THEME_COLORS=[
  {n:'Ocean Teal',c:'#3f8f8a'},{n:'Sky Blue',c:'#3ea8d8'},{n:'Azure',c:'#2f8fe0'},
  {n:'Cobalt',c:'#3457d5'},{n:'Sapphire',c:'#2f5fc4'},{n:'Steel Blue',c:'#4a7fa8'},
  {n:'Denim',c:'#3f6fa8'},{n:'Turquoise',c:'#2fb8c4'},{n:'Cyan',c:'#2fb8d8'},
  {n:'Seafoam',c:'#4fd6c4'},{n:'Indigo',c:'#5b5fc7'},{n:'Royal Purple',c:'#7c5fd6'},
  {n:'Violet',c:'#9457d6'},{n:'Orchid',c:'#b558c9'},{n:'Lavender',c:'#9f8fd6'},
  {n:'Periwinkle',c:'#7f8fd6'},{n:'Plum',c:'#8a4f8a'},{n:'Magenta',c:'#c1499e'},
  {n:'Fuchsia',c:'#d1489f'},{n:'Rose',c:'#d65a86'},{n:'Crimson',c:'#c1384f'},
  {n:'Red',c:'#c1493f'},{n:'Brick Red',c:'#a8493f'},{n:'Wine',c:'#8a3f5f'},
  {n:'Coral',c:'#e0704f'},{n:'Coral Pink',c:'#e37b8a'},{n:'Salmon',c:'#e08a6f'},
  {n:'Terracotta',c:'#c1704f'},{n:'Sienna',c:'#a8663f'},{n:'Rust',c:'#b0522f'},
  {n:'Orange',c:'#db8534'},{n:'Burnt Orange',c:'#d4763a'},{n:'Amber',c:'#d99b2b'},
  {n:'Gold',c:'#d8a84e'},{n:'Mustard',c:'#c9a227'},{n:'Copper',c:'#b3703f'},
  {n:'Bronze',c:'#9f7a3f'},{n:'Sand',c:'#c9a878'},{n:'Honey',c:'#e0a83f'},
  {n:'Tangerine',c:'#e58a3f'},{n:'Green',c:'#4a9f5c'},{n:'Forest Green',c:'#3f8f5e'},
  {n:'Emerald',c:'#2fa876'},{n:'Jade',c:'#2fa89a'},{n:'Lime',c:'#93c23e'},
  {n:'Olive',c:'#8a8f3f'},{n:'Charcoal Blue',c:'#45566a'},{n:'Graphite',c:'#5f6570'},
  {n:'Silver',c:'#8a92a0'},{n:'Taupe',c:'#a89478'},
];
const DEFAULT_THEME='#3f8f8a';
const THEME_STORAGE_KEY='gameVaultThemeAccent';

function hexToRgbStr(hex){
  const m=/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  if(!m)return '63,143,138';
  return `${parseInt(m[1],16)},${parseInt(m[2],16)},${parseInt(m[3],16)}`;
}

function applyTheme(hex,save=true){
  document.documentElement.style.setProperty('--accent-rgb',hexToRgbStr(hex));
  if(save){
    try{localStorage.setItem(THEME_STORAGE_KEY,hex);}catch(e){}
  }
  if(themeGrid){
    themeGrid.querySelectorAll('.theme-swatch').forEach(el=>{
      el.classList.toggle('active',el.dataset.hex.toLowerCase()===hex.toLowerCase());
    });
  }
}

function renderThemeGrid(){
  const current=(()=>{try{return localStorage.getItem(THEME_STORAGE_KEY)||DEFAULT_THEME;}catch(e){return DEFAULT_THEME;}})();
  themeGrid.innerHTML=THEME_COLORS.map(t=>`
    <div class="theme-swatch-cell">
      <button type="button" class="theme-swatch${t.c.toLowerCase()===current.toLowerCase()?' active':''}" style="background:${t.c}" data-hex="${t.c}" title="${escapeHtml(t.n)}" aria-label="${escapeHtml(t.n)}"></button>
      <div class="theme-swatch-label">${escapeHtml(t.n)}</div>
    </div>`).join('');
  themeGrid.querySelectorAll('.theme-swatch').forEach(el=>{
    el.onclick=()=>applyTheme(el.dataset.hex);
  });
}

function initTheme(){
  let saved=DEFAULT_THEME;
  try{saved=localStorage.getItem(THEME_STORAGE_KEY)||DEFAULT_THEME;}catch(e){}
  applyTheme(saved,false);
  renderThemeGrid();
}

// ---------- App icon & theme-from-picture ----------
// Both features need "let the user pick an arbitrary picture off their
// disk": inside the desktop window that's the native file dialog exposed
// via window.pywebview.api.pick_image_file (added to DesktopApi below); in
// the plain-browser-tab fallback mode there's no native dialog to hand off
// to, so a hidden <input type=file> + FileReader stands in for it instead.
function fileToDataURL(file){
  return new Promise((resolve,reject)=>{
    const r=new FileReader();
    r.onload=()=>resolve(r.result);
    r.onerror=()=>reject(new Error('read failed'));
    r.readAsDataURL(file);
  });
}

async function pickPictureSource(fileInput){
  if(window.pywebview&&window.pywebview.api&&window.pywebview.api.pick_image_file){
    const path=await window.pywebview.api.pick_image_file();
    return path?{path}:null;
  }
  return new Promise(resolve=>{
    fileInput.onchange=async()=>{
      const file=fileInput.files&&fileInput.files[0];
      fileInput.value='';
      if(!file){resolve(null);return;}
      try{resolve({image_data:await fileToDataURL(file)});}
      catch(e){resolve(null);}
    };
    fileInput.click();
  });
}

function applyIconDataUri(uri){
  const link=document.querySelector('link[rel="icon"]');
  if(link)link.href=uri;
  if(iconPreview)iconPreview.src=uri;
}

async function loadCurrentIconPreview(){
  // The favicon <link> already carries the current icon (baked in
  // server-side on page load), so just reuse it -- no extra request needed.
  const link=document.querySelector('link[rel="icon"]');
  if(link&&iconPreview)iconPreview.src=link.href;
}

if(iconChooseBtn)iconChooseBtn.onclick=async()=>{
  const source=await pickPictureSource(iconFileInput);
  if(!source)return;
  iconStatus.textContent='Updating icon...';
  try{
    const res=await fetch('/api/set_app_icon',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(source)});
    const data=await res.json();
    if(!res.ok){
      const messages={not_found:"That picture couldn't be found.",bad_image:"Couldn't read that as a picture.",missing_image:'Pick a picture first.'};
      iconStatus.textContent=messages[data.error]||'Could not set that as the icon.';
      return;
    }
    applyIconDataUri(data.data_uri);
    iconStatus.textContent='Icon updated. Restart the app for the taskbar icon to fully update.';
  }catch(e){
    iconStatus.textContent='Could not set that as the icon.';
  }
};

if(iconResetBtn)iconResetBtn.onclick=async()=>{
  iconStatus.textContent='Resetting...';
  try{
    const res=await fetch('/api/reset_app_icon',{method:'POST'});
    const data=await res.json();
    if(res.ok){
      applyIconDataUri(data.data_uri);
      iconStatus.textContent='Back to the original icon.';
    }else{
      iconStatus.textContent='Could not reset the icon.';
    }
  }catch(e){
    iconStatus.textContent='Could not reset the icon.';
  }
};

if(themeImageBtn)themeImageBtn.onclick=async()=>{
  const source=await pickPictureSource(themeImageInput);
  if(!source)return;
  themeImageStatus.textContent='Picking a color from the picture...';
  try{
    const res=await fetch('/api/theme_from_image',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(source)});
    const data=await res.json();
    if(!res.ok){
      const messages={
        pillow_missing:"This needs Pillow installed (pip install pillow).",
        not_found:"That picture couldn't be found.",
        bad_image:"Couldn't read that as a picture.",
        missing_image:'Pick a picture first.'
      };
      themeImageStatus.textContent=messages[data.error]||'Could not set a theme from that picture.';
      return;
    }
    applyTheme(data.hex);
    themeImageStatus.textContent=`Theme set from picture (${data.hex}).`;
  }catch(e){
    themeImageStatus.textContent='Could not set a theme from that picture.';
  }
};

// ---------- App updates ----------
// Mirrors the app-icon/theme-from-picture pattern above: a button kicks
// off a fetch to the backend routes added in the previous step, and the
// result (or a plain-language error) lands in a status line next to it.
let _latestUpdateInfo=null;
let _updatePollTimer=null;

function formatUpdateCheckError(data){
  const messages={
    no_releases:'No releases published yet.',
    network:'Could not reach GitHub. Check your connection.',
  };
  return messages[data.error]||'Could not check for updates.';
}

async function refreshUpdateInfo(userInitiated){
  if(userInitiated){
    checkUpdateBtn.disabled=true;
    checkUpdateBtn.textContent='Checking...';
    updateStatus.textContent='';
    installUpdateBtn.style.display='none';
    updateNotes.style.display='none';
  }
  try{
    const res=await fetch('/api/app_update/check');
    const data=await res.json();
    _latestUpdateInfo=data;
    if(updateCurrentVersion)updateCurrentVersion.textContent=data.current_version||'—';
    if(updateCurrentVersion&&data.version_source)updateCurrentVersion.title='Read from: '+data.version_source;
    if(!data.ok){
      updateStatus.textContent=formatUpdateCheckError(data);
      return;
    }
    if(data.update_available){
      updateStatus.textContent=`Update available: v${data.latest_version}`;
      installUpdateBtn.style.display='';
      installUpdateBtn.disabled=false;
      installUpdateBtn.textContent=`Update to v${data.latest_version}`;
      if(data.notes){
        updateNotes.textContent=data.notes;
        updateNotes.style.display='block';
      }else{
        updateNotes.style.display='none';
      }
    }else{
      if(data.skipped_newer&&data.skipped_newer.length){
        const sk=data.skipped_newer[0];
        updateStatus.textContent=`Found ${sk.tag} on GitHub, but it can't be offered: it ${sk.reason}.`;
      }else{
        updateStatus.textContent="You're on the latest version.";
      }
      installUpdateBtn.style.display='none';
      updateNotes.style.display='none';
    }
  }catch(e){
    updateStatus.textContent='Could not check for updates.';
  }finally{
    if(userInitiated){
      checkUpdateBtn.disabled=false;
      checkUpdateBtn.textContent='Check for updates';
    }
  }
}

if(checkUpdateBtn)checkUpdateBtn.onclick=()=>refreshUpdateInfo(true);

function setUpdateProgress(stage,percent){
  updateProgressWrap.style.display='block';
  updateProgressFill.style.width=(percent||0)+'%';
  const labels={
    downloading:`Downloading update... ${percent||0}%`,
    launching:'Starting installer...',
    done:'Installer launched — a Windows permission prompt may appear, then follow the setup steps. GameVault will close during install and reopen automatically.',
  };
  updateProgressLabel.textContent=labels[stage]||'';
}

function formatInstallError(err){
  const messages={
    elevation_denied:"Windows blocked the installer from getting admin rights. Try again and approve the permission prompt.",
    elevation_cancelled:"Update cancelled: the admin permission prompt was dismissed.",
    installer_missing:"The downloaded installer couldn't be found. Try again.",
    not_enough_memory:"Not enough resources to start the installer. Close some programs and try again.",
    sharing_violation:"The installer file is in use. Try again.",
  };
  return messages[err]||`Update failed: ${err||'unknown error'}`;
}

async function pollInstallStatus(){
  try{
    const res=await fetch('/api/app_update/install_status');
    const data=await res.json();
    if(data.stage==='error'){
      updateProgressLabel.textContent=formatInstallError(data.error);
      installUpdateBtn.disabled=false;
      return;
    }
    setUpdateProgress(data.stage,data.percent);
    if(data.stage==='downloading'||data.stage==='launching'){
      _updatePollTimer=setTimeout(pollInstallStatus,600);
    }else if(data.stage==='done'){
      installUpdateBtn.disabled=false;
    }
  }catch(e){
    installUpdateBtn.disabled=false;
  }
}

if(installUpdateBtn)installUpdateBtn.onclick=async()=>{
  if(!_latestUpdateInfo||!_latestUpdateInfo.download_url)return;
  clearTimeout(_updatePollTimer);
  installUpdateBtn.disabled=true;
  updateStatus.textContent='';
  try{
    const res=await fetch('/api/app_update/install',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({download_url:_latestUpdateInfo.download_url,asset_name:_latestUpdateInfo.asset_name})});
    const data=await res.json();
    if(!res.ok){
      updateProgressWrap.style.display='block';
      updateProgressLabel.textContent=data.error==='already_running'?'An update is already in progress.':'Could not start the update.';
      installUpdateBtn.disabled=false;
      return;
    }
    setUpdateProgress('downloading',0);
    pollInstallStatus();
  }catch(e){
    installUpdateBtn.disabled=false;
    updateProgressWrap.style.display='block';
    updateProgressLabel.textContent='Could not start the update.';
  }
};

function renderRecCard(g){
  const card=document.createElement('div');
  card.className='card';
  card.innerHTML=`
    <div class="card-media">
      <img src="${g.image||''}" alt="">
    </div>
    <div class="info" style="grid-template-columns:1fr auto">
      <div class="name">${escapeHtml(g.name)}</div>
      <div class="price">${fmt(g.price,g.currency)}</div>
      <button class="play-btn" style="grid-column:1;width:auto;padding:0 10px" data-add-appid="${g.appid}">+ Add to library</button>
      <a class="edit-btn" style="grid-column:2;display:flex;align-items:center;justify-content:center;text-decoration:none;font-size:.7rem;font-weight:700" href="https://store.steampowered.com/app/${g.appid}" target="_blank" rel="noopener" title="View on Steam">↗</a>
    </div>`;
  card.querySelector('[data-add-appid]').onclick=()=>addGame(g.appid);
  recommendationsGrid.append(card);
}

async function loadRecommendations(refresh){
  recommendationsGrid.innerHTML='';
  recEmptyState.style.display='none';
  recBasedOn.textContent='Finding matches on Steam…';
  recOffset=0;
  recHasMore=false;
  recLoading=false;
  await loadMoreRecommendations(refresh);
}

async function loadMoreRecommendations(refresh){
  if(recLoading) return;
  recLoading=true;
  try{
    const params=new URLSearchParams({offset:recOffset,limit:REC_PAGE_SIZE});
    if(refresh) params.set('refresh','1');
    const res=await fetch('/api/recommendations?'+params.toString());
    const data=await res.json();
    const recs=data.recommendations||[];
    recBasedOn.textContent=(data.based_on&&data.based_on.length)?('Based on: '+data.based_on.join(', ')):'';
    if(!recs.length && recOffset===0){
      recEmptyState.textContent='No recommendations yet — add a few rated games to your library first.';
      recEmptyState.style.display='block';
      recHasMore=false;
      return;
    }
    recs.forEach(renderRecCard);
    recOffset+=recs.length;
    recHasMore=!!data.has_more;
  }catch(e){
    if(recOffset===0){
      recBasedOn.textContent='';
      recEmptyState.textContent='Could not reach Steam for recommendations right now.';
      recEmptyState.style.display='block';
    }
    recHasMore=false;
  }finally{
    recLoading=false;
  }
  // With few results, the grid may not fill (let alone overflow) the
  // viewport, so the browser never fires a 'scroll' event and the listener
  // below never gets a chance to ask for more. Keep topping up right here
  // until either the page actually has something to scroll or we run out
  // of recommendations, so "scroll for more" works from the very first load.
  if(recommendationsPage.classList.contains('active') && recHasMore && !recLoading){
    const scrollable=document.documentElement.scrollHeight>window.innerHeight+40;
    if(!scrollable) await loadMoreRecommendations(false);
  }
}

window.addEventListener('scroll',()=>{
  if(!recommendationsPage.classList.contains('active')) return;
  if(!recHasMore||recLoading) return;
  const nearBottom=window.innerHeight+window.scrollY>=document.documentElement.scrollHeight-400;
  if(nearBottom) loadMoreRecommendations(false);
});

let _lastGamesSignature=null;
function setGames(games){
  // The 20s background poll (loadGames below) calls this every tick even
  // when nothing actually changed. Rebuilding the whole grid from scratch
  // (grid.innerHTML='' + re-creating every card) is wasted work — and
  // visible jank on a big library — when the data is identical to what's
  // already on screen, so skip the rebuild in that case. Any real change
  // (a new game, an edited rating, playtime ticking up on a running game,
  // etc.) still renders immediately as before.
  const sig=JSON.stringify(games);
  allGames=games;
  if(sig===_lastGamesSignature)return;
  _lastGamesSignature=sig;
  buildCategoryBoxes();
  render();
  renderCategoryDetail();
  renderLocalGamesGrid();
}

async function loadGames(){
  const res=await fetch('/api/games');
  setGames(await res.json());
}

async function removeGame(appid){
  const res=await fetch('/api/remove',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid})});
  setGames(await res.json());
  if(editingAppid===appid)closeEditPage();
}

async function uninstallGame(appid,name,isLocal){
  const msg=isLocal
    ?`Uninstall ${name}? If it has no separate uninstaller, GameVault will delete its game folder plus any leftover data it finds in AppData/Program Files.`
    :`Uninstall ${name} through Steam?`;
  if(!confirm(msg))return;
  const btn=$('editUninstall');
  const old=btn.textContent;
  btn.disabled=true;btn.textContent='Uninstalling…';
  try{
    const res=await fetch('/api/uninstall',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid})});
    const data=await res.json();
    if(!res.ok){
      const msgs={
        not_installed:`${name} isn't currently installed.`,
        exe_not_found:`${name}'s saved .exe path no longer exists.`,
        not_found:`${name} couldn't be found in your library.`,
      };
      alert(msgs[data.error]||('Could not uninstall '+name+': '+(data.detail||data.error||'unknown error')));
    }else if(data.method==='manual'){
      const n=(data.removed_paths||[]).length;
      alert(`${name} has been removed`+(n?` — cleaned up ${n} folder${n===1?'':'s'} (its game files plus any leftover data found in AppData/Program Files).`:'.'));
      await removeGame(appid);
    }else if(isLocal){
      alert(`${name}'s uninstaller has been launched.`);
    }
    // Steam handoff needs no popup here -- Steam shows its own dialog.
  }catch(e){
    alert('Could not reach GameVault to uninstall '+name+'.');
  }finally{
    btn.disabled=false;btn.textContent=old;
  }
}

async function detectSteamGames(silent){
  // Runs quietly on startup and every 3 minutes in the background -- there's
  // no button for this anymore, so `silent` is effectively always true, but
  // it's kept so this can still be called with feedback if ever needed.
  try{
    const res=await fetch('/api/detect_steam_games',{method:'POST'});
    const data=await res.json();
    if(Array.isArray(data.games))setGames(data.games);
    if(!silent&&data.added)alert(data.added.length?`Found ${data.added.length} newly installed Steam game(s).`:'No new installed Steam games found.');
  }catch(e){
    if(!silent)alert('Could not scan your Steam library.');
  }
}

async function addGame(appid){
  resultsBox.style.display='none';
  searchInput.value='';
  pendingAdd=appid;
  dlcTitle.textContent='Choose DLCs';
  dlcList.innerHTML='<div class="empty" style="margin:0">Loading DLCs quickly…<br><small>Fetching prices in a batch and caching the results.</small></div>';
  dlcModal.classList.add('open');
  try{
    const res=await fetch('/api/dlc-options/'+appid);
    const data=await res.json(); pendingBase=data.base||null; pendingDlc=data.dlcs||[]; pendingBundles=data.bundles||[];
    dlcTitle.textContent=(pendingBase?.name||'Choose DLCs')+' — DLCs';
    renderDlcList();
  }catch(e){
    pendingDlc=[];
    renderDlcList();
  }
}
function renderDlcList(){
  dlcList.innerHTML='';
  if(!pendingDlc.length){
    dlcList.innerHTML='<div class="empty" style="margin:0">This game has no DLC listed by Steam.</div>';
    updateDlcTotal(); return;
  }
  pendingDlc.forEach((d,i)=>{
    const row=document.createElement('label'); row.className='dlc-row';
    row.innerHTML=`<input type="checkbox" class="dlc-check" data-index="${i}" checked>
      <img src="${d.image||''}" alt=""><span class="dlc-row-name">${escapeHtml(d.name)}</span>
      <span class="dlc-row-price">${fmt(d.price,d.currency)}</span>`;
    dlcList.append(row);
  });
  dlcList.querySelectorAll('.dlc-check').forEach(x=>x.onchange=updateDlcTotal);
  updateDlcTotal();
}
function bestDlcPlan(selected){
  const chosen=new Map(selected.map(d=>[Number(d.appid),d])), ids=[...chosen.keys()];
  if(!ids.length)return {cost:0,bundles:[],individual:[]};
  const bundles=(pendingBundles||[]).map(b=>({...b,covered:new Set((b.covered_dlc||[]).map(Number).filter(x=>chosen.has(x)))})).filter(b=>b.covered.size);
  const n=ids.length, idx=new Map(ids.map((x,i)=>[x,i]));
  if(n>20){let rem=new Set(ids),bs=[];while(rem.size){let best=null,sv=0;for(const b of bundles){const hit=[...b.covered].filter(x=>rem.has(x));const save=hit.reduce((s,x)=>s+(Number(chosen.get(x).price)||0),0)-Number(b.price||0);if(hit.length&&save>sv){sv=save;best={b,hit}}}if(!best)break;bs.push(best.b);best.hit.forEach(x=>rem.delete(x));}const ind=[...rem],cost=bs.reduce((s,b)=>s+Number(b.price||0),0)+ind.reduce((s,x)=>s+(Number(chosen.get(x).price)||0),0);return {cost,bundles:bs,individual:ind};}
  const full=(1<<n)-1, dp=Array(full+1).fill(null);dp[0]={cost:0,bundles:[],individual:[]};
  for(let mask=0;mask<=full;mask++){const cur=dp[mask];if(!cur)continue;ids.forEach((id,i)=>{if(mask&(1<<i))return;const nm=mask|(1<<i),cost=cur.cost+(Number(chosen.get(id).price)||0);if(!dp[nm]||cost<dp[nm].cost)dp[nm]={cost,bundles:cur.bundles.slice(),individual:cur.individual.concat(id)}});bundles.forEach(b=>{let bm=0;b.covered.forEach(id=>bm|=1<<idx.get(id));const nm=mask|bm;if(nm===mask)return;const cost=cur.cost+Number(b.price||0);if(!dp[nm]||cost<dp[nm].cost)dp[nm]={cost,bundles:cur.bundles.concat(b),individual:cur.individual.slice()}})}
  return dp[full];
}
function updateDlcTotal(){
  const selected=[...dlcList.querySelectorAll('.dlc-check:checked')].map(x=>pendingDlc[Number(x.dataset.index)]).filter(Boolean);
  const base=(pendingBase?.price==null)?0:Number(pendingBase.price)||0, plan=bestDlcPlan(selected), cur=pendingBase?.currency||selected.find(d=>d.currency)?.currency||'';
  dlcTotal.textContent=fmt(base+plan.cost,cur);
  dlcTotal.title=plan.bundles.length?plan.bundles.map(b=>b.name+' — '+fmt(b.price,b.currency)).join(' + '):'Base game + selected DLCs';
}
async function confirmAddGame(){
  if(pendingAdd===null)return;
  const selected=[...dlcList.querySelectorAll('.dlc-check:checked')].map(x=>pendingDlc[Number(x.dataset.index)]).filter(Boolean), plan=bestDlcPlan(selected);
  const res=await fetch('/api/add',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid:pendingAdd,dlcs:selected,bundle_plan:plan})});
  if(res.ok){setGames(await res.json());closeDlcModal();}else closeDlcModal();
}
function closeDlcModal(){dlcModal.classList.remove('open');pendingAdd=null;pendingDlc=[];pendingBundles=[];pendingBase=null;}


async function saveEdit(){
  if(editingAppid===null)return;
  const raw=editRating.value.trim();
  const body={
    appid:editingAppid,
    rating:raw===''?null:Number(raw),
    favorite:editFavorite.checked,
    pros:editPros.value,
    cons:editCons.value,
    notes:editNotes.value,
    custom_categories:[...editCategoryChecks.querySelectorAll('input:checked')].map(x=>x.value)
  };
  const res=await fetch('/api/update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  if(res.ok){
    setGames(await res.json());
    closeEditPage();
  }
}

async function toggleFavorite(appid,current){
  const res=await fetch('/api/update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid,favorite:!current})});
  if(res.ok)setGames(await res.json());
}

function findPlayBtn(appid){
  return document.querySelector(`.play-btn[data-appid="${appid}"]`);
}

const PLAY_BTN_LABEL={launching:'Launching…',stopping:'Stopping…',playing:'■ Stop',idle:'▶ Play'};

// Single source of truth for what a play button should look like, used
// both when a card is first built (renderGameCards) and whenever a
// button's state changes afterwards (setPlayBtnState). Having one place
// that considers pendingAction (a real launch/stop in flight) is what
// stops a periodic re-render from wiping a "Launching…"/"Stopping…"
// button back to a plain, clickable "Play" mid-action.
function playBtnState(appid){
  return pendingAction.get(appid)||(playingAppids.has(appid)?'playing':'idle');
}

function playBtnMarkup(appid,name){
  const state=playBtnState(appid);
  const label=PLAY_BTN_LABEL[state];
  const disabled=(state==='launching'||state==='stopping');
  const isPlaying=(state==='playing'||state==='stopping');
  const isLocal=allGames.find(g=>g.appid===appid)?.source==='local';
  const title=state==='playing'?`Stop ${name}`:state==='idle'?(isLocal?`Launch ${name}`:`Launch ${name} in Steam`):label;
  return `<button class="play-btn ${isPlaying?'is-playing':''}" data-appid="${appid}" ${disabled?'disabled':''} title="${title}" aria-label="${title}">${label}</button>`;
}

function setPlayBtnState(appid,name,state){
  if(state==='launching'||state==='stopping')pendingAction.set(appid,state);
  else pendingAction.delete(appid);
  const btn=findPlayBtn(appid);
  if(!btn)return;
  const label=PLAY_BTN_LABEL[state];
  btn.disabled=(state==='launching'||state==='stopping');
  btn.textContent=label;
  btn.classList.toggle('is-playing',state==='playing'||state==='stopping');
  const title=state==='playing'?`Stop ${name}`:state==='idle'?`Launch ${name} in Steam`:label;
  btn.title=title;btn.setAttribute('aria-label',title);
}

function handlePlayClick(appid,name){
  if(playingAppids.has(appid))stopGame(appid,name);
  else launchGame(appid,name);
}

async function launchGame(appid,name){
  setPlayBtnState(appid,name,'launching');
  try{
    const res=await fetch('/api/launch',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid})});
    if(!res.ok){
      // Fall back to letting the OS/browser itself handle the steam:// link
      window.location.href='steam://run/'+appid;
    }
  }catch(e){
    window.location.href='steam://run/'+appid;
  }
  // The real "is it actually running" answer comes from syncRunningStatus
  // below, polling the server's process check — that's what flips this to
  // "Stop" for real. We deliberately don't guess "playing" on a fixed
  // timer here anymore: that guess used to race the real check (which can
  // take longer, e.g. slow game launchers) and lose, flipping the button
  // back to "Play" a few seconds later and making it look broken.
  // If the server can't detect real process state at all (no psutil
  // installed), there's no real signal to wait for, so fall back to an
  // optimistic guess after a few seconds — better than a button stuck on
  // "Launching…" forever.
  setTimeout(()=>{
    if(pendingAction.get(appid)!=='launching')return;  // already resolved for real, or superseded
    if(psutilSupported===false){
      playingAppids.add(appid);
      setPlayBtnState(appid,name,'playing');
    }
  },3000);
  // Safety net: if nothing ever confirms it (game failed to start, the
  // install folder couldn't be matched, etc.), don't leave the button
  // disabled on "Launching…" indefinitely — hand it back.
  setTimeout(()=>{
    if(pendingAction.get(appid)==='launching')setPlayBtnState(appid,name,'idle');
  },20000);
}

async function stopGame(appid,name){
  setPlayBtnState(appid,name,'stopping');
  try{
    const res=await fetch('/api/stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({appid})});
    const data=await res.json().catch(()=>({}));
    if(!res.ok||data.error){
      alert("Couldn't stop the game automatically ("+(data.detail||data.error||'unknown error')+"). You can still stop it from the Steam client.");
    }
  }catch(e){
    alert("Couldn't reach the app to stop the game.");
  }
  playingAppids.delete(appid);
  setPlayBtnState(appid,name,'idle');
}

// Continuously asks the server which of the library's games are actually
// running right now — powered by a background tracker on the server that
// scans for running processes on its own, whether the game was started
// from this app's Play button, straight from the Steam client, or any
// other way. This is what makes the Stop button show up even when you
// didn't hit Play here, and it's also what falls the button back to
// "Play" on its own if you quit the game from inside Steam instead of
// using the Stop button.
async function syncRunningStatus(){
  try{
    const res=await fetch('/api/running_status');
    const data=await res.json();
    psutilSupported=data.supported;
    if(!data.supported)return;  // psutil isn't installed server-side
    const runningIds=new Set(Object.keys(data.running||{}).map(Number));
    allGames.forEach(g=>{
      if(pendingAction.has(g.appid))return;  // mid "Launching…"/"Stopping…" — don't fight it
      const isRunning=runningIds.has(g.appid);
      const wasTracked=playingAppids.has(g.appid);
      if(isRunning&&!wasTracked){
        playingAppids.add(g.appid);
        setPlayBtnState(g.appid,g.name,'playing');
      }else if(!isRunning&&wasTracked){
        playingAppids.delete(g.appid);
        setPlayBtnState(g.appid,g.name,'idle');
      }
    });
  }catch(e){/* transient network hiccup — just try again next tick */}
}
setInterval(syncRunningStatus,6000);

let searchAbortCtrl=null;
searchInput.addEventListener('input',()=>{
  clearTimeout(searchTimer);
  const term=searchInput.value.trim();
  if(!term){resultsBox.style.display='none';if(searchAbortCtrl)searchAbortCtrl.abort();return}
  // Show a lightweight "Searching…" row immediately so typing feels
  // responsive even while the request to Steam is still in flight.
  resultsBox.innerHTML='<div class="result-row" style="cursor:default;opacity:.65">Searching…</div>';
  resultsBox.style.display='block';
  searchTimer=setTimeout(async()=>{
    if(searchAbortCtrl)searchAbortCtrl.abort();  // drop any still-pending older search
    const ctrl=new AbortController();
    searchAbortCtrl=ctrl;
    let items;
    try{
      const res=await fetch('/api/search?term='+encodeURIComponent(term),{signal:ctrl.signal});
      items=await res.json();
    }catch(e){
      if(e.name==='AbortError')return;  // superseded by a newer keystroke, ignore
      items=[];
    }
    resultsBox.innerHTML='';
    items.forEach(it=>{
      const row=document.createElement('div');
      row.className='result-row';
      row.innerHTML=`<img src="${it.thumb||''}" alt=""><span>${it.name}</span>`;
      row.onclick=()=>addGame(it.appid);
      resultsBox.append(row);
    });
    resultsBox.style.display=items.length?'block':'none';
  },250);
});

document.addEventListener('click',e=>{
  if(!e.target.closest('.search-box'))resultsBox.style.display='none';
});

editRating.addEventListener('input',updateRatingPreview);
librarySearchInput.addEventListener('input',render);
librarySort.addEventListener('change',render);
if(installedOnlyCheckbox){
  try{installedOnlyCheckbox.checked=localStorage.getItem(INSTALLED_ONLY_KEY)==='1';}catch(e){}
  installedOnlyCheckbox.addEventListener('change',()=>{
    try{localStorage.setItem(INSTALLED_ONLY_KEY,installedOnlyCheckbox.checked?'1':'0');}catch(e){}
    render();
  });
}
categoryBoxSearch.addEventListener('input',buildCategoryBoxes);
categoryDetailSearch.addEventListener('input',renderCategoryDetail);
categorySort.addEventListener('change',renderCategoryDetail);
$('editBack').onclick=closeEditPage;
$('editCancel').onclick=closeEditPage;
$('editSave').onclick=saveEdit;
$('editDelete').onclick=async()=>{
  if(editingAppid===null)return;
  const name=editGameName.textContent;
  if(!confirm('Remove '+name+' from your library?'))return;
  await removeGame(editingAppid);
};
$('editUninstall').onclick=()=>{
  if(editingAppid===null)return;
  const g=allGames.find(x=>x.appid===editingAppid);
  if(!g)return;
  uninstallGame(g.appid,g.name,g.source==='local');
};

async function loadCustomCategories(){
  const res=await fetch('/api/categories');
  customCategories=await res.json();
  renderCustomCategoryList();
  buildCategoryBoxes();
  if(editingAppid!==null){
    const g=allGames.find(x=>x.appid===editingAppid);
    if(g)renderEditCategoryChecks(g);
  }
}
function renderCustomCategoryList(){
  const box=$('customCatList'); if(!box)return;
  box.innerHTML=customCategories.map(c=>`<div class="custom-cat-chip"><span>${escapeHtml(c)}</span><button title="Delete ${escapeHtml(c)}" data-cat="${escapeHtml(c)}">×</button></div>`).join('');
  box.querySelectorAll('button').forEach(b=>b.onclick=async()=>{
    const name=b.dataset.cat;
    if(!confirm('Delete custom category "'+name+'"? Games will stay in your library.'))return;
    const res=await fetch('/api/category/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
    const data=await res.json();
    customCategories=data.categories||[];
    allGames=data.games||allGames;
    renderCustomCategoryList(); buildCategoryBoxes(); render();
    if(editingAppid!==null){const g=allGames.find(x=>x.appid===editingAppid);if(g)renderEditCategoryChecks(g);}
  });
}
$('addCategoryBtn').onclick=async()=>{
  const input=$('newCategoryInput'), name=input.value.trim(); if(!name)return;
  const res=await fetch('/api/category/create',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
  if(res.ok){customCategories=await res.json();input.value='';renderCustomCategoryList();buildCategoryBoxes(); if(editingAppid!==null){const g=allGames.find(x=>x.appid===editingAppid);if(g)renderEditCategoryChecks(g);}}
};
$('newCategoryInput').addEventListener('keydown',e=>{if(e.key==='Enter')$('addCategoryBtn').click();});
$('dlcSelectAll').onclick=()=>{dlcList.querySelectorAll('.dlc-check').forEach(x=>x.checked=true);updateDlcTotal();};
$('dlcSelectNone').onclick=()=>{dlcList.querySelectorAll('.dlc-check').forEach(x=>x.checked=false);updateDlcTotal();};
$('dlcCancel').onclick=closeDlcModal;
$('dlcConfirm').onclick=confirmAddGame;
dlcModal.addEventListener('click',e=>{if(e.target===dlcModal)closeDlcModal();});

$('navLibrary').onclick=showLibrary;
$('navLocal').onclick=showLocal;
$('navCategories').onclick=showCategories;
$('navRecommendations').onclick=showRecommendations;
navMainMenu.onclick=showMainMenu;
$('mmLibrary').onclick=showLibrary;
$('mmLocal').onclick=showLocal;
$('mmCategories').onclick=showCategories;
$('mmRecommendations').onclick=showRecommendations;
$('mmThemes').onclick=showThemes;
$('mmSettings').onclick=showSettings;

let localFolders=[];

function renderLocalFolders(){
  if(!localFolders.length){
    localFoldersList.innerHTML='<span class="category-desc">No watched folders yet — add one above and every .exe in it (and its subfolders) is added automatically.</span>';
    return;
  }
  localFoldersList.innerHTML=localFolders.map(f=>`<div class="custom-cat-chip"><span title="${escapeHtml(f)}">${escapeHtml(f)}</span><button title="Stop watching ${escapeHtml(f)}" data-folder="${escapeHtml(f)}">×</button></div>`).join('');
  localFoldersList.querySelectorAll('button').forEach(b=>b.onclick=async()=>{
    const path=b.dataset.folder;
    if(!confirm('Stop watching "'+path+'"? Games found there will be removed from your library.'))return;
    const res=await fetch('/api/remove_local_folder',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path})});
    const data=await res.json();
    localFolders=data.folders||[];
    setGames(data.games||allGames);
    renderLocalFolders();
  });
}

async function loadLocalFolders(){
  const res=await fetch('/api/local_folders');
  localFolders=await res.json();
  renderLocalFolders();
}

localFolderBrowseBtn.onclick=async()=>{
  if(!(window.pywebview&&window.pywebview.api&&window.pywebview.api.pick_folder)){
    localAddError.style.display='block';
    localAddError.textContent="Native folder browser isn't available in this mode — paste the folder path directly into the field.";
    return;
  }
  const path=await window.pywebview.api.pick_folder();
  if(path)localFolderPath.value=path;
};

localFolderAddBtn.onclick=async()=>{
  const path=localFolderPath.value.trim();
  localAddError.style.display='none';
  if(!path){
    localAddError.style.display='block';
    localAddError.textContent='Pick or paste a folder to watch first.';
    return;
  }
  localFolderAddBtn.disabled=true;
  const old=localFolderAddBtn.textContent;
  localFolderAddBtn.textContent='Scanning...';
  try{
    const res=await fetch('/api/add_local_folder',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path})});
    const data=await res.json();
    if(!res.ok){
      const messages={not_found:"That folder doesn't exist.",already_added:"That folder is already being watched.",missing_path:'Pick or paste a folder to watch first.'};
      localAddError.style.display='block';
      localAddError.textContent=messages[data.error]||'Could not add that folder.';
      return;
    }
    localFolders=data.folders||[];
    setGames(data.games||allGames);
    renderLocalFolders();
    localFolderPath.value='';
  }finally{
    localFolderAddBtn.disabled=false;
    localFolderAddBtn.textContent=old;
  }
};

localScanBtn.onclick=async()=>{
  localScanBtn.disabled=true;
  const old=localScanBtn.textContent;
  localScanBtn.textContent='Scanning...';
  try{
    const res=await fetch('/api/scan_local_folders',{method:'POST'});
    const data=await res.json();
    setGames(data.games||allGames);
  }finally{
    localScanBtn.disabled=false;
    localScanBtn.textContent=old;
  }
};

localBrowseBtn.onclick=async()=>{
  if(!(window.pywebview&&window.pywebview.api&&window.pywebview.api.pick_exe_file)){
    localAddError.style.display='block';
    localAddError.textContent="Native file browser isn't available in this mode — paste the .exe path directly into the field.";
    return;
  }
  const path=await window.pywebview.api.pick_exe_file();
  if(path)localExePath.value=path;
};

localAddBtn.onclick=async()=>{
  const path=localExePath.value.trim();
  localAddError.style.display='none';
  if(!path){
    localAddError.style.display='block';
    localAddError.textContent="Pick or paste a path to the game's .exe first.";
    return;
  }
  localAddBtn.disabled=true;
  try{
    const res=await fetch('/api/add_local',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path})});
    const data=await res.json();
    if(!res.ok){
      const messages={
        not_found:"That .exe path doesn't exist.",
        not_exe:'That file needs to end in .exe.',
        already_added:'That game is already in your local library.',
        missing_path:"Pick or paste a path to the game's .exe first."
      };
      localAddError.style.display='block';
      localAddError.textContent=messages[data.error]||'Could not add that game.';
      return;
    }
    setGames(data);
    localExePath.value='';
  }finally{
    localAddBtn.disabled=false;
  }
};
$('backCategories').onclick=closeCategory;
$('recRefreshBtn').onclick=()=>loadRecommendations(true);
navThemes.onclick=showThemes;
navSettings.onclick=showSettings;

$('navToggle').onclick=()=>{
  const hidden=$('appShell').classList.toggle('nav-hidden');
  const label=hidden?'Show navigation':'Hide navigation';
  $('navToggle').title=label;
  $('navToggle').ariaLabel=label;
};

// Fullscreen: works both inside the pywebview desktop window (via the
// exposed toggle_fullscreen bridge) and in a plain browser tab (via the
// standard Fullscreen API), whichever mode this happens to be running in.
let pywebviewReady=false;
window.addEventListener('pywebviewready',()=>{pywebviewReady=true;});
function toggleFullscreen(){
  if(window.pywebview&&pywebviewReady&&window.pywebview.api&&window.pywebview.api.toggle_fullscreen){
    window.pywebview.api.toggle_fullscreen();
  }else if(document.fullscreenElement){
    document.exitFullscreen();
  }else if(document.documentElement.requestFullscreen){
    document.documentElement.requestFullscreen().catch(()=>{});
  }
}
$('fullscreenToggle').onclick=toggleFullscreen;
document.addEventListener('keydown',e=>{
  if(e.key==='F11'){e.preventDefault();toggleFullscreen();}
});

$('refreshBtn').onclick=async e=>{
  const btn=e.currentTarget,old=btn.textContent;
  btn.disabled=true;btn.textContent='Refreshing Steam data...';
  try{
    const res=await fetch('/api/refresh',{method:'POST'});
    setGames(await res.json());
  }finally{
    btn.disabled=false;btn.textContent=old;
  }
};

initTheme();
loadCurrentIconPreview();
loadGames().then(loadCustomCategories).then(syncRunningStatus);
loadLocalFolders();
detectSteamGames(true);  // quiet scan on startup so newly installed Steam games show up automatically
// Re-pull games periodically so playtime numbers keep ticking up while a
// game runs, without a manual refresh. Skipped while editing a game so it
// doesn't disturb an in-progress edit.
setInterval(()=>{ if(editingAppid===null) loadGames(); },20000);
// Quietly re-scan the Steam library every 3 minutes so newly installed (or
// uninstalled) Steam games are picked up automatically -- no button needed.
setInterval(()=>{ if(editingAppid===null) detectSteamGames(true); },180000);
</script>
</body>
</html>
"""


# NOTE: __APP_ICON_DATA_URI__ is deliberately left unresolved here -- it's
# substituted per-request in index() above so a custom icon takes effect
# immediately, without needing to restart the app.


class DesktopApi:
    """Exposed to the page's JS as window.pywebview.api.* inside the
    desktop window, so the in-page fullscreen button / F11 key can toggle
    the actual OS window instead of just an in-page CSS trick."""

    def toggle_fullscreen(self):
        if not webview.windows:
            return
        win = webview.windows[0]
        win.toggle_fullscreen()
        # The underlying web-view control (WebView2 on Windows, WebKitGTK/Qt
        # on Linux/mac) doesn't always re-measure its own bounds the instant
        # the OS window snaps to fullscreen — for a moment the page is still
        # laid out for the old window size, which is what shows up as a
        # sliver of the desktop/old background along an edge. Nudging the
        # size by a pixel and back forces the control to recompute its
        # bounds against the *new* window size immediately instead of
        # waiting for the next real resize event.
        def _nudge():
            time.sleep(0.12)
            try:
                w, h = win.width, win.height
                win.resize(w - 1, h)
                win.resize(w, h)
            except Exception:
                pass  # older pywebview versions without .resize() — harmless no-op
        threading.Thread(target=_nudge, daemon=True).start()

    def pick_exe_file(self):
        """Native "open file" dialog for the Local Games "Browse..." button,
        restricted to .exe files. Returns the chosen path, or None if the
        user cancelled or this isn't running inside the desktop window
        (e.g. the plain-browser-tab fallback mode, which has no native
        dialog to hand off to)."""
        if not webview.windows:
            return None
        win = webview.windows[0]
        try:
            result = win.create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=("Executable files (*.exe)", "All files (*.*)"),
            )
        except Exception:
            return None
        return result[0] if result else None

    def pick_folder(self):
        """Native "select folder" dialog for watching a folder in Local
        Games. Returns the chosen path, or None if cancelled / unavailable
        (plain-browser-tab fallback mode has no native dialog to use)."""
        if not webview.windows:
            return None
        win = webview.windows[0]
        try:
            result = win.create_file_dialog(webview.FOLDER_DIALOG)
        except Exception:
            return None
        return result[0] if result else None

    def pick_image_file(self):
        """Native "open file" dialog for the Themes page's "App icon" and
        "Theme from a picture" pickers. Returns the chosen path, or None if
        cancelled / unavailable (plain-browser-tab fallback mode uses a
        hidden <input type=file> instead -- see pickPictureSource in the
        page's JS)."""
        if not webview.windows:
            return None
        win = webview.windows[0]
        try:
            result = win.create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=("Image files (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)", "All files (*.*)"),
            )
        except Exception:
            return None
        return result[0] if result else None


def _run_flask():
    # use_reloader must be off: the reloader spawns a second process, which
    # would open two desktop windows / bind the port twice.
    try:
        app.run(debug=False, port=5000, use_reloader=False, threaded=True)
    except OSError as e:
        # Most likely port 5000 is already taken (another copy of this app
        # still running in the background, or something else on that port).
        # A thread crashing silently here is exactly the "opens then
        # disappears with no error" symptom — flag it loudly instead.
        _report_fatal_error(
            "Couldn't start the local server on port 5000 — it's probably "
            "already in use by another running copy of this app (check Task "
            "Manager) or another program.\n\n" + repr(e)
        )


def _report_fatal_error(message):
    """Last-resort error reporting for a windowed/noconsole .exe: a plain
    print() here goes nowhere (there's no console to see it), and an
    unhandled exception just makes the process disappear with zero
    feedback — which is exactly the "opens for a second then vanishes in
    Task Manager" symptom. So this always writes the full detail to
    crash_log.txt next to the app, and additionally pops up a native
    Windows message box when possible so a double-click user actually
    sees *something* instead of nothing."""
    import traceback
    text = f"{time.strftime('%Y-%m-%d %H:%M:%S')}\n{message}\n{'-'*60}\n"
    try:
        with open(os.path.join(DATA_DIR, "crash_log.txt"), "a", encoding="utf-8") as f:
            f.write(text)
    except OSError:
        pass
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                "Game Vault couldn't start:\n\n" + message[:800] +
                f"\n\nFull details were written to:\ncrash_log.txt (next to the app)",
                "Game Vault — startup error", 0x10  # MB_ICONERROR
            )
        except Exception:
            pass


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        # Without this, Windows identifies the running process as
        # "python.exe" for taskbar purposes, not as this app. That's what
        # causes a pinned taskbar icon to revert to the plain Python icon
        # after launch — Windows creates a separate taskbar entry tied to
        # python.exe's own identity/icon instead of matching the pinned
        # shortcut. Setting an explicit AppUserModelID here gives this
        # process its own identity so the pin keeps its icon.
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GameVault.DesktopApp")
        except Exception:
            pass  # non-fatal — worst case, taskbar behavior falls back to default

    try:
        _start_playtime_tracker()
        if HAS_WEBVIEW:
            # Run Flask in a background thread and show it inside a native
            # desktop window (no browser chrome, no visible URL bar) via
            # pywebview. Same Flask app, same UI, just displayed differently.
            server_thread = threading.Thread(target=_run_flask, daemon=True)
            server_thread.start()
            icon_path = _write_app_icon_files()
            _ensure_gamevault_shortcuts(icon_path)  # create/update normal Desktop + Start shortcuts
            time.sleep(0.7)  # give Flask a moment to bind before we load the page

            try:
                webview.create_window("Game Vault", "http://127.0.0.1:5000",
                                       width=1280, height=860, min_size=(900, 600),
                                       js_api=DesktopApi(),
                                       # matches the page's own --bg so a fullscreen
                                       # resize never flashes white before the page
                                       # has painted / the webview control catches up
                                       background_color="#14181f")
            except TypeError:
                # Older pywebview build whose create_window() doesn't accept
                # background_color — retry without it rather than crashing.
                webview.create_window("Game Vault", "http://127.0.0.1:5000",
                                       width=1280, height=860, min_size=(900, 600),
                                       js_api=DesktopApi())

            try:
                webview.start(icon=icon_path)  # icon= only takes effect on GTK/Qt (Linux)
            except TypeError:
                webview.start()  # older pywebview without the icon= kwarg
        else:
            # Fallback: pywebview isn't installed, so behave like the old
            # local-web-app version and just open a browser tab.
            print("pywebview not installed — falling back to browser mode.")
            print("For a real desktop window: pip install pywebview")
            threading.Timer(0.8, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
            app.run(debug=False, port=5000)
    except Exception as e:
        import traceback
        _report_fatal_error(traceback.format_exc())
        raise
