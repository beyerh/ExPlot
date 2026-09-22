"""
Platform helpers for the ExPlot GUI: interface scaling and file drag & drop.

Both are optional conveniences: if something is unavailable the app runs unchanged.
"""

import json
import os
import shutil
import subprocess
import sys
import tkinter.font as tkfont
from pathlib import Path

UI_SCALE_OPTIONS = ["Auto", "100%", "125%", "150%", "175%", "200%", "250%", "300%"]
DATA_FILE_EXTENSIONS = (".xlsx", ".xls", ".csv", ".tsv", ".txt")
PROJECT_FILE_EXTENSIONS = (".explt",)


def config_dir():
    """User config directory for settings (cross-platform)."""
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "ExPlot"
    if sys.platform.startswith("win"):
        return Path(os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming"))) / "ExPlot"
    return Path.home() / ".config" / "ExPlot"


# ----------------------------------------------------------------------------
# Interface scaling
# ----------------------------------------------------------------------------

def _run(cmd):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=2).stdout
    except Exception:
        return ""


def detect_linux_scale(root=None):
    """Desktop scale factor for X11 / XWayland apps on Linux, 1.0 if unknown.

    Windows and macOS scale Tk themselves, so this is only used on Linux.
    Sources, in order: EXPLOT_UI_SCALE, GDK_SCALE, Xft.dpi (set by KDE/GNOME/xrdb),
    and compositors that leave XWayland apps unscaled (niri, Hyprland) – the latter
    only if the X screen is larger than the compositor's logical desktop, i.e.
    X apps really see unscaled native pixels.
    """
    for var in ("EXPLOT_UI_SCALE", "GDK_SCALE"):
        try:
            value = float(os.environ.get(var, ""))
            if value > 0:
                return value
        except ValueError:
            pass
    if shutil.which("xrdb"):
        for line in _run(["xrdb", "-query"]).splitlines():
            if line.startswith("Xft.dpi:"):
                try:
                    dpi = float(line.split(":", 1)[1])
                    if dpi > 0 and dpi != 96:
                        return dpi / 96
                except ValueError:
                    pass
    if os.environ.get("WAYLAND_DISPLAY") and root is not None:
        outputs = []  # (logical width, scale)
        try:
            if shutil.which("niri"):
                data = json.loads(_run(["niri", "msg", "--json", "outputs"]) or "{}")
                outputs = [(o["logical"]["width"], o["logical"]["scale"]) for o in data.values() if o.get("logical")]
            elif shutil.which("hyprctl"):
                data = json.loads(_run(["hyprctl", "-j", "monitors"]) or "[]")
                outputs = [(m["width"] / m.get("scale", 1), m.get("scale", 1)) for m in data]
        except Exception:
            outputs = []
        if outputs:
            logical_width = sum(w for w, _ in outputs)
            if root.winfo_screenwidth() > logical_width * 1.1:
                return max(s for _, s in outputs)
    return 1.0


def read_ui_scale_preference():
    try:
        with open(config_dir() / "default_settings.json", encoding="utf-8") as f:
            return json.load(f).get("ui_scale", "Auto")
    except Exception:
        return "Auto"


def apply_ui_scaling(root, preference=None):
    """Scale fonts and widgets. Call right after creating the Tk root, before building widgets.

    ``preference`` is "Auto" or a percentage like "150%". "Auto" only changes
    anything on Linux; on Windows and macOS it leaves Tk's native scaling alone.
    Returns the applied factor.
    """
    preference = preference or read_ui_scale_preference()
    if preference == "Auto":
        factor = detect_linux_scale(root) if sys.platform.startswith("linux") else 1.0
    else:
        try:
            factor = float(str(preference).rstrip("%")) / 100
        except ValueError:
            factor = 1.0
    if factor <= 0 or abs(factor - 1.0) < 0.01:
        return 1.0
    root.tk.call("tk", "scaling", float(root.tk.call("tk", "scaling")) * factor)
    # Fonts defined in pixels (negative size) do not follow 'tk scaling'
    for name in tkfont.names(root):
        font = tkfont.nametofont(name, root=root)
        size = font.cget("size")
        if size < 0:
            font.configure(size=round(size * factor))
    return factor


# ----------------------------------------------------------------------------
# File drag & drop
# ----------------------------------------------------------------------------

def drop_events_reach_app():
    """Whether files dropped from other apps can actually reach the Tk window.

    Tk is an X11 app. On Wayland desktops whose XWayland bridge doesn't forward
    drag & drop from Wayland clients, registration succeeds but no drop events
    ever arrive. Known case: niri + xwayland-satellite (drops unimplemented as
    of satellite 0.8, Supreeeme/xwayland-satellite issues #133/#330). GNOME, KDE,
    Hyprland, Sway and plain X11 bridge it fine.
    """
    if not (sys.platform.startswith("linux") and os.environ.get("WAYLAND_DISPLAY")):
        return True
    desktop = os.environ.get("XDG_CURRENT_DESKTOP", "").lower()
    if "niri" in desktop:
        return False
    return True


def enable_file_drop(root, widgets, on_files):
    """Make ``widgets`` accept dropped files; ``on_files(list_of_paths)`` is called on drop.

    Uses tkinterdnd2 (tkdnd). Returns False if drag & drop is unavailable.
    """
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        TkinterDnD._require(root)
    except Exception as e:
        print(f"Drag and drop unavailable: {e}")
        return False

    def handle(data):
        paths = [p for p in root.tk.splitlist(data) if p]
        root.after_idle(lambda: on_files(paths))
        return "copy"

    # Plain Tcl calls also work for the Tk root (tkinterdnd2's widget methods do not).
    # tkdnd walks up from the widget under the cursor, so registering the toplevel covers the whole window.
    command = root.register(handle)
    for widget in widgets:
        try:
            root.tk.call("tkdnd::drop_target", "register", widget._w, (DND_FILES,))
            root.tk.call("bind", widget._w, "<<Drop>>", f"{command} %D")
        except Exception as e:
            print(f"Could not register drop target {widget}: {e}")
            return False
    return True
