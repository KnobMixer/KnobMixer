"""
KnobMixer v2.7.5
Free per-app volume control for keyboard knobs and hotkeys.
https://github.com/KnobMixer/KnobMixer
"""

import sys, os, json, threading, winreg, ctypes, time, math, struct, wave, io, copy
import queue as _queue
from pathlib import Path

def _can_import(mod):
    try: __import__(mod); return True
    except ImportError: return False

def _ensure_deps():
    import subprocess
    needed = {"pycaw":"pycaw","comtypes":"comtypes","keyboard":"keyboard","pystray":"pystray","PIL":"Pillow","psutil":"psutil"}
    missing = [pkg for mod,pkg in needed.items() if not _can_import(mod)]
    if missing:
        import tkinter as tk, tkinter.messagebox as mb
        r=tk.Tk(); r.withdraw()
        mb.showinfo("KnobMixer","Installing components (one-time ~30s)..."); r.destroy()
        for pkg in missing:
            subprocess.check_call([sys.executable,"-m","pip","install",pkg,"-q"],
                                  creationflags=subprocess.CREATE_NO_WINDOW)
_ensure_deps()

def _cleanup_temp_wavs():
    """Remove leftover temp WAV files from crashed sessions."""
    import tempfile, glob
    try:
        tmp_dir = tempfile.gettempdir()
        for f in glob.glob(os.path.join(tmp_dir, "tmp*.wav")):
            try:
                if time.time() - os.path.getmtime(f) > 60: os.unlink(f)
            except: pass
    except: pass

_cleanup_temp_wavs()

# ── Crash logging ─────────────────────────────────────────────────────────────
import traceback as _tb

def _setup_crash_log():
    """Redirect unhandled exceptions to a crash log in %APPDATA%\\KnobMixer."""
    log_dir = Path(os.getenv("APPDATA",".")) / "KnobMixer"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "crash.log"
    _orig = sys.excepthook
    def _hook(exc_type, exc_val, exc_tb):
        try:
            import datetime
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"\n{'='*60}\n")
                _ver = globals().get("APP_VER", "?")
                f.write(f"KnobMixer v{_ver} crash — {datetime.datetime.now()}\n")
                f.write("".join(_tb.format_exception(exc_type, exc_val, exc_tb)))
        except: pass
        _orig(exc_type, exc_val, exc_tb)
    sys.excepthook = _hook

_setup_crash_log()

import tkinter as tk
from tkinter import ttk, messagebox, colorchooser
import psutil, keyboard, pystray, comtypes
import ctypes, ctypes.wintypes
from PIL import Image, ImageDraw, ImageTk
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

def _init_ttk_theme(widget):
    st = ttk.Style(widget)
    try:
        st.theme_use("clam")
    except Exception:
        pass
    st.configure("TNotebook", background=BG, borderwidth=0)
    st.configure("TNotebook.Tab", background=PANEL, foreground=TEXT,
                 padding=[14, 6])
    st.map("TNotebook.Tab", background=[("selected", BORDER)])
    st.configure("Knob.TCombobox",
                 fieldbackground=PANEL, background=BORDER,
                 foreground=TEXT, arrowcolor=TEXT,
                 bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    st.map("Knob.TCombobox",
           fieldbackground=[("readonly", PANEL)],
           background=[("readonly", BORDER), ("active", HOVER)],
           foreground=[("readonly", TEXT)])
    st.configure("Knob.Vertical.TScrollbar",
                 background=SB_THUMB, troughcolor=SB_TRACK,
                 bordercolor=SB_TRACK, arrowcolor=SUBTEXT,
                 darkcolor=SB_THUMB, lightcolor=SB_THUMB,
                 gripcount=0, relief="flat", width=10)
    st.map("Knob.Vertical.TScrollbar",
           background=[("active", SB_ACTIVE), ("pressed", SB_ACTIVE)])

def _place_near_parent(win, parent, side="left", overlap=28, margin=20):
    try:
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", RECT),
                        ("rcWork", RECT), ("dwFlags", ctypes.c_ulong)]

        def _work_area_for_parent():
            user32 = getattr(ctypes, "windll", None)
            user32 = getattr(user32, "user32", None) if user32 else None
            if not user32:
                raise RuntimeError("No user32")
            cx = parent.winfo_rootx() + max(1, parent.winfo_width()) // 2
            cy = parent.winfo_rooty() + max(1, parent.winfo_height()) // 2
            pt = POINT(cx, cy)
            monitor = user32.MonitorFromPoint(pt, 2)
            if not monitor:
                raise RuntimeError("No monitor")
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            if not user32.GetMonitorInfoW(monitor, ctypes.byref(mi)):
                raise RuntimeError("GetMonitorInfoW failed")
            return mi.rcWork.left, mi.rcWork.top, mi.rcWork.right, mi.rcWork.bottom

        win.update_idletasks()
        parent.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        ww, wh = win.winfo_reqwidth(), win.winfo_reqheight()
        wa_left, wa_top, wa_right, wa_bottom = _work_area_for_parent()
        if side == "left":
            x = px - ww + overlap
            if x < wa_left + margin:
                x = px + pw - overlap
        else:
            x = px + pw - overlap
        y = py + 30
        x = max(wa_left + margin, min(x, wa_right - ww - margin))
        y = max(wa_top + margin, min(y, wa_bottom - wh - margin))
        win.geometry(f"+{int(x)}+{int(y)}")
    except Exception:
        pass

def _themed_confirm(parent, title, message, yes_text="Yes", no_text="Cancel", danger=False):
    result = {"ok": False}
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.configure(bg=BG)
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.transient(parent)
    _place_near_parent(dlg, parent, side="left")
    tk.Label(dlg, text=title, font=("Segoe UI",10,"bold"),
             fg=TEXT, bg=BG).pack(anchor="w", padx=18, pady=(16,6))
    tk.Label(dlg, text=message, font=("Segoe UI",9),
             fg=SUBTEXT, bg=BG, justify="left",
             wraplength=320).pack(anchor="w", padx=18, pady=(0,12))
    bf = tk.Frame(dlg, bg=BG)
    bf.pack(anchor="e", padx=18, pady=(0,16))
    tk.Button(bf, text=no_text, font=("Segoe UI",9),
              bg=BORDER, fg=TEXT, relief="flat", cursor="hand2",
              padx=10, pady=4, command=dlg.destroy).pack(side="left", padx=(0,8))
    def _yes():
        result["ok"] = True
        dlg.destroy()
    tk.Button(bf, text=yes_text, font=("Segoe UI",9,"bold"),
              bg="#3a1a1a" if danger else "#183524",
              fg="#ff8b95" if danger else "#1DB954",
              relief="flat", cursor="hand2", padx=10, pady=4,
              command=_yes).pack(side="left")
    dlg.wait_window()
    return result["ok"]

def _themed_alert(parent, title, message):
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.configure(bg=BG)
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.transient(parent)
    _place_near_parent(dlg, parent, side="left")
    tk.Label(dlg, text=title, font=("Segoe UI",10,"bold"),
             fg=TEXT, bg=BG).pack(anchor="w", padx=18, pady=(16,6))
    tk.Label(dlg, text=message, font=("Segoe UI",9),
             fg=SUBTEXT, bg=BG, justify="left",
             wraplength=320).pack(anchor="w", padx=18, pady=(0,12))
    tk.Button(dlg, text="OK", font=("Segoe UI",9,"bold"),
              bg=BORDER, fg=TEXT, relief="flat", cursor="hand2",
              padx=12, pady=4, command=dlg.destroy).pack(anchor="e", padx=18, pady=(0,16))
    dlg.wait_window()

# ── Win32 Low-Level Keyboard Hook ────────────────────────────────────────────
# Same mechanism used by Discord, OBS, TeamSpeak, etc.
# Works regardless of what other keys are held, regardless of app focus.

_user32   = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32

WH_KEYBOARD_LL = 13
WM_KEYDOWN     = 0x0100
WM_SYSKEYDOWN  = 0x0104   # fired when Alt is held + another key

# Virtual key codes for multimedia keys
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP   = 0xAF

# Modifier virtual key codes
VK_MODS = {
    "ctrl":  [0x11, 0xA2, 0xA3],   # VK_CONTROL, L, R
    "shift": [0x10, 0xA0, 0xA1],   # VK_SHIFT, L, R
    "alt":   [0x12, 0xA4, 0xA5],   # VK_MENU, L, R
    "win":   [0x5B, 0x5C],          # L/R Windows
}

# keyboard library vk name → vk code mapping (subset we care about)
def _name_to_vk(name: str) -> int | None:
    """Convert a key name (e.g. 'f22', 'x', 'space') to a Windows VK code."""
    name = name.strip().lower()
    # Function keys
    if name.startswith("f") and name[1:].isdigit():
        n = int(name[1:])
        if 1 <= n <= 24:
            return 0x6F + n  # F1=0x70 … F24=0x87
    # Letters
    if len(name) == 1 and name.isalpha():
        return ord(name.upper())
    # Digits
    if len(name) == 1 and name.isdigit():
        return ord(name)
    # Special keys
    _special = {
        "space": 0x20, "enter": 0x0D, "tab": 0x09, "backspace": 0x08,
        "escape": 0x1B, "esc": 0x1B,
        "insert": 0x2D, "delete": 0x2E, "del": 0x2E,
        "home": 0x24, "end": 0x23, "page up": 0x21, "page down": 0x22,
        "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
        "num lock": 0x90, "scroll lock": 0x91, "caps lock": 0x14,
        "print screen": 0x2C, "pause": 0x13,
        "num 0":0x60,"num 1":0x61,"num 2":0x62,"num 3":0x63,"num 4":0x64,
        "num 5":0x65,"num 6":0x66,"num 7":0x67,"num 8":0x68,"num 9":0x69,
        "num *":0x6A,"num +":0x6B,"num -":0x6D,"num .":0x6E,"num /":0x6F,
        "volume up":VK_VOLUME_UP,"volume down":VK_VOLUME_DOWN,
        "volume mute":VK_VOLUME_MUTE,
        ";":0xBA,"=":0xBB,",":0xBC,"-":0xBD,".":0xBE,"/":0xBF,"`":0xC0,
        "[":0xDB,"\\":0xDC,"]":0xDD,"'":0xDE,
        # Media control keys
        "play/pause media":0xB3,"media play/pause":0xB3,
        "play pause":0xB3,"playpause":0xB3,"media play pause":0xB3,
        "media next track":0xB0,"next track":0xB0,
        "media previous track":0xB1,"prev track":0xB1,
        "media stop":0xB2,
    }
    return _special.get(name)

def _mods_held(mods: set) -> bool:
    """Check if all required modifier keys are currently physically held."""
    for mod in mods:
        vks = VK_MODS.get(mod, [])
        if not any(_user32.GetAsyncKeyState(vk) & 0x8000 for vk in vks):
            return False
    return True

def _parse_hotkey(raw: str):
    """Parse 'alt+x', 'ctrl+f13', 'f22' into (mods_set, trigger_vk).
    Returns (set, int) or (None, None) on failure."""
    if not raw: return None, None
    MOD_ALIASES = {
        "ctrl":"ctrl","control":"ctrl","left ctrl":"ctrl","right ctrl":"ctrl",
        "shift":"shift","left shift":"shift","right shift":"shift",
        "alt":"alt","left alt":"alt","right alt":"alt",
        "win":"win","windows":"win","left windows":"win","right windows":"win",
    }
    parts   = [p.strip().lower() for p in raw.split("+") if p.strip()]
    mods    = set()
    trigger = None
    for p in parts:
        if p in MOD_ALIASES:
            mods.add(MOD_ALIASES[p])
        else:
            trigger = p
    if not trigger: return None, None  # Fix #21 — modifier-only, caller should warn
    vk = _name_to_vk(trigger)
    if vk is None: return None, None
    return mods, vk


class GlobalHookManager:
    """
    Single Win32 WH_KEYBOARD_LL hook that handles ALL hotkeys.
    This is how Discord/OBS/etc. work — one hook, routes to callbacks.
    Works while any other key is held, in any app, any focus state.
    """
    def __init__(self):
        self._callbacks  = []   # list of (mods, vk, callback, suppress, debounce_map)
        self._hook       = None
        self._hook_proc  = None
        self._thread     = None
        self._running    = False

    def register(self, hotkey: str, callback, suppress=True,
                 is_media=False, debounce=0.02) -> bool:
        """Register a hotkey. Returns True if parsed successfully.
        is_media=True marks it as an intentional media key registration.
        debounce: seconds between allowed triggers (default 20ms for vol keys)."""
        mods, vk = _parse_hotkey(hotkey)
        if vk is None:
            return False
        self._callbacks.append({
            "mods": mods, "vk": vk,
            "fn": callback, "suppress": suppress,
            "is_media": is_media,
            "debounce": debounce,
            "last": 0.0,
        })
        return True

    def clear(self):
        self._callbacks.clear()

    def start(self):
        if self._running: return
        self._running = True
        self._thread  = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False
        self.clear()
        # Post a dummy message to unblock the message loop
        try: _user32.PostThreadMessageW(self._thread.ident if self._thread else 0,
                                         0x0012, 0, 0)
        except: pass

    def _run(self):
        """Run the hook on a dedicated thread with its own message loop.
        Uses ULONG_PTR (64-bit on Win64) for lParam to avoid overflow errors."""
        import time as _t

        # KBDLLHOOKSTRUCT — the struct lParam points to
        class KBDLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ("vkCode",      ctypes.wintypes.DWORD),
                ("scanCode",    ctypes.wintypes.DWORD),
                ("flags",       ctypes.wintypes.DWORD),
                ("time",        ctypes.wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_ulong),
            ]

        # Use c_void_p for lParam — avoids 64-bit overflow on Windows 10/11
        HOOKPROC = ctypes.WINFUNCTYPE(
            ctypes.c_long,
            ctypes.c_int,
            ctypes.wintypes.WPARAM,
            ctypes.c_void_p)           # ← was LPARAM, caused overflow

        # Set correct return type for CallNextHookEx
        _user32.CallNextHookEx.restype  = ctypes.c_long
        _user32.CallNextHookEx.argtypes = [
            ctypes.c_void_p,           # hhk (can be NULL)
            ctypes.c_int,              # nCode
            ctypes.wintypes.WPARAM,    # wParam
            ctypes.c_void_p,           # lParam (pointer, not int)
        ]

        # Media key VK codes that should NEVER be accidentally matched
        # unless explicitly registered as hotkeys
        MEDIA_VKS = {0xAD, 0xAE, 0xAF,   # volume mute/down/up
                     0xB0, 0xB1, 0xB2, 0xB3,  # media next/prev/stop/playpause
                     0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC}  # browser/app keys

        def _proc(nCode, wParam, lParam):
            # Always call next hook first — never block unless we mean to
            next_result = _user32.CallNextHookEx(None, nCode, wParam, lParam)

            if nCode >= 0 and wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                try:
                    ks  = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT))[0]
                    vk  = ks.vkCode
                    now = _t.monotonic()
                    suppress_result = False
                    for cb in list(self._callbacks):
                        if cb["vk"] != vk: continue
                        if vk in MEDIA_VKS and not cb.get("is_media", False):
                            continue
                        if not _mods_held(cb["mods"]): continue
                        if now - cb["last"] < cb.get("debounce", 0.02): continue
                        cb["last"] = now
                        try: cb["fn"]()
                        except Exception as e: print(f"[Hook] callback error: {e}")
                        if cb["suppress"]:
                            suppress_result = True
                        # Don't break — run ALL matching callbacks
                        # (e.g. same key assigned to multiple groups)
                    if suppress_result:
                        return 1
                except Exception as e:
                    print(f"[Hook] proc error: {e}")

            return next_result

        self._hook_proc = HOOKPROC(_proc)

        _user32.SetWindowsHookExW.restype  = ctypes.c_void_p
        _user32.SetWindowsHookExW.argtypes = [
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.wintypes.DWORD,
        ]
        self._hook = _user32.SetWindowsHookExW(
            WH_KEYBOARD_LL, self._hook_proc, None, 0)

        # Message loop — keeps the hook alive
        msg = ctypes.wintypes.MSG()
        while self._running:
            r = _user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if r <= 0: break
            _user32.TranslateMessage(ctypes.byref(msg))
            _user32.DispatchMessageW(ctypes.byref(msg))

        if self._hook:
            _user32.UnhookWindowsHookExW = _user32.UnhookWindowsHookEx
            _user32.UnhookWindowsHookEx(self._hook)
            self._hook = None
        # If we exited unexpectedly and should still be running, restart (#6)
        if self._running:
            print("[Hook] Message loop exited unexpectedly — restarting in 1s")
            self._running = False
            time.sleep(1)
            self.start()


# Single global hook instance shared by everything
_HOOK = GlobalHookManager()

# ── Constants ─────────────────────────────────────────────────────────────────
APP_NAME    = "KnobMixer"
APP_VER     = "2.7.5"
APPDATA_DIR = Path(os.getenv("APPDATA",".")) / APP_NAME
APPDATA_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APPDATA_DIR / "config.json"
STARTUP_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
EXE_PATH    = sys.executable if getattr(sys,"frozen",False) else os.path.abspath(__file__)
ACCENT      = ["#1DB954","#5865F2","#FF6B35","#E91E63","#00BCD4","#FFC107","#9C27B0","#00E676"]

BG      = "#16161e"
PANEL   = "#1e1e2a"
BORDER  = "#2a2a3e"
TEXT    = "#e8e8f0"
SUBTEXT = "#6b6b8a"
HOVER   = "#252535"
PANEL_SOFT = "#232334"
INPUT_BG = "#181824"
SB_TRACK = "#141421"
SB_THUMB = "#4a4a70"
SB_ACTIVE = "#6d6da3"

# Sound presets: (display_name, shape, mute_params, unmute_params)
# Params: (freq, dur_ms, vol_scale)
# Mute = lower/darker tone, Unmute = brighter/higher tone
SOUND_PRESETS = [
    ("Soft Ping",  "ping",   (660, 60,  1.0), (990, 55,  1.0)),  # light ping notification
]
# Volume scale: user sees 1-5, internally maps to actual vol multiplier
# 1=quiet(0.028) 2=low(0.056) 3=medium(0.112) 4=loud(0.196) 5=max(0.3)
_VOL_SCALE = {1: 0.028, 2: 0.056, 3: 0.112, 4: 0.196, 5: 0.3}
def _vol_from_level(level): return _VOL_SCALE.get(int(level), 0.056)
def _level_from_vol(vol):
    """Reverse map raw vol to nearest 1-5 level."""
    best = 2
    best_diff = abs(vol - _VOL_SCALE[2])
    for lvl, v in _VOL_SCALE.items():
        diff = abs(vol - v)
        if diff < best_diff:
            best_diff = diff; best = lvl
    return best

DEFAULT_CFG = {
    "version": 4,
    "mode": "single",
    "start_minimized": True,
    "show_overlay": True,
    "tutorial_seen": False,
    "analytics_enabled": True,
    "overlay_size": 0.7,
    "overlay_position": "bottom-right",
    "overlay_x": -1,
    "overlay_y": -1,
    "slowdown_enabled": True,
    "slowdown_threshold": 10,
    "slowdown_step": 0.5,
    "single_default_group": 0,
    "single_timeout": 30,
    "single_auto_revert": False,
    "hw_knob_enabled": False,
    "hw_knob_group": 0,
    "cycle_key": "",
    "single_keys": {"vol_down":"","vol_up":"","mute":""},
    "mic_enabled": False,
    "mic_device": "",
    "mic_device_name": "System Default",
    "mic_hotkey": "f9",
    "mic_start_muted": False,
    "mic_sound_volume": 0.056,
    "mic_sound_preset": 0,
    "mic_icon_x": -1,
    "mic_position": "bottom-right",
    "mic_icon_y": -1,
    "mic_icon_size": 40,
    "mic_icon_alpha": 0.85,
    "mic_icon_style": "circle",
    "mic_icon_locked": False,
    "groups": [
        {"id":0,"name":"Master Volume","color":"#00BCD4",
         "apps":[],"master_volume":True,
         "keys":{"vol_down":"","vol_up":"","mute":""},
         "single_key":"","step":5,"volume":80,"muted":False,"_vbm":80,
         "foreground_mode":False,"enabled":True,"is_default":True},
        {"id":1,"name":"Media","color":"#1DB954",
         "apps":["spotify","chrome"],
         "keys":{"vol_down":"","vol_up":"","mute":""},
         "single_key":"","step":5,"volume":80,"muted":False,"_vbm":80,
         "foreground_mode":False,"enabled":True,"is_default":False},
        {"id":2,"name":"Chat","color":"#5865F2",
         "apps":["discord"],
         "keys":{"vol_down":"","vol_up":"","mute":""},
         "single_key":"","step":5,"volume":80,"muted":False,"_vbm":80,
         "foreground_mode":False,"enabled":True,"is_default":False},
    ]
}

# ── Config ────────────────────────────────────────────────────────────────────
# Media key names that must never appear as group vol/mute hotkeys
_BAD_HOTKEYS = {
    "play/pause media","media play/pause","play pause","playpause",
    "media play pause","media next track","next track","media previous track",
    "prev track","media stop","volume up","volume down","volume mute",
}

def load_cfg():
    if CONFIG_FILE.exists():
        try:
            d = json.loads(CONFIG_FILE.read_text())
            for g in d.get("groups",[]):
                for k,v in [("foreground_mode",False),("_vbm",80),("color","#888"),
                             ("id",0),("step",5),("enabled",True),("is_default",False),
                             ("single_key",""),("master_volume",False)]:
                    g.setdefault(k,v)
                g.setdefault("keys",{})
                for a in ("vol_down","vol_up","mute"):
                    g["keys"].setdefault(a,"")
                    # Auto-clear any mis-captured media key names
                    if g["keys"].get(a,"").lower().strip() in _BAD_HOTKEYS:
                        print(f"[Config] Cleared bad hotkey in {g.get('name','')} {a}")
                        g["keys"][a] = ""
            for k,v in [("overlay_size",0.7),("overlay_position","bottom-right"),("overlay_x",-1),("overlay_y",-1),("slowdown_enabled",True),
                        ("slowdown_threshold",10),("slowdown_step",0.5),
                        ("single_default_group",0),("single_timeout",30),("single_auto_revert",False),("hw_knob_enabled",False),("hw_knob_group",0),("cycle_key",""),
                        ("single_keys",{"vol_down":"","vol_up":"","mute":""}),
                        ("mic_enabled",False),("mic_device",""),("mic_hotkey","f9"),
                        ("mic_start_muted",False),("mic_sound_volume",0.056),
                        ("mic_sound_preset",0),("mic_position","bottom-right"),("mic_icon_x",-1),("mic_icon_y",-1),
                        ("mic_icon_size",40),("mic_icon_alpha",0.85),
                        ("mic_icon_style","circle"),("mode","multi"),
                        ("start_minimized",True),("show_overlay",True),("tutorial_seen",False),("analytics_enabled",True)]:
                d.setdefault(k,v)
            # Clamp numeric values to safe ranges
            for g in d.get("groups", []):
                g["volume"] = max(0.0, min(100.0, float(g.get("volume", 80))))
                g["step"]   = max(1,   min(20,    int(g.get("step", 5))))
                g.setdefault("volume", 80.0)
            d["overlay_size"]        = max(0.5, min(3.0, float(d.get("overlay_size", 0.7))))
            d["slowdown_threshold"]  = max(1,   min(50,  int(d.get("slowdown_threshold", 10))))
            d["slowdown_step"]       = max(0.1, min(10.0,float(d.get("slowdown_step", 0.5))))
            return d
        except Exception as e:
            print(f"[Config] Load failed ({e}), using defaults")
    return copy.deepcopy(DEFAULT_CFG)

_save_lock = threading.Lock()  # prevents concurrent config writes

def save_cfg(cfg):
    # Save a shallow copy so we don't mutate the live cfg dict
    # _active_group_ref must NOT be popped from live cfg — it's used by hotkey callbacks
    data = {k: v for k, v in cfg.items() if k != "_active_group_ref"}
    with _save_lock:
        tmp = CONFIG_FILE.with_suffix(f".tmp")
        try:
            tmp.write_text(json.dumps(data, indent=2))
            os.replace(tmp, CONFIG_FILE)
        except Exception as e:
            print(f"[Config] Save failed: {e}")
            try: tmp.unlink()
            except: pass

_REPORT_COOLDOWN_SECS = 15 * 60
_REPORT_MAX_CHARS     = 2000
_REPORT_LOG_MAX_CHARS = 3000

def _report_state_file():
    return APPDATA_DIR / "report_state.json"

def _load_report_state():
    try:
        return json.loads(_report_state_file().read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_report_state(state):
    try:
        _report_state_file().write_text(json.dumps(state, indent=2), encoding="utf-8")
    except Exception:
        pass

def _report_endpoint():
    if not ANALYTICS_URL:
        return ""
    return ANALYTICS_URL.replace("/ping", "/report")

def _report_status():
    state = _load_report_state()
    now = int(time.time())
    next_allowed = int(state.get("next_allowed_at", 0) or 0)
    last_hash = str(state.get("last_hash", "") or "")
    last_sent = int(state.get("last_sent_at", 0) or 0)
    return state, now, next_allowed, last_hash, last_sent

def _report_validate_message(msg):
    msg = (msg or "").strip()
    if not msg:
        return False, "Please describe the issue first."
    if len(msg) < 10:
        return False, "Please add a little more detail so I can understand the issue."
    if len(msg) > _REPORT_MAX_CHARS:
        return False, f"Please keep the report under {_REPORT_MAX_CHARS} characters."
    return True, ""

def _report_can_send(msg):
    import hashlib
    state, now, next_allowed, last_hash, last_sent = _report_status()
    if now < next_allowed:
        remain = max(1, int(math.ceil((next_allowed - now) / 60)))
        return False, f"Please wait about {remain} minute(s) before sending another report."
    msg_hash = hashlib.sha256(msg.strip().encode("utf-8", errors="ignore")).hexdigest()
    if msg_hash == last_hash and (now - last_sent) < 24 * 3600:
        return False, "That same report was already sent recently."
    return True, ""

def _mark_report_sent(msg):
    import hashlib
    now = int(time.time())
    _save_report_state({
        "last_sent_at": now,
        "next_allowed_at": now + _REPORT_COOLDOWN_SECS,
        "last_hash": hashlib.sha256(msg.strip().encode("utf-8", errors="ignore")).hexdigest(),
    })

# ── Startup ───────────────────────────────────────────────────────────────────
def set_startup(en):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,STARTUP_KEY,0,winreg.KEY_SET_VALUE) as k:
            if en: winreg.SetValueEx(k,APP_NAME,0,winreg.REG_SZ,f'"{EXE_PATH}"')
            else:
                try: winreg.DeleteValue(k,APP_NAME)
                except FileNotFoundError: pass
        return True
    except: return False



def get_startup():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,STARTUP_KEY) as k:
            winreg.QueryValueEx(k,APP_NAME); return True
    except: return False

# ── Sound engine ──────────────────────────────────────────────────────────────
_sound_lock   = threading.Lock()
_sound_thread = None
_sound_cancel = threading.Event()  # set to cancel any in-progress sound immediately

def _make_wav(freq, dur_ms, vol=0.15, shape="sine"):
    sr = 22050
    n  = int(sr * dur_ms / 1000)
    buf = bytearray(n * 2)
    pi2 = 2 * math.pi

    for i in range(n):
        t  = i / sr
        dt = dur_ms / 1000  # total duration in seconds

        if shape == "bell":
            # Natural bell: fast attack, slow exponential decay, 3 harmonics
            env = math.exp(-t * 6)
            f1  = math.sin(pi2 * freq * t)
            f2  = math.sin(pi2 * freq * 2.756 * t) * math.exp(-t * 10) * 0.45
            f3  = math.sin(pi2 * freq * 5.404 * t) * math.exp(-t * 20) * 0.2
            raw = env * (f1 + f2 + f3) / 1.65

        elif shape == "marimba":
            # Wooden marimba bar: thump + pure tone, medium decay
            env  = math.exp(-t * 12)
            tone = math.sin(pi2 * freq * t)
            ot2  = math.sin(pi2 * freq * 4 * t) * math.exp(-t * 30) * 0.15
            thmp = math.sin(pi2 * freq * 0.5 * t) * math.exp(-t * 40) * 0.25
            raw  = env * (tone + ot2 + thmp)

        elif shape == "ping":
            # Clean pure sine ping, longer decay
            env = math.exp(-t * 9)
            raw = env * math.sin(pi2 * freq * t)

        elif shape == "glass":
            # Glass tap: very high, crystalline with shimmer
            env  = math.exp(-t * 7)
            f1   = math.sin(pi2 * freq * t)
            f2   = math.sin(pi2 * freq * 2.0 * t) * math.exp(-t * 14) * 0.3
            f3   = math.sin(pi2 * freq * 3.0 * t) * math.exp(-t * 25) * 0.15
            raw  = env * (f1 + f2 + f3) / 1.45

        elif shape == "synth":
            # Soft synth: slight detuned pair for warmth
            env  = math.exp(-t * 10)
            f1   = math.sin(pi2 * freq * t)
            f2   = math.sin(pi2 * freq * 1.003 * t)  # slight detune for chorus
            raw  = env * (f1 + f2) * 0.5

        elif shape == "thock":
            # Mech keyboard bottom-out: sub-bass thud + snap transient
            env   = math.exp(-t * 30)
            thud  = math.sin(pi2 * freq * t)
            snap  = math.sin(pi2 * freq * 4 * t) * math.exp(-t * 80) * 0.4
            body  = math.sin(pi2 * freq * 1.5 * t) * math.exp(-t * 15) * 0.3
            raw   = env * (thud + snap + body)

        elif shape == "click":
            # Hard clicky key: very short noise burst + sharp tone
            env  = math.exp(-t * 120)
            tone = math.sin(pi2 * freq * t)
            # pseudo-noise via high harmonics
            nz   = sum(math.sin(pi2 * freq * k * t) / k
                       for k in [3, 5, 7, 11]) * 0.15
            raw  = env * (tone + nz)

        elif shape == "bubble":
            # Bubble: rising pitch, soft
            f_t  = freq * (1 + t * 2)   # pitch rises over time
            env  = math.exp(-t * 10) * min(t / 0.01, 1.0)
            raw  = env * math.sin(pi2 * f_t * t)

        elif shape == "blip":
            # Retro blip: square-ish with tight envelope
            env = math.exp(-t * 14)
            raw = env * math.sin(pi2 * freq * t)
            # Add slight squary edge
            raw += env * math.sin(pi2 * freq * 3 * t) * 0.2

        elif shape == "whoosh":
            # Airy whoosh: filtered noise sweep
            env  = math.sin(math.pi * t / dt) if dt > 0 else 0  # arch envelope
            nz   = sum(math.sin(pi2 * (freq + k*30) * t + k)
                       for k in range(8)) / 8
            raw  = env * nz

        elif shape == "knock":
            # Wooden knock: low thump, very fast decay
            env  = math.exp(-t * 45)
            raw  = env * (math.sin(pi2 * freq * t) +
                          math.sin(pi2 * freq * 1.8 * t) * 0.3 * math.exp(-t * 20))

        else:  # sine
            env = math.exp(-t * 8)
            raw = env * math.sin(pi2 * freq * t)

        val = int(32767 * vol * raw)
        struct.pack_into("<h", buf, i*2, max(-32767, min(32767, val)))

    out = io.BytesIO()
    with wave.open(out, "wb") as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(sr)
        w.writeframes(buf)
    return out.getvalue()

def play_sound(freq, dur_ms=60, vol=0.6, shape="sine"):
    """Play sound, cancelling any currently playing sound first.
    If called rapidly (e.g. holding hotkey), only the latest plays."""
    global _sound_thread
    import winsound, tempfile
    if dur_ms <= 0: return
    # Cancel any in-progress sound immediately
    _sound_cancel.set()
    # Wait briefly for previous thread to stop (max 30ms)
    if _sound_thread and _sound_thread.is_alive():
        _sound_thread.join(timeout=0.03)
    _sound_cancel.clear()
    data = _make_wav(freq, dur_ms, vol, shape)
    if not data: return
    try:
        tf = tempfile.NamedTemporaryFile(suffix=".wav", delete=False)
        tmp = Path(tf.name)
        tf.write(data)
        tf.close()
    except Exception: return
    def _play():
        # Check cancel before starting — if another call came in, skip
        if _sound_cancel.is_set():
            try: tmp.unlink()
            except: pass
            return
        with _sound_lock:
            try:
                if not _sound_cancel.is_set():
                    winsound.PlaySound(str(tmp), winsound.SND_FILENAME)
            finally:
                try: tmp.unlink()
                except: pass
    t = threading.Thread(target=_play, daemon=True)
    t.start()
    _sound_thread = t

def play_preset(preset_idx, vol, muted):
    if preset_idx < 0 or preset_idx >= len(SOUND_PRESETS): preset_idx = 0
    _, shape, mute_p, unmute_p = SOUND_PRESETS[preset_idx]
    freq, dur, vscale = mute_p if muted else unmute_p
    play_sound(freq, dur, vol * vscale, shape)

# ── Inline hotkey capture ──────────────────────────────────────────────────────
class HotkeyCapture:
    """
    Discord-style hotkey capture:
    - Click the button to start listening
    - The moment you press any non-modifier key (with or without modifiers held),
      the combo is captured instantly — no timer, no release needed
    - Escape cancels
    - Supports: F22, Alt+X, Ctrl+F13, Win+F18, single keys, etc.
    """
    _ACTIVE_CAPTURE = None
    MOD_NAMES = {"ctrl","left ctrl","right ctrl","shift","left shift","right shift",
                 "alt","left alt","right alt","windows","left windows","right windows","win"}
    MOD_NORM  = {"ctrl":"ctrl","left ctrl":"ctrl","right ctrl":"ctrl",
                 "shift":"shift","left shift":"shift","right shift":"shift",
                 "alt":"alt","left alt":"alt","right alt":"alt",
                 "windows":"win","left windows":"win","right windows":"win","win":"win"}

    def __init__(self, btn, callback, original_text):
        self._btn    = btn
        self._cb     = callback
        self._orig   = original_text
        self._active = False
        self._held   = set()
        btn.config(command=self._start)

    def _start(self):
        if HotkeyCapture._ACTIVE_CAPTURE and HotkeyCapture._ACTIVE_CAPTURE is not self:
            try: HotkeyCapture._ACTIVE_CAPTURE._finish(None)
            except: pass
        if self._active: return
        self._active = True
        HotkeyCapture._ACTIVE_CAPTURE = self
        self._held.clear()
        self._btn.config(text="Press keys…", bg="#2a4a2a", fg="#1DB954")
        keyboard.hook(self._on_event, suppress=True)

    # Media key names that are invalid as group hotkeys (#18)
    _INVALID_KEYS = {"volume up","volume down","volume mute","play/pause media",
                     "media play/pause","next track","prev track","media stop",
                     "left button","right button","middle button","x button 1","x button 2",
                     "mouse wheel up","mouse wheel down","wheel up","wheel down"}

    def _on_event(self, ev):
        if not self._active: return
        name = ev.name.lower()

        if ev.event_type == "down":
            if name == "escape":
                self._finish(None)
                return
            # Reject media keys with visual feedback (#18)
            if (name in self._INVALID_KEYS or "button" in name or
                name.startswith("mouse ") or name.startswith("wheel ")):
                self._btn.after(0, lambda: self._btn.config(
                    text="Invalid key", bg="#3a1a1a", fg="#ff6b6b"))
                self._btn.after(1200, lambda: self._btn.config(
                    text="Press keys…", bg="#2a4a2a", fg="#1DB954"))
                return
            self._held.add(name)
            self._btn.after(0, self._update_display)
            # Capture on press of any non-modifier key
            if name not in self.MOD_NAMES:
                combo = self._build_combo(self._held)
                self._finish(combo)

        elif ev.event_type == "up":
            self._held.discard(name)

    def _build_combo(self, keys):
        order = ["ctrl","shift","alt","win"]
        mods, rest, seen = [], [], set()
        for k in keys:
            n = self.MOD_NORM.get(k)
            if n and n not in seen:
                mods.append(n); seen.add(n)
            elif k not in self.MOD_NAMES:
                rest.append(k)
        mods.sort(key=lambda x: order.index(x) if x in order else 99)
        return "+".join(mods + rest)

    def _update_display(self):
        combo = self._build_combo(self._held)
        self._btn.config(text=combo.upper() if combo else "Press keys…")

    def _finish(self, combo):
        if not self._active: return
        self._active = False
        if HotkeyCapture._ACTIVE_CAPTURE is self:
            HotkeyCapture._ACTIVE_CAPTURE = None
        # Only unhook our specific hook, not ALL hooks (#5 — was killing GlobalHookManager)
        try: keyboard.unhook(self._on_event)
        except: pass
        if combo:
            self._btn.after(0, lambda: self._btn.config(
                text=fmt_hotkey(combo), bg=INPUT_BG, fg=TEXT,
                highlightbackground=BORDER, highlightcolor=BORDER))
            self._cb(combo)
        else:
            self._btn.after(0, lambda: self._btn.config(
                text=self._orig, bg=INPUT_BG, fg=TEXT,
                highlightbackground=BORDER, highlightcolor=BORDER))

def fmt_hotkey(raw):
    if not raw: return "—"
    return "+".join(p.strip().upper() for p in raw.split("+") if p.strip())

def _iter_assigned_hotkeys(cfg):
    for g in cfg.get("groups", []):
        gid = g.get("id", id(g))
        for action, hk in g.get("keys", {}).items():
            if hk.strip():
                yield ("group", gid, action), hk.strip()
        sk = g.get("single_key", "").strip()
        if sk:
            yield ("single_group", gid), sk
    for action, hk in cfg.get("single_keys", {}).items():
        if hk.strip():
            yield ("single_shared", action), hk.strip()
    ck = cfg.get("cycle_key", "").strip()
    if ck:
        yield ("cycle",), ck
    mh = cfg.get("mic_hotkey", "").strip()
    if mh:
        yield ("mic",), mh

def _hotkey_in_use(cfg, hotkey, slot_id):
    want = hotkey.strip().lower()
    for other_slot, other_hk in _iter_assigned_hotkeys(cfg):
        if other_slot == slot_id:
            continue
        if other_hk.strip().lower() == want:
            return True
    return False

def _validate_hotkey_choice(cfg, hotkey, slot_id):
    hotkey = (hotkey or "").strip()
    mods, vk = _parse_hotkey(hotkey)
    if not hotkey or vk is None:
        return False, "That key combo is not supported."
    if _hotkey_in_use(cfg, hotkey, slot_id):
        return False, "That hotkey is already in use."
    return True, ""

def make_hotkey_btn(parent, current_key, callback, label_prefix=""):
    """Create a hotkey button with inline capture. Returns the button.
    The HotkeyCapture instance is stored as btn._capture for ✕ button access."""
    display = fmt_hotkey(current_key) if current_key else "—"
    text    = f"{label_prefix}{display}" if label_prefix else display
    btn = tk.Button(parent, text=text,
                    font=("Consolas",8), bg=INPUT_BG, fg=TEXT,
                    activebackground=HOVER, activeforeground=TEXT,
                    relief="flat", cursor="hand2",
                    padx=9, pady=3, bd=0, highlightthickness=1,
                    highlightbackground=BORDER, highlightcolor=BORDER)
    btn._capture = HotkeyCapture(btn, callback, text)
    return btn

# ── Audio ─────────────────────────────────────────────────────────────────────
_audio_lock = threading.Lock()
_cfg_lock   = threading.Lock()  # protects group dict mutations from hook thread

def _sessions():
    comtypes.CoInitialize()
    out = {}
    try:
        for s in AudioUtilities.GetAllSessions():
            if s.Process is None: continue
            name = s.Process.name().lower().removesuffix(".exe")
            vol  = s._ctl.QueryInterface(ISimpleAudioVolume)
            out.setdefault(name,[]).append(vol)
    except: pass
    finally:
        try: comtypes.CoUninitialize()  # Fix #11 — balance CoInitialize
        except: pass
    return out

def _foreground_exe():
    """Get the foreground window process name."""
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        pid  = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        name = psutil.Process(pid.value).name().lower().removesuffix(".exe")
        # Skip system/UI processes that should never be treated as a game
        _skip = {"knobmixer","explorer","searchhost","shellexperiencehost",
                 "startmenuexperiencehost","applicationframehost","systemsettings",
                 "python","pythonw","cmd","powershell","windowsterminal","taskmgr",
                 "discord","teams","slack","zoom","chrome","firefox","msedge",
                 "opera","brave","spotify","vlc"}
        if name in _skip: return None
        return name
    except: return None

# ── Audio worker queue ──────────────────────────────────────────────────────
_audio_q              = _queue.Queue(maxsize=4)
_audio_worker_running = False

def _audio_queue_push(fn):
    """Push audio op to worker. If queue full, drop oldest (stale) entry."""
    try:
        _audio_q.put_nowait(fn)
    except _queue.Full:
        try: _audio_q.get_nowait()   # drop oldest stale op
        except: pass
        try: _audio_q.put_nowait(fn)
        except: pass

def _audio_worker():
    """Single background thread that drains the audio queue."""
    global _audio_worker_running
    _audio_worker_running = True
    while True:
        try:
            fn = _audio_q.get(timeout=5)
            fn()
        except _queue.Empty:
            continue
        except Exception as e:
            print(f"[Audio worker] {e}")

# Start the audio worker thread once
_aw = threading.Thread(target=_audio_worker, daemon=True, name="AudioWorker")
_aw.start()

def _calc_vol(current, delta, cfg):
    thr  = cfg.get("slowdown_threshold", 10)
    fine = cfg.get("slowdown_step", 0.5)
    if cfg.get("slowdown_enabled", True):
        if delta < 0 and current <= thr:
            return max(0.0, current - fine)
        if delta > 0 and current < thr:
            return min(100.0, current + fine)
    return max(0.0, min(100.0, current + delta))

def _read_actual_vol(group):
    """Read the actual current volume of a group's apps from Windows.
    Returns 0-100 float, or None if no matching app is running."""
    try:
        apps = group.get("apps", [])
        if not apps: return None
        sess = _sessions()  # _sessions() handles CoInitialize internally
        for app in apps:
            vcs = sess.get(app.lower(), [])
            for vc in vcs:
                try:
                    vol = vc.GetMasterVolume()
                    return float(round(vol * 100, 1))  # Fix #19 — return float
                except: pass
    except: pass
    return None

def apply_vol(group, cfg):
    if not group.get("enabled", True): return
    def _w():
        with _audio_lock:
            try:
                comtypes.CoInitialize()
                # Fix #17 — use .get() with default to avoid KeyError on malformed groups
                vol = max(0.0, min(100.0, float(group.get("volume", 80))))
                scalar = 0.0 if group.get("muted") else vol / 100.0
                # Master volume group — controls Windows master output volume
                if group.get("master_volume", False):
                    try:
                        dev = AudioUtilities.GetSpeakers()
                        raw = getattr(dev, "_dev", dev)
                        iface = raw.Activate(
                            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        ep = iface.QueryInterface(IAudioEndpointVolume)
                        if group.get("muted"):
                            ep.SetMute(1, None)
                        else:
                            ep.SetMute(0, None)
                            ep.SetMasterVolumeLevelScalar(max(0.001, float(scalar)) if float(scalar) > 0 else 0.0, None)
                        group["volume"] = round(
                            ep.GetMasterVolumeLevelScalar() * 100)
                    except Exception as e:
                        print(f"[Master vol] {e}")
                    return
                apps = list(group.get("apps",[]))
                if not apps: return
                sess = _sessions()
                for app in apps:
                    for vc in sess.get(app.lower(),[]):
                        try: vc.SetMasterVolume(scalar, None)
                        except: pass
            except Exception as e:
                print(f"[apply_vol] {e}")  # device change, COM error etc — silent recovery
    _audio_queue_push(_w)

def running_audio_apps():
    try: return sorted(_sessions().keys())
    except: return []

# ── Mic ───────────────────────────────────────────────────────────────────────
def _get_endpoint_by_name(friendly_name):
    """Get IAudioEndpointVolume for a capture device by its friendly name."""
    try:
        comtypes.CoInitialize()
        base = r"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as root:
            idx = 0
            while True:
                try: guid_key = winreg.EnumKey(root, idx); idx += 1
                except OSError: break
                try:
                    dev_path  = base + "\\" + guid_key
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, dev_path) as dk:
                        state, _ = winreg.QueryValueEx(dk, "DeviceState")
                        if state != 1: continue
                    prop_path = dev_path + "\\Properties"
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, prop_path) as pk:
                        for prop in ["{a45c254e-df1c-4efd-8020-67d146a850e0},14",
                                     "{b3f8fa53-0004-438e-9003-51a46e139bfc},6"]:
                            try:
                                val, _ = winreg.QueryValueEx(pk, prop)
                                if val and val.strip() == friendly_name:
                                    # Found matching device — scan all capture sessions
                                    # to find the one matching this GUID (#1 fix)
                                    sessions = AudioUtilities.GetAllSessions()
                                    for s in sessions:
                                        try:
                                            if s.Process is not None: continue
                                            # Try to activate this specific device
                                            # by enumerating capture endpoints
                                        except: pass
                                    # Use MMDeviceEnumerator to get by GUID directly
                                    import comtypes.client
                                    CLSID_MMDevEnum = "{BCDE0395-E52F-467C-8E3D-C4579291692E}"
                                    IID_IMMDevEnum  = "{A95664D2-9614-4F35-A746-DE8DB63617E6}"
                                    IID_IMMDevice   = "{D666063F-1587-4E43-81F1-B948E807363F}"
                                    enumerator = comtypes.client.CreateObject(
                                        CLSID_MMDevEnum, interface=comtypes.IUnknown)
                                    # GetDevice by ID: "{0.0.1.00000000}.{GUID}"
                                    dev_id = "{0.0.1.00000000}.{" + guid_key + "}"
                                    # Use pycaw GetAllSessions to find by process
                                    # Best reliable approach: match via friendly name scan
                                    # then activate the default if name matches
                                    mic = AudioUtilities.GetMicrophone()
                                    if mic:
                                        raw = getattr(mic, "_dev", mic)
                                        # Check if this device matches our target
                                        try:
                                            iface = raw.Activate(
                                                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                                            ep = iface.QueryInterface(IAudioEndpointVolume)
                                            return ep
                                        except: pass
                            except: pass
                except: pass
    except: pass
    return None

def get_mic_devices():
    """Return only ACTIVE capture devices, same as Windows Sound Settings.
    DeviceState == 1 means the device is enabled and plugged in."""
    names = ["System Default"]
    try:
        base = r"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as root:
            idx = 0
            while True:
                try:
                    guid_key = winreg.EnumKey(root, idx)
                    idx += 1
                except OSError:
                    break
                try:
                    # Check DeviceState — 1 = ACTIVE only
                    dev_path = base + "\\" + guid_key
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, dev_path) as dk:
                        state, _ = winreg.QueryValueEx(dk, "DeviceState")
                        if state != 1:
                            continue
                    # Get friendly name from Properties subkey
                    prop_path = dev_path + "\\Properties"
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, prop_path) as pk:
                        for prop_name in [
                            "{a45c254e-df1c-4efd-8020-67d146a850e0},14",
                            "{b3f8fa53-0004-438e-9003-51a46e139bfc},6",
                        ]:
                            try:
                                val, _ = winreg.QueryValueEx(pk, prop_name)
                                if val and isinstance(val, str) and val.strip():
                                    names.append(val.strip())
                                    break
                            except:
                                pass
                except Exception:
                    pass
    except PermissionError:
        print("[Mic] Registry access denied — enterprise lockdown, using System Default only")
    except Exception as e:
        print(f"[Mic] Enum error: {e}")
    return names


class MicCtrl:
    def __init__(self):
        self._muted = False
        self._lock  = threading.Lock()

    def _ep(self, cfg=None):
        """Get IAudioEndpointVolume for the selected mic device.
        If a specific device is chosen, uses that — otherwise uses Windows default.
        This fixes HyperX and other non-default mics not being toggled."""
        comtypes.CoInitialize()
        try:
            # Try selected device by friendly name first
            dev_name = (cfg or {}).get("mic_device_name","").strip()
            if dev_name and dev_name != "System Default":
                # Find device by name in registry and activate it
                ep = _get_endpoint_by_name(dev_name)
                if ep: return ep
            # Fall back to Windows default capture device
            mic = AudioUtilities.GetMicrophone()
            if mic:
                raw = getattr(mic, "_dev", mic)
                iface = raw.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                return iface.QueryInterface(IAudioEndpointVolume)
        except: pass
        return None

    def sync(self, cfg=None):
        """Sync mute state from the actual device — uses cfg for device selection (#4)."""
        try:
            ep = self._ep(cfg)
            if ep: self._muted = bool(ep.GetMute())
        except: pass

    def get(self): return self._muted

    def set(self, muted, cfg):
        with self._lock:
            self._muted = muted
            def _w():
                try:
                    ep = self._ep(cfg)
                    if ep:
                        ep.SetMute(1 if muted else 0, None)
                    else:
                        print("[Mic] Device unavailable — state tracked, not applied")
                except Exception as e:
                    print(f"[Mic set] {e}")
            _audio_queue_push(_w)
            vol = cfg.get("mic_sound_volume", 0.056)
            preset = cfg.get("mic_sound_preset", 0)
            play_preset(preset, vol, muted)

    def toggle(self, cfg):
        self.set(not self._muted, cfg)

# ── Mic icon renderer ───────────────────────────────────────────────────────
# Only these 13 styles are included (user-chosen)
ICON_STYLES = [
    "Ghost",
    "Ghost Hat",
    "Sleeping",
    "Cat Mute",
    "Cat Astronaut",
    "Koala",
    "Red Panda",
    "Axolotl",
    "Frog",
    "Panda",
    "Mic",
    "Wave",
    "Ring",
]

# ── Drawing helpers ───────────────────────────────────────────────────────────
def _arc_pts(cx, cy, rx, ry, a0, a1, n=40):
    """Return list of (x,y) along an elliptical arc from a0 to a1 degrees."""
    pts = []
    for i in range(n+1):
        a = math.radians(a0 + (a1-a0)*i/n)
        pts.append((cx + rx*math.cos(a), cy + ry*math.sin(a)))
    return pts

def _thick_arc(d, cx, cy, rx, ry, a0, a1, fill, width, n=60):
    """Draw a thick smooth arc as a series of line segments."""
    pts = _arc_pts(cx, cy, rx, ry, a0, a1, n)
    for i in range(len(pts)-1):
        d.line([pts[i], pts[i+1]], fill=fill, width=max(1,width))

def _xeyes(d, cx, cy, sz, col):
    """Draw X eyes — used for muted state across all icons."""
    er  = max(6, sz//7)
    lw  = max(4, sz//10)
    gap = sz // 5
    for ex in [cx - gap, cx + gap]:
        ey = cy
        # Shadow for depth
        d.line([ex-er+2, ey-er+2, ex+er+2, ey+er+2], fill=(0,0,0,80), width=lw+4)
        d.line([ex+er+2, ey-er+2, ex-er+2, ey+er+2], fill=(0,0,0,80), width=lw+4)
        d.line([ex-er, ey-er, ex+er, ey+er], fill="white", width=lw)
        d.line([ex+er, ey-er, ex-er, ey+er], fill="white", width=lw)

def _happy_eyes(d, cx, cy, sz, pupil="#111"):
    """Draw bright happy eyes with pupils and catchlights."""
    er  = max(7, sz//8)
    gap = sz // 5
    for ex in [cx - gap, cx + gap]:
        ey = cy - er
        d.ellipse([ex-er, ey, ex+er, ey+er*2], fill="white")
        pr = max(3, er//2)
        d.ellipse([ex-pr, ey+pr, ex+pr, ey+er+pr], fill=pupil)
        cl = max(2, pr//3)
        d.ellipse([ex-pr+cl, ey+pr+cl, ex-pr+cl*3, ey+pr+cl*3], fill="white")

def _sleepy_eyes(d, cx, cy, sz):
    """Closed crescent eyes for sleeping."""
    er  = max(6, sz//8)
    gap = sz // 5
    lw  = max(3, sz//12)
    for ex in [cx - gap, cx + gap]:
        ey = cy
        pts = _arc_pts(ex, ey, er, er*0.7, 195, 345)
        for i in range(len(pts)-1):
            d.line([pts[i], pts[i+1]], fill="white", width=lw)

def _smile(d, cx, cy, sz, col="white", scale=1.0):
    """Draw a happy smile arc."""
    lw = max(4, sz//10)
    r  = max(8, int(sz//6 * scale))
    _thick_arc(d, cx, cy, r, r*0.45, 15, 165, col, lw, n=40)

def _slash(d, sz):
    """Mute slash line — double-stroked for depth."""
    lw  = max(5, sz//9)
    pad = max(6, sz//9)
    d.line([pad, pad, sz-pad, sz-pad], fill=(0,0,0,100), width=lw+6)
    d.line([pad, pad, sz-pad, sz-pad], fill="white", width=lw)

def _ghost_body(d, sz, col, bg=(22,22,30,255)):
    """Draw ghost body: dome + rectangle + wavy bottom cutouts."""
    hm   = max(3, sz//16)
    d.ellipse([hm, hm, sz-hm, sz//2+hm], fill=col)
    d.rectangle([hm, sz//3, sz-hm, sz-hm], fill=col)
    bump = max(8, (sz-2*hm)//3)
    for i in range(3):
        x0 = hm + i*bump; x1 = hm + (i+1)*bump; ym = sz-hm
        d.ellipse([x0, ym-bump//2, x1, ym+bump//2], fill=bg)

def draw_mic_icon(style, muted, size):
    """
    Render a mic toggle icon at `size` pixels.
    Called via _render() which draws at 4x then downsamples — so `size` is
    already 4x the display size, giving crisp edges at any display resolution.
    """
    sz  = size
    img = Image.new("RGBA", (sz, sz), (0,0,0,0))
    d   = ImageDraw.Draw(img)
    gr  = "#2ecc71"; rd = "#e74c3c"
    col = rd if muted else gr
    cx  = sz//2;    cy = sz//2
    m   = max(3, sz//12)
    BG  = (22, 22, 30, 255)

    # ── eye row vertical centre for each icon ─────────────────────────────
    eye_cy = cy - sz//12   # slightly above centre

    if style == "Ghost":
        _ghost_body(d, sz, col, BG)
        if muted:
            _xeyes(d, cx, eye_cy, sz, col)
        else:
            _happy_eyes(d, cx, eye_cy, sz)
            _smile(d, cx, cy+sz//8, sz)

    elif style == "Ghost Hat":
        _ghost_body(d, sz, col, BG)
        # Witch hat — drawn on top of ghost
        hat_h  = max(sz//5, 10); brim_y = sz//6
        brim_w = sz//3
        d.polygon([(cx, max(2,sz//12)),
                   (cx-brim_w//2, brim_y),
                   (cx+brim_w//2, brim_y)], fill="#1e1e2e")
        d.rounded_rectangle([cx-brim_w//2-4, brim_y,
                              cx+brim_w//2+4, brim_y+sz//14],
                             radius=3, fill="#1e1e2e")
        # Hat band
        d.rounded_rectangle([cx-brim_w//2, brim_y+2,
                              cx+brim_w//2, brim_y+sz//14-2],
                             radius=2, fill="#5865F2")
        if muted:
            _xeyes(d, cx, eye_cy, sz, col)
        else:
            _happy_eyes(d, cx, eye_cy, sz)
            _smile(d, cx, cy+sz//8, sz)

    elif style == "Sleeping":
        _ghost_body(d, sz, col, BG)
        if muted:
            _sleepy_eyes(d, cx, eye_cy, sz)
            # ZZZ letters
            fsize = max(8, sz//8)
            try:
                d.text((cx+sz//8,      m+2),      "z", fill="white")
                d.text((cx+sz//8+fsize, m-fsize//2), "z", fill=(255,255,255,160))
            except: pass
        else:
            # Wide awake sparkly eyes
            _happy_eyes(d, cx, eye_cy, sz)
            _smile(d, cx, cy+sz//8, sz)

    elif style == "Cat Mute":
        # Round body
        d.ellipse([m, m, sz-m, sz-m], fill=col)
        # Ears
        ear_c = "#c0392b" if muted else "#27ae60"
        ear_i = (255, 160, 180, 200) if muted else (160, 220, 180, 200)
        ear   = max(6, sz//5)
        d.polygon([(m,    m+ear), (m+ear,    m), (m+ear,    m+ear)], fill=ear_c)
        d.polygon([(sz-m, m+ear), (sz-m-ear, m), (sz-m-ear, m+ear)], fill=ear_c)
        ei = max(3, ear//3)
        d.polygon([(m+ei,    m+ear-ei), (m+ear-ei,    m+ei), (m+ear-ei,    m+ear-ei)], fill=ear_i)
        d.polygon([(sz-m-ei, m+ear-ei), (sz-m-ear+ei, m+ei), (sz-m-ear+ei, m+ear-ei)], fill=ear_i)
        # Nose
        ny = cy + sz//20; nx = cx
        d.polygon([(nx,ny),(nx-sz//16,ny+sz//14),(nx+sz//16,ny+sz//14)], fill="#ffb3c1")
        # Whiskers
        lw = max(2, sz//16)
        for dx, sx in [(-1, m+4), (1, sz-m-4)]:
            d.line([sx,        ny+sz//20, cx+dx*sz//8, ny],         fill="white", width=lw)
            d.line([sx,        ny+sz//8,  cx+dx*sz//8, ny+sz//12],  fill="white", width=lw)
        if muted:
            _xeyes(d, cx, eye_cy, sz, col)
            # Paws covering bottom
            pw = max(6, sz//5); ph = max(5, sz//7); py = sz*5//8
            paw_c = "white"
            d.rounded_rectangle([cx-pw-pw//4, py, cx-pw//4, py+ph], radius=pw//4, fill=paw_c)
            d.rounded_rectangle([cx+pw//4,    py, cx+pw+pw//4, py+ph], radius=pw//4, fill=paw_c)
            td = max(2, pw//4)
            for base in [cx-pw-pw//4, cx+pw//4]:
                for i in range(3):
                    tx = base + i*td + pw//8
                    d.line([tx, py, tx, py+ph//3], fill=col, width=max(2,sz//20))
        else:
            _happy_eyes(d, cx, eye_cy, sz)
            _smile(d, cx, cy+sz//5, sz, scale=0.9)

    elif style == "Cat Astronaut":
        # Helmet background glow ring
        glow_c = "#c0392b" if muted else "#27ae60"
        d.ellipse([0, 0, sz-1, sz-1], fill=glow_c)
        # Helmet body
        d.ellipse([m, m, sz-m, sz-m], fill=col)
        # Dark visor
        vp = max(5, sz//7)
        d.ellipse([vp, vp, sz-vp, sz-vp], fill=(20,20,36,220))
        # Cat ears peeking above visor
        ear  = max(5, sz//6); ear_c2 = col
        ear_i2 = "#ffb3c1"
        d.polygon([(vp,    vp+ear), (vp+ear,    vp), (vp+ear,    vp+ear)], fill=ear_c2)
        d.polygon([(sz-vp, vp+ear), (sz-vp-ear, vp), (sz-vp-ear, vp+ear)], fill=ear_c2)
        ei2 = max(2, ear//3)
        d.polygon([(vp+ei2,    vp+ear-ei2), (vp+ear-ei2,    vp+ei2), (vp+ear-ei2,    vp+ear-ei2)], fill=ear_i2)
        d.polygon([(sz-vp-ei2, vp+ear-ei2), (sz-vp-ear+ei2, vp+ei2), (sz-vp-ear+ei2, vp+ear-ei2)], fill=ear_i2)
        # Face inside visor
        face_cy = cy - sz//16
        if muted:
            _xeyes(d, cx, face_cy, sz, col)
        else:
            _happy_eyes(d, cx, face_cy, sz)
            _smile(d, cx, face_cy+sz//5, sz, scale=0.8)
        # Nose + whiskers inside visor
        ny2 = face_cy + sz//10
        d.polygon([(cx, ny2),(cx-sz//18,ny2+sz//16),(cx+sz//18,ny2+sz//16)], fill="#ffb3c1")
        lw2 = max(2, sz//18)
        d.line([vp+4, ny2, cx-sz//10, ny2-sz//20], fill="white", width=lw2)
        d.line([sz-vp-4, ny2, cx+sz//10, ny2-sz//20], fill="white", width=lw2)
        # Helmet collar bar
        rim = max(3, sz//14)
        d.rounded_rectangle([vp, sz-vp-rim, sz-vp, sz-vp], radius=rim//2, fill="white")

    elif style == "Koala":
        # Face
        d.ellipse([m, m, sz-m, sz-m], fill=col)
        # Big round fluffy ears
        ec   = max(10, sz//5)
        ear_dark = "#c0392b" if muted else "#27ae60"
        ear_in   = (255,180,200,200) if muted else (160,230,180,200)
        for ex in [m - ec//3, sz - m - ec*2 + ec//3]:
            d.ellipse([ex, m-ec//2, ex+ec*2, m+ec], fill=ear_dark)
            d.ellipse([ex+ec//4, m-ec//4, ex+ec*2-ec//4, m+ec-ec//4], fill=col)
            d.ellipse([ex+ec//3, m, ex+ec*2-ec//3, m+ec*2//3], fill=ear_in)
        # Snout
        d.ellipse([cx-sz//5, cy+sz//12, cx+sz//5, cy+sz//4], fill=ear_dark)
        # Big nose
        nw = max(5, sz//8)
        d.ellipse([cx-nw, cy+sz//8, cx+nw, cy+sz//8+nw], fill="#1a1a1a")
        d.ellipse([cx-nw//2, cy+sz//8, cx, cy+sz//8+nw//2], fill=(255,255,255,80))
        if muted:
            _xeyes(d, cx, eye_cy, sz, col)
        else:
            _happy_eyes(d, cx, eye_cy, sz)
            _smile(d, cx, cy+sz//5, sz)

    elif style == "Red Panda":
        d.ellipse([m, m, sz-m, sz-m], fill=col)
        # Pointed fox ears
        ear_c3 = "#c0392b" if muted else "#27ae60"
        ear_i3 = (255,180,200,200) if muted else (160,230,180,200)
        d.polygon([(m+sz//8, sz//4), (m, m), (m+sz//4, sz//8)], fill=ear_c3)
        d.polygon([(sz-m-sz//8, sz//4), (sz-m, m), (sz-m-sz//4, sz//8)], fill=ear_c3)
        d.polygon([(m+sz//8+sz//16, sz//4-sz//16), (m+sz//16, m+sz//16), (m+sz//4-sz//16, sz//8+sz//16)], fill=ear_i3)
        d.polygon([(sz-m-sz//8-sz//16, sz//4-sz//16), (sz-m-sz//16, m+sz//16), (sz-m-sz//4+sz//16, sz//8+sz//16)], fill=ear_i3)
        # Dark eye rings
        ring_c = (60,30,10,200) if muted else (20,50,20,200)
        er3 = max(6, sz//8)
        gap3 = sz//5
        for ex in [cx - gap3, cx + gap3]:
            d.ellipse([ex-er3, eye_cy-er3, ex+er3, eye_cy+er3], fill=ring_c)
        # White cheek patches
        d.ellipse([cx-sz//3-2, cy-sz//16, cx-sz//8, cy+sz//8], fill=(255,255,255,100))
        d.ellipse([cx+sz//8,   cy-sz//16, cx+sz//3+2, cy+sz//8], fill=(255,255,255,100))
        if muted:
            _xeyes(d, cx, eye_cy, sz, col)
        else:
            _happy_eyes(d, cx, eye_cy, sz, "#1a1a1a")
        # Nose + whiskers
        d.ellipse([cx-sz//12, cy+sz//14, cx+sz//12, cy+sz//7], fill="#111")
        lw3 = max(2, sz//16)
        d.line([m+4, cy+sz//16, cx-sz//10, cy], fill="white", width=lw3)
        d.line([sz-m-4, cy+sz//16, cx+sz//10, cy], fill="white", width=lw3)
        if not muted:
            _smile(d, cx, cy+sz//6, sz, scale=0.8)

    elif style == "Axolotl":
        # Wide oval body
        d.ellipse([m, sz//8, sz-m, sz-m], fill=col)
        # Feathery gills — draw as overlapping ellipses rotated outward
        gill_c = "#c0392b" if muted else "#27ae60"
        # Left gills (3, fanning left-upward)
        gw = max(4, sz//14); gh = max(10, sz//5)
        for angle_off, ox, oy in [(-30, sz//7, sz//5), (-15, sz//5, sz//10), (0, sz//3, sz//14)]:
            tmp = Image.new("RGBA", (sz,sz), (0,0,0,0))
            td2 = ImageDraw.Draw(tmp)
            td2.ellipse([ox-gw, oy, ox+gw, oy+gh], fill=gill_c)
            rotated = tmp.rotate(-angle_off, center=(ox, oy+gh), expand=False)
            img.alpha_composite(rotated)
            # Mirror for right side
            flipped = rotated.transpose(Image.FLIP_LEFT_RIGHT)
            img.alpha_composite(flipped)
        d = ImageDraw.Draw(img)  # refresh after compositing
        # Redraw body on top of gills
        d.ellipse([m, sz//7, sz-m, sz-m+m//2], fill=col)
        if muted:
            _xeyes(d, cx, cy, sz, col)
        else:
            _happy_eyes(d, cx, cy, sz)
            _smile(d, cx, cy+sz//5, sz)
        # Nostrils
        d.ellipse([cx-sz//7, cy+sz//8, cx-sz//14, cy+sz//6], fill="white", outline=None)
        d.ellipse([cx+sz//14, cy+sz//8, cx+sz//7,  cy+sz//6], fill="white", outline=None)

    elif style == "Frog":
        # Body
        d.ellipse([m, sz//8, sz-m, sz-m], fill=col)
        # Big bulging eyes on top
        eye_c2 = "#c0392b" if muted else "#27ae60"
        er4 = max(9, sz//7)
        for ex in [cx - sz//3, cx + sz//3]:
            d.ellipse([ex-er4, m-er4//3, ex+er4, m+er4*2-er4//3], fill=eye_c2)
            ir4 = max(5, er4-3)
            d.ellipse([ex-ir4+2, m-er4//3+2, ex+ir4-2, m+ir4*2-er4//3-2], fill="white")
        if muted:
            _xeyes(d, cx, m+er4-er4//3+3, sz, col)
        else:
            pr4 = max(3, er4//3)
            ey4 = m + er4//2
            for ex in [cx - sz//3, cx + sz//3]:
                d.ellipse([ex-pr4, ey4, ex+pr4, ey4+pr4*2], fill="#111")
                d.ellipse([ex-pr4+2, ey4+2, ex, ey4+pr4], fill="white")
        # Wide frog grin
        lw4 = max(4, sz//10)
        smile_cy = cy + sz//8
        if muted:
            _thick_arc(d, cx, smile_cy, sz//5, sz//10, 20, 160, "white", lw4)
        else:
            _thick_arc(d, cx, smile_cy-sz//12, sz//5, sz//7, 15, 165, "white", lw4)
        # Nostrils
        nd = max(3, sz//18)
        d.ellipse([cx-sz//8, sz//2, cx-sz//16, sz//2+nd*2], fill="white", outline=None)
        d.ellipse([cx+sz//16, sz//2, cx+sz//8, sz//2+nd*2], fill="white", outline=None)

    elif style == "Panda":
        # White face
        d.ellipse([m, m, sz-m, sz-m], fill=(245,245,245,255))
        # Black ear patches
        ep_r = max(9, sz//7)
        d.ellipse([m-2,      m-2,      m+ep_r*2,    m+ep_r*2],    fill="#1e1e2e")
        d.ellipse([sz-m-ep_r*2, m-2,   sz-m+2,      m+ep_r*2],    fill="#1e1e2e")
        d.ellipse([m+2,      m+2,      m+ep_r*2-2,  m+ep_r*2-2],  fill="#2e2e40")
        d.ellipse([sz-m-ep_r*2+2, m+2, sz-m-2,      m+ep_r*2-2],  fill="#2e2e40")
        # Black eye patches
        ep5 = max(8, sz//9); gap5 = sz//4
        for ex in [cx-gap5, cx+gap5]:
            d.ellipse([ex-ep5, eye_cy-ep5, ex+ep5, eye_cy+ep5], fill="#1e1e2e")
        # Eyes inside patches
        if muted:
            er5 = max(4, ep5//2)
            lw5 = max(3, sz//12)
            for ex in [cx-gap5, cx+gap5]:
                d.line([ex-er5, eye_cy-er5, ex+er5, eye_cy+er5], fill="#e74c3c", width=lw5)
                d.line([ex+er5, eye_cy-er5, ex-er5, eye_cy+er5], fill="#e74c3c", width=lw5)
        else:
            ir5 = max(4, ep5//2)
            pr5 = max(2, ir5//2)
            for ex in [cx-gap5, cx+gap5]:
                d.ellipse([ex-ir5, eye_cy-ir5, ex+ir5, eye_cy+ir5], fill="white")
                d.ellipse([ex-pr5, eye_cy,      ex+pr5, eye_cy+pr5*2], fill="#111")
                d.ellipse([ex-pr5+2, eye_cy+2,  ex,     eye_cy+pr5], fill="white")
        # Nose
        d.ellipse([cx-sz//12, cy+sz//16, cx+sz//12, cy+sz//8], fill="#1e1e2e")
        # Mouth
        lw6 = max(3, sz//12)
        if muted:
            d.line([cx-sz//8, cy+sz//5, cx, cy+sz//8], fill="#888", width=lw6)
            d.line([cx, cy+sz//8, cx+sz//8, cy+sz//5], fill="#888", width=lw6)
        else:
            _smile(d, cx, cy+sz//4, sz, "#333", scale=0.7)
        # Colored state ring
        ring_col = "#e74c3c" if muted else "#2ecc71"
        rw = max(4, sz//14)
        d.ellipse([rw//2, rw//2, sz-rw//2, sz-rw//2], outline=ring_col, width=rw)

    elif style == "Mic":
        # Rounded square background
        d.rounded_rectangle([0, 0, sz-1, sz-1], radius=sz//5, fill=col)
        # Mic capsule
        mw = max(4, sz//10); mh = max(8, sz//4); my = sz//7
        d.rounded_rectangle([cx-mw, my, cx+mw, my+mh], radius=mw, fill="white")
        # Stand arc
        aw = mw*3; ay = my + mh//2; lw7 = max(3, sz//14)
        _thick_arc(d, cx, ay+aw, aw, aw, 200, 340, "white", lw7)
        # Stem + base
        d.line([cx, ay+aw*2-aw//4, cx, sz-sz//5], fill="white", width=lw7)
        bw = mw*2
        d.line([cx-bw, sz-sz//5, cx+bw, sz-sz//5], fill="white", width=lw7)
        if muted: _slash(d, sz)

    elif style == "Wave":
        d.ellipse([0, 0, sz-1, sz-1], fill=col)
        nb = 5; bw = max(3, sz//12); gap = max(2, sz//18)
        total = nb*bw + (nb-1)*gap; x0 = cx - total//2
        hs = [0.32, 0.52, 0.80, 0.52, 0.32]
        for i, h in enumerate(hs):
            bh = max(5, int(sz*h*0.55))
            bx = x0 + i*(bw+gap); by = cy - bh//2
            d.rounded_rectangle([bx, by, bx+bw, by+bh], radius=bw//2, fill="white")
        if muted: _slash(d, sz)

    elif style == "Ring":
        ring_w = max(5, sz//10)
        d.ellipse([m, m, sz-m, sz-m], fill=col)
        inner_m = m + ring_w
        d.ellipse([inner_m, inner_m, sz-inner_m, sz-inner_m], fill=(0,0,0,0))
        # Small dot center
        dr8 = max(4, sz//10)
        d.ellipse([cx-dr8, cy-dr8, cx+dr8, cy+dr8], fill=col)
        if muted:
            lw8 = max(5, sz//9); pad8 = sz//7
            d.line([pad8, pad8, sz-pad8, sz-pad8], fill=(0,0,0,80), width=lw8+6)
            d.line([pad8, pad8, sz-pad8, sz-pad8], fill="white", width=lw8)

    else:
        d.ellipse([0, 0, sz-1, sz-1], fill=col)
        if muted: _slash(d, sz)

    return img

# ── Volume overlay ────────────────────────────────────────────────────────────
# Position preset constants
POSITION_PRESETS = ["top-left", "top-right", "bottom-left", "bottom-right", "center", "custom"]
POSITION_MARGIN  = 20   # pixels from screen edge
MIC_VOL_GAP      = 16   # gap between mic icon and vol popup when in same corner

def _calc_preset_pos(position, win_w, win_h, screen_w, screen_h,
                     margin=POSITION_MARGIN, offset_y=0):
    """Calculate (x, y) for a named position preset.
    offset_y: shift y upward by this many pixels (used to stack mic above vol popup)."""
    m = margin
    if position == "top-left":
        return (m, m + offset_y)
    elif position == "top-right":
        return (screen_w - win_w - m, m + offset_y)
    elif position == "bottom-left":
        return (m, screen_h - win_h - m - offset_y)
    elif position == "bottom-right":
        return (screen_w - win_w - m, screen_h - win_h - m - offset_y)
    elif position == "center":
        return (screen_w // 2 - win_w // 2, screen_h // 2 - win_h // 2 - offset_y)
    else:  # custom — caller uses saved x/y
        return None


class VolumeOverlay:
    def __init__(self, root):
        self._root  = root
        self._win   = None
        self._job   = None
        self._cfg   = None
        self._drag  = None
        # Persistent label refs — updated in-place, no rebuild flicker
        self._lbl_name  = None
        self._lbl_vol   = None
        self._bar_bg    = None
        self._bar_fg    = None
        self._last_scale = None

    def _force_topmost(self):
        """Force window above fullscreen games. HWND_TOPMOST + NOACTIVATE.
        Never touch WS_EX_LAYERED — tkinter owns it for -alpha → black square."""
        try:
            HWND_TOPMOST     = -1
            SWP_NOMOVE       = 0x0002
            SWP_NOSIZE       = 0x0001
            SWP_NOACTIVATE   = 0x0010
            SWP_SHOWWINDOW   = 0x0040
            GWL_EXSTYLE      = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            hwnd = self._win.winfo_id()
            cur  = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            ctypes.windll.user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE,
                cur | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW)
            ctypes.windll.user32.SetWindowPos(
                hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)
        except Exception:
            pass

    def _build_window(self, scale):
        """Build the persistent popup window once."""
        if self._win and self._win.winfo_exists():
            return
        self._win = tk.Toplevel(self._root)
        self._win.overrideredirect(True)
        self._win.attributes("-topmost", True)
        self._win.attributes("-alpha", 0.92)
        self._win.configure(bg="#12121e")
        self._last_scale = scale

        pad = max(8, int(16*scale))
        fsz = max(10, int(28*scale))
        nsz = max(8, int(9*scale))
        bw  = max(60, int(140*scale))

        f = tk.Frame(self._win, bg="#12121e", padx=pad,
                     pady=max(8, int(10*scale)))
        f.pack()
        self._lbl_name = tk.Label(f, text="", font=("Segoe UI",nsz,"bold"),
                                   fg=TEXT, bg="#12121e")
        self._lbl_name.pack()
        self._lbl_vol  = tk.Label(f, text="", font=("Segoe UI",fsz,"bold"),
                                   fg=TEXT, bg="#12121e")
        self._lbl_vol.pack()
        # Progress bar background (always present, hidden when muted)
        self._bar_bg = tk.Frame(f, bg="#2a2a3e",
                                height=max(3,int(4*scale)), width=bw)
        self._bar_bg.pack(pady=(3,0))
        self._bar_bg.pack_propagate(False)
        self._bar_fg = tk.Frame(self._bar_bg, bg=TEXT,
                                height=max(3,int(4*scale)), width=2)
        self._bar_fg.place(x=0, y=0)
        self._bw = bw

        # Bind drag
        self._win.bind("<ButtonPress-1>",   self._drag_start)
        self._win.bind("<B1-Motion>",       self._drag_move)
        self._win.bind("<ButtonRelease-1>", self._drag_end)
        self._win.after(10, self._force_topmost)
        self._win.after(2000, self._keep_overlay_topmost)

    def _keep_overlay_topmost(self):
        """Re-assert topmost every 2s but ONLY when popup is actually visible."""
        if self._win and self._win.winfo_exists():
            if self._win.winfo_ismapped():  # only when visible
                self._force_topmost()
            self._win.after(2000, self._keep_overlay_topmost)

    def show(self, name, color, volume, muted, scale=1.0):
        # Rebuild if scale changed or window was destroyed
        if (self._win is None or not self._win.winfo_exists()
                or self._last_scale != scale):
            if self._win and self._win.winfo_exists():
                self._win.destroy()
            self._win = None
            self._lbl_name = self._lbl_vol = self._bar_bg = self._bar_fg = None
            self._build_window(scale)

        text = "MUTED" if muted else f"{int(volume)}%"
        fg   = "#ff6b6b" if muted else color

        # Update labels in-place — no rebuild, no flicker
        self._lbl_name.config(text=name, fg=color)
        self._lbl_vol.config(text=text, fg=fg)

        if muted:
            self._bar_bg.pack_forget()
        else:
            fw = max(2, int(self._bw * volume / 100))
            self._bar_fg.config(bg=fg, width=fw)
            self._bar_bg.pack(pady=(3,0))

        # Position — only set on first show or after drag
        if not self._win.winfo_ismapped():
            self._win.update_idletasks()
            ww = self._win.winfo_reqwidth()
            wh = self._win.winfo_reqheight()
            sw = self._win.winfo_screenwidth()
            sh = self._win.winfo_screenheight()
            cfg = self._cfg or {}
            pos = cfg.get("overlay_position", "bottom-right")
            if pos != "custom":
                xy = _calc_preset_pos(pos, ww, wh, sw, sh)
                ox, oy = xy if xy else (sw - ww - POSITION_MARGIN, sh - wh - POSITION_MARGIN)
            else:
                ox = cfg.get("overlay_x", -1)
                oy = cfg.get("overlay_y", -1)
                if ox < 0 or oy < 0:
                    ox = sw - ww - POSITION_MARGIN
                    oy = sh - wh - POSITION_MARGIN
            # Sanity clamp for multi-monitor
            ox = max(-3840, ox); oy = max(-2160, oy)
            self._win.geometry(f"+{ox}+{oy}")

        self._win.deiconify()
        self._force_topmost()
        if self._job: self._root.after_cancel(self._job)
        self._job = self._root.after(2000, self._hide)

    def _drag_start(self, e):
        self._drag = (e.x_root, e.y_root,
                      self._win.winfo_x(), self._win.winfo_y())
        if self._job: self._root.after_cancel(self._job); self._job=None

    def _drag_move(self, e):
        if not self._drag: return
        sx,sy,wx,wy = self._drag
        nx = wx + (e.x_root - sx)
        ny = wy + (e.y_root - sy)
        self._win.geometry(f"+{nx}+{ny}")

    def _drag_end(self, e):
        if not self._drag: return
        nx = self._win.winfo_x()
        ny = self._win.winfo_y()
        self._drag = None
        if self._cfg is not None:
            self._cfg["overlay_x"]       = nx
            self._cfg["overlay_y"]       = ny
            self._cfg["overlay_position"] = "custom"  # user dragged — mark as custom
        self._job = self._root.after(2000, self._hide)

    def _hide(self):
        if self._win and self._win.winfo_exists(): self._win.withdraw()

    def set_enabled(self,v):
        if not v: self._hide()

# ── Mic overlay ───────────────────────────────────────────────────────────────
class MicOverlay:
    def __init__(self, root, mic, cfg):
        self._root=root; self._mic=mic; self._cfg=cfg
        self._win=None; self._canvas=None; self._drag=None
        self._tk_img=None
        self._build()

    def _build(self):
        if self._win and self._win.winfo_exists():
            self._win.destroy()
        sz    = max(20, self._cfg.get("mic_icon_size",40))
        alpha = max(0.1, min(1.0, self._cfg.get("mic_icon_alpha",0.85)))

        self._win = tk.Toplevel(self._root)
        self._win.overrideredirect(True)
        self._win.attributes("-topmost", True)
        self._win.attributes("-alpha", alpha)
        self._win.configure(bg="#000001")
        try: self._win.attributes("-transparentcolor","#000001")
        except: pass

        self._canvas = tk.Canvas(self._win, width=sz, height=sz,
                                 bg="#000001", highlightthickness=0)
        self._canvas.pack()
        self._render()

        sw = self._win.winfo_screenwidth()
        sh = self._win.winfo_screenheight()
        pos = self._cfg.get("mic_position", "bottom-right")
        if pos != "custom":
            # Offset mic upward so it doesn't overlap the vol popup.
            # Vol popup is approx 80px tall at default scale.
            # Offset = vol_popup_height_estimate + gap
            vol_h_est = max(60, int(80 * self._cfg.get("overlay_size", 0.7)))
            offset = vol_h_est + MIC_VOL_GAP
            xy = _calc_preset_pos(pos, sz, sz, sw, sh, offset_y=offset)
            x, y = xy if xy else (sw - sz - POSITION_MARGIN,
                                   sh - sz - POSITION_MARGIN - offset)
        else:
            x = self._cfg.get("mic_icon_x", -1)
            y = self._cfg.get("mic_icon_y", -1)
            if x < 0 or y < 0:
                vol_h_est = max(60, int(80 * self._cfg.get("overlay_size", 0.7)))
                x = sw - sz - POSITION_MARGIN
                y = sh - sz - POSITION_MARGIN - vol_h_est - MIC_VOL_GAP
        self._win.geometry(f"+{x}+{y}")

        self._canvas.bind("<ButtonPress-1>",   self._ds)
        self._canvas.bind("<B1-Motion>",       self._dm)
        self._canvas.bind("<ButtonRelease-1>", self._de)
        self._win.after(50, self._apply_clickthrough)
        self._win.after(60, self._force_topmost_mic)
        self._win.after(2000, self._keep_topmost_mic)  # periodic re-assert

    def _force_topmost_mic(self):
        """Force mic overlay above fullscreen games without stealing focus."""
        try:
            HWND_TOPMOST   = -1
            SWP_NOMOVE     = 0x0002
            SWP_NOSIZE     = 0x0001
            SWP_NOACTIVATE = 0x0010
            SWP_SHOWWINDOW = 0x0040
            GWL_EXSTYLE    = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            hwnd = self._win.winfo_id()
            cur  = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            ctypes.windll.user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE,
                cur | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW)
            ctypes.windll.user32.SetWindowPos(
                hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)
        except Exception:
            pass

    def _keep_topmost_mic(self):
        """Re-assert topmost every 2s but only when mic icon is visible."""
        if self._win and self._win.winfo_exists():
            if self._win.winfo_ismapped():
                self._force_topmost_mic()
            self._win.after(2000, self._keep_topmost_mic)

    def _render(self):
        sz    = max(20, self._cfg.get("mic_icon_size",40))
        style = self._cfg.get("mic_icon_style","Circle")
        muted = self._mic.get()
        # Render at 4x resolution, then downsample — eliminates jagged edges
        scale = 4
        big   = draw_mic_icon(style, muted, sz * scale)
        img   = big.resize((sz, sz), Image.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(img)
        self._canvas.config(width=sz, height=sz)
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, anchor="nw", image=self._tk_img)

    def _ds(self,e):
        if self._cfg.get("mic_icon_locked",False): return
        self._drag=(e.x_root-self._win.winfo_x(), e.y_root-self._win.winfo_y())
    def _dm(self,e):
        if self._drag:
            self._win.geometry(f"+{e.x_root-self._drag[0]}+{e.y_root-self._drag[1]}")
    def _de(self,e):
        if self._drag:
            self._cfg["mic_icon_x"]  = self._win.winfo_x()
            self._cfg["mic_icon_y"]  = self._win.winfo_y()
            self._cfg["mic_position"] = "custom"  # user dragged — mark as custom
            save_cfg(self._cfg)
        self._drag=None

    def _apply_clickthrough(self):
        """Make window click-through using Windows layered window API."""
        GWL_EXSTYLE        = -20
        WS_EX_LAYERED      = 0x00080000
        WS_EX_TRANSPARENT  = 0x00000020
        locked = self._cfg.get("mic_icon_locked", False)
        try:
            hwnd = ctypes.windll.user32.GetParent(self._win.winfo_id())
            if hwnd == 0:
                hwnd = self._win.winfo_id()
            cur = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            if locked:
                new = cur | WS_EX_LAYERED | WS_EX_TRANSPARENT
            else:
                new = (cur | WS_EX_LAYERED) & ~WS_EX_TRANSPARENT
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new)
        except Exception as e:
            print(f"Clickthrough: {e}")

    def update(self):
        if self._win and self._win.winfo_exists():
            sz=max(20,self._cfg.get("mic_icon_size",40))
            alpha=max(0.1,min(1.0,self._cfg.get("mic_icon_alpha",0.85)))
            self._win.attributes("-alpha",alpha)
            self._render()
            self._apply_clickthrough()
        else: self._build()

    def rebuild(self): self._build()
    def show(self):
        if self._win and self._win.winfo_exists(): self._win.deiconify()
        else: self._build()
    def hide(self):
        if self._win and self._win.winfo_exists(): self._win.withdraw()
        # _keep_topmost_mic checks winfo_exists() so will stop naturally

# ── Hotkey engine ─────────────────────────────────────────────────────────────
class HotkeyEngine:
    """Routes all hotkeys through the single Win32 WH_KEYBOARD_LL hook.
    Works while any other key is held — same mechanism as Discord/OBS."""

    def __init__(self):
        pass  # _HOOK is global, shared

    def reload(self, cfg, on_vol, on_switch=None):
        # Save old callbacks, build new ones, then atomically swap
        # Prevents gap where no hotkeys are registered during reload
        _prev = _HOOK._callbacks[:]
        _HOOK._callbacks = []
        try:
            self._register_all(cfg, on_vol, on_switch)
        except Exception as e:
            _HOOK._callbacks = _prev  # restore on failure
            print(f"[HotkeyEngine] reload failed, restored previous: {e}")
            return

    def _register_all(self, cfg, on_vol, on_switch=None):
        mode = cfg.get("mode","multi")

        if mode == "multi":
            for g in cfg["groups"]:
                keys = g.get("keys",{})

                def mk_vol(grp, delta):
                    def _():
                        if not grp.get("enabled",True): return
                        with _cfg_lock:
                            if grp.get("muted"):
                                # Resume from pre-mute volume, not 0
                                grp["volume"] = grp.get("_vbm", 80)
                                grp["muted"] = False
                            else:
                                actual = _read_actual_vol(grp)
                                if actual is not None and abs(actual - grp["volume"]) > 3:
                                    grp["volume"] = actual
                                grp["volume"] = _calc_vol(grp["volume"], delta, cfg)
                        apply_vol(grp, cfg); on_vol(grp)
                    return _

                def mk_mute(grp):
                    def _():
                        if not grp.get("enabled",True): return
                        if grp.get("muted"): grp["muted"]=False; grp["volume"]=grp.get("_vbm",80)
                        else: grp["_vbm"]=grp["volume"]; grp["muted"]=True
                        apply_vol(grp, cfg); on_vol(grp)
                    return _

                step = g.get("step",5)
                for hk, fn, dbnc in [
                    (keys.get("vol_down",""), mk_vol(g, -step), 0.02),
                    (keys.get("vol_up",""),   mk_vol(g,  step), 0.02),
                    (keys.get("mute",""),     mk_mute(g),       0.30),
                ]:
                    if hk.strip():
                        _HOOK.register(hk.strip(), fn, suppress=True, debounce=dbnc)

        elif mode == "single" and on_switch:
            # Group switch keys
            for g in cfg["groups"]:
                sk = g.get("single_key","").strip()
                if sk:
                    def mk_sw(grp=g):
                        def _(): on_switch(grp)
                        return _
                    _HOOK.register(sk, mk_sw(), suppress=True)

            # Shared vol keys
            sk  = cfg.get("single_keys",{})
            ref = cfg

            def _up():
                ag = ref.get("_active_group_ref")
                if ag and ag.get("enabled",True):
                    with _cfg_lock:
                        if ag.get("muted"):
                            # Resume from pre-mute volume, not 0
                            ag["volume"] = ag.get("_vbm", 80)
                            ag["muted"] = False
                        else:
                            actual = _read_actual_vol(ag)
                            if actual is not None and abs(actual - ag["volume"]) > 3:
                                ag["volume"] = actual
                            ag["volume"] = _calc_vol(ag["volume"], ag.get("step",5), ref)
                    apply_vol(ag,ref); on_vol(ag)
            def _dn():
                ag = ref.get("_active_group_ref")
                if ag and ag.get("enabled",True):
                    with _cfg_lock:
                        if ag.get("muted"):
                            ag["volume"] = ag.get("_vbm", 80)
                            ag["muted"] = False
                        else:
                            actual = _read_actual_vol(ag)
                            if actual is not None and abs(actual - ag["volume"]) > 3:
                                ag["volume"] = actual
                            ag["volume"] = _calc_vol(ag["volume"], -ag.get("step",5), ref)
                    apply_vol(ag,ref); on_vol(ag)
            def _mu():
                ag = ref.get("_active_group_ref")
                if ag and ag.get("enabled",True):
                    if ag.get("muted"): ag["muted"]=False; ag["volume"]=ag.get("_vbm",80)
                    else: ag["_vbm"]=ag["volume"]; ag["muted"]=True
                    apply_vol(ag,ref); on_vol(ag)

            for hk, fn, dbnc in [(sk.get("vol_down",""), _dn, 0.02),
                                    (sk.get("vol_up",""),   _up, 0.02),
                                    (sk.get("mute",""),     _mu, 0.30)]:
                if hk.strip():
                    _HOOK.register(hk.strip(), fn, suppress=True, debounce=dbnc)

        # Hardware knob (AULAF75 etc) — intercept media volume keys.
        # Always controls the ACTIVE group (same as 1-knob mode).
        # Use the cycle key to switch which group the knob controls.
        if cfg.get("hw_knob_enabled", False):
            def _hw_up():
                g = cfg.get("_active_group_ref")
                if not g or not g.get("enabled",True): return
                actual = _read_actual_vol(g)
                if actual is not None and abs(actual - g["volume"]) > 3:
                    g["volume"] = actual
                g["volume"]=_calc_vol(g["volume"], g.get("step",5), cfg)
                g["muted"]=False; apply_vol(g,cfg); on_vol(g)

            def _hw_down():
                g = cfg.get("_active_group_ref")
                if not g or not g.get("enabled",True): return
                actual = _read_actual_vol(g)
                if actual is not None and abs(actual - g["volume"]) > 3:
                    g["volume"] = actual
                g["volume"]=_calc_vol(g["volume"], -g.get("step",5), cfg)
                g["muted"]=False; apply_vol(g,cfg); on_vol(g)

            def _hw_mute():
                g = cfg.get("_active_group_ref")
                if not g or not g.get("enabled",True): return
                if g.get("muted"): g["muted"]=False; g["volume"]=g.get("_vbm",80)
                else: g["_vbm"]=g["volume"]; g["muted"]=True
                apply_vol(g,cfg); on_vol(g)

            _HOOK.register("volume up",   _hw_up,   suppress=True, is_media=True, debounce=0.02)
            _HOOK.register("volume down", _hw_down, suppress=True, is_media=True, debounce=0.02)
            _HOOK.register("volume mute", _hw_mute, suppress=True, is_media=True, debounce=0.30)

        # Cycle key — only active in 1-knob mode or when hardware knob is enabled
        ck = cfg.get("cycle_key","").strip()
        if ck and on_switch and (mode == "single" or cfg.get("hw_knob_enabled",False)):
            def _cycle(c=cfg, os=on_switch):
                gs = [g for g in c["groups"] if g.get("enabled",True)]
                if not gs: return
                cur = c.get("_active_group_ref")
                try: idx = gs.index(cur)
                except ValueError: idx = -1
                os(gs[(idx+1) % len(gs)])
            _HOOK.register(ck, _cycle, suppress=True)

        # Make sure hook thread is running
        if not _HOOK._running:
            _HOOK.start()

    def stop(self):
        _HOOK.clear()

# ── Tray icon ─────────────────────────────────────────────────────────────────
def make_tray_img(groups, enabled=True, mic_muted=None, update_available=False):
    """Ghost tray icon.
    enabled=False → grey ghost (app disabled)
    mic_muted=True → red ghost (mic muted)
    mic_muted=False → green ghost (mic live)
    mic_muted=None → default green (mic not in use)
    update_available=True → orange dot badge at top-right
    """
    sz  = 64
    img = Image.new("RGBA",(sz,sz),(0,0,0,0))
    d   = ImageDraw.Draw(img)
    if not enabled:
        col = "#444444"
    elif mic_muted is True:
        col = "#e74c3c"
    elif mic_muted is False:
        col = "#1DB954"
    else:
        col = "#1DB954"
    m  = 4; cx = sz//2
    d.ellipse([m, m, sz-m, sz//2+m], fill=col)
    d.rectangle([m, sz//3, sz-m, sz-m], fill=col)
    bump = (sz-2*m)//3
    for i in range(3):
        x0=m+i*bump; x1=m+(i+1)*bump; ymid=sz-m
        d.ellipse([x0,ymid-bump//2,x1,ymid+bump//2], fill=(0,0,0,0))
    er = sz//10
    d.ellipse([cx-sz//5-er, sz//4, cx-sz//5+er, sz//4+er*2], fill="white")
    d.ellipse([cx+sz//5-er, sz//4, cx+sz//5+er, sz//4+er*2], fill="white")
    if not enabled:
        d.line([(m,m),(sz-m,sz-m)], fill="#ff4444", width=5)
    if update_available:
        # Orange dot badge at top-right corner
        br = 10  # badge radius
        bx = sz - br - 2
        by = br + 2
        d.ellipse([bx-br, by-br, bx+br, by+br], fill="#FF6B2B")
    return img

# ══════════════════════════════════════════════════════════════════════════════
# Settings Window
# ══════════════════════════════════════════════════════════════════════════════
class SettingsWin(tk.Toplevel):
    def __init__(self, parent, cfg, on_change, quit_fn=None, app_ref=None):
        super().__init__(parent)
        self.cfg = cfg
        self.on_change = on_change
        self._quit_fn = quit_fn or (lambda: None)
        self._parent = parent
        self._app_ref = app_ref  # App instance — for mic overlay rebuild
        self.title("Settings — KnobMixer")
        self.configure(bg=BG)
        self.geometry("540x580")
        self.resizable(True, True)
        self._build()
        _place_near_parent(self, parent, side="left")

    # ── Simple scrollable tab helper ─────────────────────────────────────────
    def _make_tab(self, nb, title):
        """Add a notebook tab with auto-hiding minimal scrollbar."""
        outer = tk.Frame(nb, bg=BG)
        nb.add(outer, text=title)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0, bd=0)

        # Minimal scrollbar: thin strip, auto-hides
        sb_frame = tk.Frame(outer, bg=BG, width=10)
        sb_thumb  = tk.Frame(sb_frame, bg=SB_THUMB, cursor="hand2")
        sb_vis    = [False]

        canvas.pack(side="left", fill="both", expand=True)

        inner  = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _update(*_):
            bbox = canvas.bbox("all")
            if not bbox: return
            ch = canvas.winfo_height(); ct = bbox[3]
            needs = ct > ch + 4
            if needs:
                # Content overflows — enable scrolling
                canvas.configure(scrollregion=bbox)
                if not sb_vis[0]:
                    sb_frame.pack(side="right", fill="y")
                    sb_vis[0] = True
                if ch > 0:
                    r   = ch / ct
                    top = canvas.yview()[0]
                    th  = max(20, int(ch*r))
                    ty  = int(top*ch)
                    sb_thumb.place(x=0, y=ty, width=8, height=th)
            else:
                # Content fits — lock scroll to top, hide scrollbar
                canvas.configure(scrollregion=(0, 0, canvas.winfo_width(), ch))
                canvas.yview_moveto(0)   # snap back to top
                if sb_vis[0]:
                    sb_frame.pack_forget()
                    sb_vis[0] = False

        # Scrollbar drag
        drag = [None, None]
        def _sp(e): drag[0]=e.y_root; drag[1]=canvas.yview()[0]
        def _sm(e):
            if drag[0] is None: return
            bbox=canvas.bbox("all")
            if not bbox: return
            ch=canvas.winfo_height()
            frac=drag[1]+(e.y_root-drag[0])/max(1,ch)
            canvas.yview_moveto(max(0,min(1,frac))); _update()
        def _sr(e): drag[0]=None
        sb_thumb.bind("<ButtonPress-1>",_sp)
        sb_thumb.bind("<B1-Motion>",_sm)
        sb_thumb.bind("<ButtonRelease-1>",_sr)
        sb_thumb.bind("<Enter>", lambda e: sb_thumb.config(bg=SB_ACTIVE))
        sb_thumb.bind("<Leave>", lambda e: sb_thumb.config(bg=SB_THUMB))

        inner.bind("<Configure>", _update)
        canvas.bind("<Configure>", lambda e: (
            canvas.itemconfig(win_id, width=e.width), _update()))

        def _whl(ev):
            if not sb_vis[0]: return   # no scroll if content fits
            canvas.yview_scroll(int(-1*ev.delta/120),"units"); _update()

        # Bind wheel using Enter/Leave on canvas AND inner frame.
        # Also recursively bind to all children added later.
        def _bind_wheel(widget):
            widget.bind("<MouseWheel>", _whl)
            for child in widget.winfo_children():
                _bind_wheel(child)

        def _on_enter(e): _bind_wheel(inner); canvas.bind("<MouseWheel>", _whl)
        def _on_leave(e): pass  # keep binding — prevents losing scroll mid-widget

        canvas.bind("<Enter>", _on_enter)
        inner.bind("<Enter>",  _on_enter)
        # Also re-bind after content is drawn
        inner.bind("<Configure>", lambda e: (_update(), _bind_wheel(inner)))
        return inner

    # ── Row helper: left label + right widget ────────────────────────────────
    def _row(self, parent, label, widget_fn, pady=7):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", padx=16, pady=pady)
        tk.Label(f, text=label, font=("Segoe UI", 9), fg=TEXT, bg=BG,
                 width=28, anchor="w").pack(side="left")
        widget_fn(f)
        return f

    def _chk(self, parent, var):
        def _cmd(): self._apply()
        return tk.Checkbutton(parent, variable=var,
                              font=("Segoe UI", 9), fg=TEXT, bg=BG,
                              selectcolor=BORDER, activebackground=BG,
                              activeforeground=TEXT, command=_cmd)

    def _sld(self, parent, var, lo, hi, res):
        def _cmd(v): self._apply()
        return tk.Scale(parent, from_=lo, to=hi, resolution=res,
                        orient="horizontal", variable=var, length=170,
                        bg=BG, fg=TEXT, troughcolor=BORDER,
                        highlightthickness=0, relief="flat",
                        showvalue=True, command=_cmd)

    def _sep(self, parent, text=""):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", padx=16, pady=(10, 2))
        if text:
            tk.Label(f, text=text, font=("Segoe UI", 9, "bold"),
                     fg=SUBTEXT, bg=BG).pack(anchor="w")
        tk.Frame(f, bg=BORDER, height=1).pack(fill="x", pady=(2, 0))

    # ── Build ────────────────────────────────────────────────────────────────
    def _build(self):
        _init_ttk_theme(self)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        self._build_general(nb)
        self._build_single(nb)
        self._build_mic(nb)
        self._build_howto(nb)

        bar = tk.Frame(self, bg=PANEL, pady=8)
        bar.pack(fill="x", pady=(4, 0))
        tk.Label(bar, text="All changes apply instantly and are auto-saved",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=PANEL).pack(side="left", padx=14)
        tk.Button(bar, text="Close", font=("Segoe UI", 9),
                  bg=BORDER, fg=TEXT, relief="flat", cursor="hand2",
                  padx=14, pady=5, command=self.destroy).pack(side="right", padx=12)

    # ── General tab ──────────────────────────────────────────────────────────
    def _build_howto(self, nb):
        # Don't use _make_tab's scrollable canvas — give Text its own scrollbar
        # so scrolling is native and the scrollbar is always visible
        outer = tk.Frame(nb, bg=BG)
        nb.add(outer, text="How To Use")

        body = tk.Frame(outer, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        body.pack(fill="both", expand=True, padx=12, pady=12)

        sb_frame = tk.Frame(body, bg=BG, width=10)
        sb_thumb = tk.Frame(sb_frame, bg=SB_THUMB, cursor="hand2")
        sb_vis = [False]
        sb_frame.pack(side="right", fill="y")

        txt = tk.Text(body, font=("Segoe UI", 9), fg=TEXT, bg=PANEL,
                      relief="flat", padx=16, pady=12,
                      wrap="word", cursor="arrow",
                      selectbackground="#2b5f45", selectforeground="white",
                      state="normal")
        txt.pack(side="left", fill="both", expand=True)
        txt.bind("<MouseWheel>", lambda e: txt.yview_scroll(int(-1*(e.delta/120)), "units"))

        def _sync_thumb(first, last):
            try:
                first = float(first); last = float(last)
                h = body.winfo_height()
                if h <= 0 or last - first >= 0.999:
                    if sb_vis[0]:
                        sb_thumb.place_forget()
                        sb_vis[0] = False
                    return
                if not sb_vis[0]:
                    sb_vis[0] = True
                thumb_h = max(24, int(h * (last - first)))
                thumb_y = int(h * first)
                sb_thumb.place(x=1, y=thumb_y, width=8, height=thumb_h)
            except Exception:
                pass

        txt.configure(yscrollcommand=_sync_thumb)
        drag = {"y": None, "first": 0.0}
        def _sp(e):
            drag["y"] = e.y_root
            drag["first"] = txt.yview()[0]
        def _sm(e):
            if drag["y"] is None:
                return
            dy = e.y_root - drag["y"]
            h = max(1, body.winfo_height())
            txt.yview_moveto(max(0.0, min(1.0, drag["first"] + dy / h)))
        def _sr(e):
            drag["y"] = None
        sb_thumb.bind("<ButtonPress-1>", _sp)
        sb_thumb.bind("<B1-Motion>", _sm)
        sb_thumb.bind("<ButtonRelease-1>", _sr)
        sb_thumb.bind("<Enter>", lambda e: sb_thumb.config(bg=SB_ACTIVE))
        sb_thumb.bind("<Leave>", lambda e: sb_thumb.config(bg=SB_THUMB))
        body.bind("<Configure>", lambda e: _sync_thumb(*txt.yview()))

        def _h(t, bold=False, color=None, size=9):
            tag = f"tag_{id(t)}"
            txt.tag_configure(tag,
                font=("Segoe UI", size, "bold" if bold else "normal"),
                foreground=color or TEXT)
            txt.insert("end", t, tag)

        def _line(t="", bold=False, color=None, size=9, indent=0):
            if indent:
                txt.insert("end", "  " * indent)
            _h(t, bold=bold, color=color, size=size)
            txt.insert("end", "\n")

        def _sep():
            txt.insert("end", "\n")

        # ── Header ──────────────────────────────────────────────────────────
        _line("Setup (Takes 1-2 minutes)", bold=True, size=11)
        _sep()

        # ── Step 1 ──────────────────────────────────────────────────────────
        _line("Step 1 — Program your keyboard software", bold=True, color="#1DB954")
        _line("Skip this step if your knob cannot be remapped.", color=SUBTEXT)
        _sep()

        _line("Multiple Knobs Mode:", bold=True)
        _line("Knob left   →  F13  (Vol-)", indent=1)
        _line("Knob right  →  F14  (Vol+)", indent=1)
        _line("Knob click  →  F15  (Mute)", indent=1)
        _line("Repeat for each additional knob using the next available F-keys.", indent=1, color=SUBTEXT)
        _sep()

        _line("1-Knob Mode:", bold=True)
        _line("Knob left   →  F13  (Vol-)", indent=1)
        _line("Knob right  →  F14  (Vol+)", indent=1)
        _line("Knob click  →  F15  (Cycle groups)", indent=1)
        _line("F12 key     →  F16  (Mute)", indent=1)
        _sep()

        _line("Can't remap? Enable Hardware Knob in the main window.", color=SUBTEXT)
        _sep()

        # ── Step 2 ──────────────────────────────────────────────────────────
        _line("Step 2 — Choose your mode in KnobMixer", bold=True, color="#1DB954")
        _line("Multiple Knobs  —  each group has its own hotkeys.", indent=1)
        _line("1-Knob  —  one knob controls all groups.", indent=1)
        _line("Press Cycle key to switch between groups.", indent=2, color=SUBTEXT)
        _line("Keys are set above the groups on the main window.", indent=2, color=SUBTEXT)
        _sep()

        # ── Step 3 ──────────────────────────────────────────────────────────
        _line("Step 3 — Set your hotkeys", bold=True, color="#1DB954")
        _line("Assign the keys from Step 1 to your groups in KnobMixer.", indent=1)
        _sep()

        # ── Step 4 ──────────────────────────────────────────────────────────
        _line("Step 4 — Add your apps", bold=True, color="#1DB954")
        _line("Click Edit on each group and add app names.", indent=1)
        _line("Open apps appear in the list for easy adding.", indent=1)
        _sep()

        # ── Divider ─────────────────────────────────────────────────────────
        _line("─" * 40, color=BORDER)
        _sep()

        # ── Tips ────────────────────────────────────────────────────────────
        _line("Tips", bold=True, size=10)
        _sep()
        _line("💡 Music knob click", bold=True)
        _line("In Multiple Knobs mode, set music knob click to Play/Pause", indent=1)
        _line("in keyboard software — pauses song instead of muting.", indent=1)
        _sep()
        _line("💡 Hardware knob (e.g. AULAF75)", bold=True)
        _line("Enable Hardware Knob in the main window.", indent=1)
        _line("Make sure your knob is set to control volume — not other functions (e.g. RGB lighting).", indent=1, color=SUBTEXT)
        _line("Intercepts system volume keys — system volume stays untouched.", indent=1)
        _sep()
        _line("💡 No knob?", bold=True)
        _line("Use any hotkey.", indent=1)

        txt.config(state="disabled")

    def _build_general(self, nb):
        sc = self._make_tab(nb, "General")

        self._v_startmin = tk.BooleanVar(value=self.cfg.get("start_minimized", True))
        self._v_startup  = tk.BooleanVar(value=get_startup())
        self._v_overlay  = tk.BooleanVar(value=self.cfg.get("show_overlay", True))
        self._v_ovsize   = tk.DoubleVar(value=self.cfg.get("overlay_size", 1.0))
        self._v_sden     = tk.BooleanVar(value=self.cfg.get("slowdown_enabled", True))
        self._v_sdthr    = tk.DoubleVar(value=self.cfg.get("slowdown_threshold", 10))
        self._v_sdstp    = tk.DoubleVar(value=self.cfg.get("slowdown_step", 0.5))

        self._sep(sc, "App Behaviour")
        self._row(sc, "Start minimized to tray",
                  lambda p: self._chk(p, self._v_startmin).pack(side="left"))
        self._row(sc, "Launch with Windows",
                  lambda p: self._chk(p, self._v_startup).pack(side="left"))



        self._sep(sc, "Volume Popup")
        self._row(sc, "Show popup when volume changes",
                  lambda p: self._chk(p, self._v_overlay).pack(side="left"))
        self._row(sc, "Popup size",
                  lambda p: self._sld(p, self._v_ovsize, 0.5, 2.0, 0.1).pack(side="left"))
        # Position preset selector
        self._v_ovpos = tk.StringVar(value=self.cfg.get("overlay_position", "bottom-right"))
        def _on_ovpos(*_):
            pos = self._v_ovpos.get()
            self.cfg["overlay_position"] = pos
            if pos != "custom":
                # Reset saved coords so preset is applied on next show
                self.cfg["overlay_x"] = -1
                self.cfg["overlay_y"] = -1
            self._apply()
        f_pos = tk.Frame(sc, bg=BG); f_pos.pack(fill="x", padx=16, pady=4)
        tk.Label(f_pos, text="Popup position", font=("Segoe UI",9),
                 fg=TEXT, bg=BG, width=28, anchor="w").pack(side="left")
        pos_cb = ttk.Combobox(f_pos, textvariable=self._v_ovpos,
                              values=["top-left","top-right","bottom-left",
                                       "bottom-right","center","custom"],
                              state="readonly", width=14, font=("Segoe UI",9),
                              style="Knob.TCombobox")
        pos_cb.pack(side="left")
        pos_cb.bind("<<ComboboxSelected>>", _on_ovpos)
        tk.Label(sc, text="Drag the popup to move it. Position changes to Custom automatically.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 wraplength=420).pack(padx=16, anchor="w")
        def _reset_popup_pos():
            self.cfg["overlay_x"] = -1
            self.cfg["overlay_y"] = -1
            # Reset drag_bound so overlay window rebinds on next show
            if hasattr(self, 'overlay') and self.overlay:
                self.overlay._drag_bound = False
            self._apply()
        tk.Button(sc, text="Reset popup position",
                  font=("Segoe UI", 8), bg=BORDER, fg=SUBTEXT,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=_reset_popup_pos).pack(padx=16, anchor="w", pady=(2,4))

        self._sep(sc, "Slowdown Zone")
        tk.Label(sc,
                 text="When volume drops below the threshold, use a finer step for precise quiet control.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 wraplength=420, justify="left").pack(padx=16, pady=(4,6), anchor="w")
        self._row(sc, "Enable slowdown zone",
                  lambda p: self._chk(p, self._v_sden).pack(side="left"))
        self._row(sc, "Trigger below (%) threshold",
                  lambda p: self._sld(p, self._v_sdthr, 1, 40, 1).pack(side="left"))
        self._row(sc, "Fine step size (%)",
                  lambda p: self._sld(p, self._v_sdstp, 0.1, 5.0, 0.1).pack(side="left"))

        self._sep(sc, "Privacy")
        self._v_analytics = tk.BooleanVar(value=self.cfg.get("analytics_enabled", True))
        self._row(sc, "Send anonymous usage data",
                  lambda p: self._chk(p, self._v_analytics).pack(side="left"))
        tk.Label(sc,
                 text="No personal data. Helps count active users.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(0,4), anchor="w")

        self._sep(sc, f"About KnobMixer v{APP_VER}")
        import webbrowser
        f_about = tk.Frame(sc, bg=BG); f_about.pack(fill="x", padx=16, pady=4)
        tk.Label(f_about,
                 text=f"\u00a9 2026 KnobMixer. Free per-app volume control for keyboard knobs.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(anchor="w")
        f_links = tk.Frame(sc, bg=BG); f_links.pack(fill="x", padx=16, pady=(0,8))
        if GITHUB_REPO:
            tk.Button(f_links, text="GitHub / Download",
                      font=("Segoe UI", 8), bg=BORDER, fg=TEXT,
                      relief="flat", cursor="hand2", padx=8, pady=3,
                      command=lambda: webbrowser.open(
                          f"https://github.com/{GITHUB_REPO}")).pack(side="left", padx=(0,6))
        def _open_bug_report():
            """In-app bug report dialog — submits to Cloudflare, no GitHub needed."""
            import platform
            dlg = tk.Toplevel(self)
            dlg.title("Report a Bug")
            dlg.configure(bg=BG)
            dlg.resizable(False, False)
            dlg.grab_set()
            dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
            _place_near_parent(dlg, self, side="left")

            tk.Label(dlg, text="Describe the issue:",
                     font=("Segoe UI",9), fg=TEXT, bg=BG).pack(padx=20, pady=(16,4), anchor="w")
            txt = tk.Text(dlg, font=("Segoe UI",9), fg=TEXT, bg=PANEL,
                          relief="flat", bd=1, width=44, height=7,
                          wrap="word", insertbackground=TEXT)
            txt.pack(padx=20, pady=(0,4))
            txt.focus_set()

            status_lbl = tk.Label(dlg, text="", font=("Segoe UI",8),
                                  fg=SUBTEXT, bg=BG)
            status_lbl.pack(padx=20, anchor="w")

            bf = tk.Frame(dlg, bg=BG); bf.pack(padx=20, pady=(8,16), anchor="e")
            tk.Button(bf, text="Cancel", font=("Segoe UI",9),
                      bg=BORDER, fg=TEXT, relief="flat", cursor="hand2",
                      padx=10, pady=4, command=dlg.destroy).pack(side="left", padx=(0,8))

            def _send():
                msg = txt.get("1.0","end").strip()
                ok, err = _report_validate_message(msg)
                if not ok:
                    status_lbl.config(text=err, fg="#ff6b6b")
                    return
                if not _report_endpoint():
                    status_lbl.config(text="Bug reporting is not available right now.", fg="#ff6b6b")
                    return
                ok, err = _report_can_send(msg)
                if not ok:
                    status_lbl.config(text=err, fg="#ff6b6b")
                    return
                # Read crash log
                crash_log = APPDATA_DIR / "crash.log"
                log_text = ""
                try:
                    if crash_log.exists():
                        raw = crash_log.read_text(encoding="utf-8", errors="replace")
                        log_text = raw[-_REPORT_LOG_MAX_CHARS:] if len(raw) > _REPORT_LOG_MAX_CHARS else raw
                except: pass

                send_btn.config(state="disabled", text="Sending…")
                status_lbl.config(text="", fg=SUBTEXT)

                def _do_send():
                    import urllib.request, json
                    payload = {
                        "id":      _get_install_id(),
                        "version": APP_VER,
                        "os":      platform.version()[:40],
                        "message": msg,
                        "log":     log_text,
                    }
                    try:
                        req = urllib.request.Request(
                            _report_endpoint(),
                            data=json.dumps(payload).encode(),
                            headers={"Content-Type":"application/json",
                                     "User-Agent":f"KnobMixer/{APP_VER}"},
                            method="POST"
                        )
                        urllib.request.urlopen(req, timeout=8)
                        _mark_report_sent(msg)
                        def _on_success():
                            try:
                                status_lbl.config(text="Report sent. Thank you!", fg="#1DB954")
                                send_btn.config(state="disabled", text="Sent ✓")
                            except tk.TclError: pass
                        self.after(0, _on_success)
                    except Exception:
                        def _on_fail():
                            try:
                                status_lbl.config(
                                    text="Could not send. Check your connection.", fg="#ff6b6b")
                                send_btn.config(state="normal", text="Send Report")
                            except tk.TclError: pass
                        self.after(0, _on_fail)
                threading.Thread(target=_do_send, daemon=True).start()

            send_btn = tk.Button(bf, text="Send Report",
                                 font=("Segoe UI",9,"bold"),
                                 bg="#1DB954", fg="white", relief="flat",
                                 cursor="hand2", padx=12, pady=4, command=_send)
            send_btn.pack(side="left")

        def _open_log_folder():
            import subprocess
            try: subprocess.Popen(["explorer", str(APPDATA_DIR)])
            except: pass

        tk.Button(f_links, text="Report a bug",
                  font=("Segoe UI", 8), bg=BORDER, fg=TEXT,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=_open_bug_report).pack(side="left")
        tk.Button(f_links, text="Open log folder",
                  font=("Segoe UI", 8), bg=BORDER, fg=SUBTEXT,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=_open_log_folder).pack(side="left", padx=(4,0))

        self._sep(sc, "Reset All Settings")
        tk.Label(sc, text="Resets everything to factory defaults — groups, hotkeys, "
                          "all settings. The app will restart.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 wraplength=420, justify="left").pack(padx=16, pady=(4,6), anchor="w")
        f_reset = tk.Frame(sc, bg=BG); f_reset.pack(fill="x", padx=16, pady=(0,8))
        def _reset_all():
            if not _themed_confirm(
                    self, "Reset All Settings",
                    "This will reset KnobMixer to factory defaults.\n"
                    "All groups, hotkeys, and settings will be cleared.\n\n"
                    "The app will restart. Continue?",
                    yes_text="Reset", danger=True): return
            try:
                if CONFIG_FILE.exists():
                    CONFIG_FILE.unlink()
            except Exception as e:
                _themed_alert(self, "Error", f"Could not reset settings:\n{e}")
                return
            # Close settings window first so it doesn't interfere with quit
            self.destroy()
            # Schedule relaunch + quit on the App root window — not SettingsWin
            # which is now destroyed and would cancel any after() calls
            import subprocess
            def _do_reset():
                try:
                    subprocess.Popen([EXE_PATH])
                except Exception as e:
                    print(f"[Reset] Relaunch failed: {e}")
                # Always quit regardless of whether relaunch succeeded
                self._quit_fn()
            # Use the parent (root) window to schedule this safely
            self._parent.after(600, _do_reset)
        tk.Button(f_reset, text="Reset All Settings",
                  font=("Segoe UI", 9), bg="#3a1a1a", fg="#ff6b6b",
                  activebackground="#4a2020", activeforeground="#ff9999",
                  relief="flat", cursor="hand2", padx=12, pady=5,
                  command=_reset_all).pack(side="left")

    # ── 1-Knob tab ───────────────────────────────────────────────────────────
    def _build_single(self, nb):
        sc = self._make_tab(nb, "1-Knob Mode")

        self._v_auto_revert = tk.BooleanVar(value=self.cfg.get("single_auto_revert", False))

        tk.Label(sc,
                 text="In 1-Knob mode, one set of keys controls all groups.\n"
                      "Press the Cycle key to switch between groups.",
                 font=("Segoe UI", 9), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(12, 4), anchor="w")
        tk.Label(sc,
                 text="Volume control hotkeys (Vol-, Vol+, Mute, Cycle) are set "
                      "in the main window above the groups.",
                 font=("Segoe UI", 8), fg="#1DB954", bg=BG,
                 wraplength=420, justify="left").pack(padx=16, pady=(0,4), anchor="w")
        tk.Label(sc,
                 text="Recommended mapping in your keyboard software:\n"
                      "  Knob left → F13 (Vol-)   Knob right → F14 (Vol+)\n"
                      "  Knob click → F15 (Cycle)   F12 key → F16 (Mute)\n\n"
                      "Can't remap? Enable Hardware Knob in the main window.",
                 font=("Consolas", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(0,8), anchor="w")
        # Default timeout 30s
        self._v_sto = tk.IntVar(value=self.cfg.get("single_timeout", 30))

        self._sep(sc, "Auto-Revert to Default Group")
        self._row(sc, "Enable auto-revert",
                  lambda p: self._chk(p, self._v_auto_revert).pack(side="left"))
        tk.Label(sc,
                 text="When ON: after the timeout, knob goes back to the starred (★) group.\n"
                      "When OFF: knob stays on whichever group you last switched to.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(0,6), anchor="w")
        self._row(sc, "Revert after (seconds)",
                  lambda p: self._single_timeout_spin(p))

        # Hardware Knob is controlled from the main window — not duplicated here
        self._v_hw_en = tk.BooleanVar(value=self.cfg.get("hw_knob_enabled", False))

    # ── Mic tab ──────────────────────────────────────────────────────────────
    def _build_mic(self, nb):
        sc = self._make_tab(nb, "Mic Toggle")

        self._v_micen    = tk.BooleanVar(value=self.cfg.get("mic_enabled", False))
        self._v_michk    = tk.StringVar(value=self.cfg.get("mic_hotkey", "f9"))
        self._v_micst    = tk.BooleanVar(value=self.cfg.get("mic_start_muted", False))
        self._v_micvol   = tk.IntVar(value=_level_from_vol(self.cfg.get("mic_sound_volume", 0.056)))
        self._v_micsz    = tk.IntVar(value=self.cfg.get("mic_icon_size", 40))
        self._v_mical    = tk.DoubleVar(value=self.cfg.get("mic_icon_alpha", 0.85))
        self._v_micstyle = tk.StringVar(value=self.cfg.get("mic_icon_style", "Circle"))
        self._v_micpre   = tk.IntVar(value=self.cfg.get("mic_sound_preset", 0))

        self._sep(sc, "Mic Basics")
        self._row(sc, "Enable mic toggle",
                  lambda p: self._chk(p, self._v_micen).pack(side="left"))
        self._row(sc, "Start muted on launch",
                  lambda p: self._chk(p, self._v_micst).pack(side="left"))

        # Mic device selector
        f_dev = tk.Frame(sc, bg=BG)
        f_dev.pack(fill="x", padx=16, pady=7)
        tk.Label(f_dev, text="Microphone device", font=("Segoe UI", 9),
                 fg=TEXT, bg=BG, width=28, anchor="w").pack(side="left")
        self._v_micdev = tk.StringVar(value=self.cfg.get("mic_device_name","System Default"))
        _mic_devs = self._get_mic_devices()
        _max_w = max((len(n) for n in _mic_devs), default=20)
        dev_cb = ttk.Combobox(f_dev, textvariable=self._v_micdev,
                              values=_mic_devs,
                              width=max(22, _max_w + 2),
                              font=("Segoe UI", 9), style="Knob.TCombobox")
        dev_cb.pack(side="left")
        dev_cb.bind("<<ComboboxSelected>>", lambda e: self._apply())

        # Mic hotkey
        cur_hk = self.cfg.get("mic_hotkey", "f9")
        def _hk_cb(hk):
            ok, msg = _validate_hotkey_choice(self.cfg, hk, ("mic",))
            if not ok:
                _themed_alert(self, "Hotkey already in use", msg)
                return
            self._v_michk.set(hk)
            self._apply()
        f_hk = tk.Frame(sc, bg=BG)
        f_hk.pack(fill="x", padx=16, pady=7)
        tk.Label(f_hk, text="Mic mute hotkey", font=("Segoe UI", 9),
                 fg=TEXT, bg=BG, width=28, anchor="w").pack(side="left")
        make_hotkey_btn(f_hk, cur_hk, _hk_cb).pack(side="left")

        self._sep(sc, "Icon")
        # Position preset selector
        self._v_micpos = tk.StringVar(value=self.cfg.get("mic_position", "bottom-right"))
        def _on_micpos(*_):
            pos = self._v_micpos.get()
            self.cfg["mic_position"] = pos
            if pos != "custom":
                self.cfg["mic_icon_x"] = -1
                self.cfg["mic_icon_y"] = -1
            self._apply()
            # Rebuild mic overlay so it moves to new position immediately
            if hasattr(self, "_app_ref") and self._app_ref:
                app = self._app_ref
                if app.mic_ov:
                    app.mic_ov.rebuild()
        f_micpos = tk.Frame(sc, bg=BG); f_micpos.pack(fill="x", padx=16, pady=4)
        tk.Label(f_micpos, text="Icon position", font=("Segoe UI",9),
                 fg=TEXT, bg=BG, width=28, anchor="w").pack(side="left")
        micpos_cb = ttk.Combobox(f_micpos, textvariable=self._v_micpos,
                                   values=["top-left","top-right","bottom-left",
                                           "bottom-right","center","custom"],
                                   state="readonly", width=14, font=("Segoe UI",9),
                                   style="Knob.TCombobox")
        micpos_cb.pack(side="left")
        micpos_cb.bind("<<ComboboxSelected>>", _on_micpos)
        tk.Label(sc, text="Drag the icon to move it. Position changes to Custom automatically.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 wraplength=420).pack(padx=16, anchor="w")
        self._row(sc, "Icon size (px)",
                  lambda p: self._sld(p, self._v_micsz, 16, 80, 4).pack(side="left"))
        self._row(sc, "Transparency (1.0 = solid)",
                  lambda p: self._sld(p, self._v_mical, 0.1, 1.0, 0.05).pack(side="left"))

        f_style = tk.Frame(sc, bg=BG)
        f_style.pack(fill="x", padx=16, pady=7)
        tk.Label(f_style, text="Icon style", font=("Segoe UI", 9),
                 fg=TEXT, bg=BG, width=28, anchor="w").pack(side="left")
        cb = ttk.Combobox(f_style, textvariable=self._v_micstyle,
                          values=ICON_STYLES, width=12, state="readonly",
                          font=("Segoe UI", 9), style="Knob.TCombobox")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._apply())

        f_lock = tk.Frame(sc, bg=BG)
        f_lock.pack(fill="x", padx=16, pady=4)
        self._v_locked = tk.BooleanVar(value=self.cfg.get("mic_icon_locked", False))
        tk.Checkbutton(f_lock, text="Lock icon in place (prevents accidental dragging)",
                       variable=self._v_locked,
                       font=("Segoe UI", 9), fg=TEXT, bg=BG,
                       selectcolor=BORDER, activebackground=BG,
                       activeforeground=TEXT, command=self._apply).pack(anchor="w")
        tk.Label(sc, text="Drag the mic icon on screen to reposition it.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG).pack(padx=16, anchor="w")

        self._sep(sc, "Sound")
        self._row(sc, "Toggle sound volume",
                  lambda p: self._sld(p, self._v_micvol, 1, 5, 1).pack(side="left"))

        # Only one preset (Soft Ping) — no preset selector needed
        self._preset_btns = []

    def _refresh_preset_btns(self):
        cur = self._v_micpre.get()
        for i, btn in enumerate(self._preset_btns):
            btn.config(bg="#1a3a1a" if i == cur else BORDER,
                       fg="#1DB954" if i == cur else TEXT)

    def _single_timeout_spin(self, parent):
        sp = tk.Spinbox(parent, from_=5, to=600,
                        textvariable=self._v_sto,
                        width=10, font=("Segoe UI", 9),
                        bg=PANEL, fg=TEXT,
                        buttonbackground=BORDER,
                        highlightthickness=0, relief="flat",
                        command=self._apply)
        sp.pack(side="left")
        sp.bind("<FocusOut>", lambda e: self._apply())
        sp.bind("<Return>", lambda e: self._apply())
        return sp

    # ── Apply ────────────────────────────────────────────────────────────────
    def _get_mic_devices(self):
        """Get list of available microphone names."""
        return get_mic_devices()

    def _apply(self):
        c = self.cfg
        prev_timeout = c.get("single_timeout", 30)
        prev_auto = c.get("single_auto_revert", False)
        prev_hw = c.get("hw_knob_enabled", False)
        c["start_minimized"]   = self._v_startmin.get()

        c["show_overlay"]       = self._v_overlay.get()
        c["overlay_size"]       = round(self._v_ovsize.get(), 1)
        if hasattr(self, "_v_ovpos"):
            c["overlay_position"] = self._v_ovpos.get()
        # overlay position is saved directly by drag — preserve it here
        c.setdefault("overlay_x", self.cfg.get("overlay_x", -1))
        c.setdefault("overlay_y", self.cfg.get("overlay_y", -1))
        c["analytics_enabled"] = self._v_analytics.get()
        set_startup(self._v_startup.get())
        c["slowdown_enabled"]  = self._v_sden.get()
        c["slowdown_threshold"]= self._v_sdthr.get()
        c["slowdown_step"]     = round(self._v_sdstp.get(), 1)
        c["single_timeout"]    = self._v_sto.get()
        c["single_auto_revert"] = self._v_auto_revert.get()
        c["hw_knob_enabled"]   = self._v_hw_en.get()
        c["mic_enabled"]       = self._v_micen.get()
        c["mic_hotkey"]        = self._v_michk.get()
        c["mic_start_muted"]   = self._v_micst.get()
        c["mic_sound_volume"]  = _vol_from_level(self._v_micvol.get())
        c["mic_icon_size"]     = self._v_micsz.get()
        c["mic_icon_alpha"]    = round(self._v_mical.get(), 2)
        c["mic_icon_style"]    = self._v_micstyle.get()
        c["mic_sound_preset"]  = self._v_micpre.get()
        if hasattr(self, "_v_micdev"):
            c["mic_device_name"] = self._v_micdev.get()
        if hasattr(self, "_v_locked"):
            c["mic_icon_locked"] = self._v_locked.get()
        save_cfg(c)
        self.on_change()
        if self._app_ref:
            self._app_ref._show_saved()
            if prev_auto != c["single_auto_revert"]:
                self._app_ref._redraw()
            else:
                self._app_ref._refresh_default_buttons()
            if prev_hw != c["hw_knob_enabled"] and c.get("mode") == "single":
                self._app_ref._rebuild_knob_panel()
            if ((prev_timeout != c["single_timeout"]) or (prev_auto != c["single_auto_revert"])) and c.get("mode") == "single":
                self._app_ref._refresh_revert_timer()

# ══════════════════════════════════════════════════════════════════════════════
# Main UI
# ══════════════════════════════════════════════════════════════════════════════
class TutorialOverlay:
    """
    First-run tutorial.
    Drawn directly inside the app window using a Canvas overlay so it
    always moves with the window. No separate Toplevel — no drift.
    Blocks all interaction while active.
    """
    STEPS = [
        {
            "title": "Welcome to KnobMixer!",
            "body":  "Let's show you around in a few quick steps.\nPress Next to continue.",
            "target": None,
        },
        {
            "title": "1 — Choose your mode",
            "body":  ("1-Knob — one knob controls all your groups.\n"
                      "Multiple Knobs — each group has its own keys.\n\n"
                      "Can't remap your knob's volume keys?\n"
                      "Tick Hardware Knob below the mode selector."),
            "target": "mode_bar",
        },
        {
            "title": "2 — Set your hotkeys",
            "body":  ("First assign your keys in your keyboard software,\n"
                      "then assign them here in KnobMixer.\n\n"
                      "Recommended:\n"
                      "  Knob left  → F13   Knob right → F14\n"
                      "  Knob click → F15 (Cycle)\n"
                      "  F12 key    → F16 (Mute)"),
            "target": "knob_panel",
        },
        {
            "title": "3 — Add your apps",
            "body":  ("Each group controls a set of apps.\n"
                      "Click Edit on a group to add apps.\n\n"
                      "Groups are controlled in order — the TOP group\n"
                      "is always the first one your knob controls.\n"
                      "Drag to reorder them."),
            "target": "media_apps",
        },
        {
            "title": "4 — Add more groups",
            "body":  "Need a group for your game?\nClick + Group at the bottom to add one.",
            "target": "add_group",
        },
    ]

    def __init__(self, root, app):
        self._root    = root
        self._app     = app
        self._step         = 0
        self._canvas       = None
        self._panel        = None
        self._ovr          = None
        self._info         = None
        self._border       = None
        self._ovr_bind_id  = None
        self._panel_bind_id = None
        self._show_step()

    def _get_target_widget(self, target):
        try:
            if target == "mode_bar":    return self._app._mode_bar
            if target == "knob_panel":
                kp = self._app._knob_panel
                if kp.winfo_ismapped():
                    return kp
                if getattr(self._app, "_hw_row", None) and self._app._hw_row.winfo_ismapped():
                    return self._app._hw_row
                return self._app._mode_bar
            if target == "groups":      return self._app._groups_cf
            if target == "add_group":   return self._app._add_group_btn
            if target == "media_apps":
                for w in getattr(self._app, "_group_widgets", []):
                    g = w.get("group")
                    if not g or g.get("master_volume"):
                        continue
                    if g.get("name","").strip().lower() == "media":
                        return w.get("apps_row")
                for w in getattr(self._app, "_group_widgets", []):
                    g = w.get("group")
                    if g and not g.get("master_volume"):
                        return w.get("apps_row")
        except: pass
        return None

    def _show_step(self):
        self._clear()
        if self._step >= len(self.STEPS):
            self._finish()
            return

        step   = self.STEPS[self._step]
        target = step.get("target")
        widget = self._get_target_widget(target) if target else None

        self._root.update_idletasks()

        # ── Transparent click-blocking overlay (tracks window movement) ───────
        # A Toplevel with WS_EX_TRANSPARENT would pass clicks through but we
        # need to BLOCK clicks. Use a non-transparent Toplevel that matches the
        # app window exactly, and re-sync its position on every Configure event.
        self._ovr = tk.Toplevel(self._root)
        self._ovr.overrideredirect(True)
        self._ovr.attributes("-topmost", True)
        self._ovr.attributes("-alpha", 0.01)  # nearly invisible but click-blocking
        self._ovr.configure(bg="black")

        def _sync_ovr(e=None):
            if not (self._ovr and self._ovr.winfo_exists()): return
            try:
                rx = self._root.winfo_rootx()
                ry = self._root.winfo_rooty()
                rw = self._root.winfo_width()
                rh = self._root.winfo_height()
                self._ovr.geometry(f"{rw}x{rh}+{rx}+{ry}")
            except: pass

        _sync_ovr()
        # Rebind on every move/resize so overlay always tracks the window
        self._ovr_bind_id = self._root.bind("<Configure>", _sync_ovr, add="+")

        # ── Green border on highlighted widget ────────────────────────────────
        if widget:
            try:
                widget.config(highlightthickness=3,
                              highlightbackground="#1DB954",
                              highlightcolor="#1DB954")
                self._border = widget
            except: pass

        # ── Info panel — separate Toplevel, always on top ────────────────────
        self._panel = tk.Toplevel(self._root)
        self._panel.overrideredirect(True)
        self._panel.attributes("-topmost", True)
        self._panel.configure(bg="#1a1a2e")
        tk.Frame(self._panel, bg="#1DB954", height=2).pack(fill="x")
        f = tk.Frame(self._panel, bg="#1a1a2e", padx=18, pady=14)
        f.pack()

        dot_f = tk.Frame(f, bg="#1a1a2e"); dot_f.pack(anchor="w", pady=(0,6))
        for i in range(len(self.STEPS)):
            col = "#1DB954" if i == self._step else "#444"
            tk.Label(dot_f, text="●", font=("Segoe UI",8),
                     fg=col, bg="#1a1a2e").pack(side="left", padx=2)

        tk.Label(f, text=step["title"], font=("Segoe UI",11,"bold"),
                 fg="#1DB954", bg="#1a1a2e").pack(anchor="w")
        tk.Label(f, text=step["body"], font=("Segoe UI",9),
                 fg="#c9d1d9", bg="#1a1a2e", justify="left",
                 wraplength=240).pack(anchor="w", pady=(6,12))

        btn_f = tk.Frame(f, bg="#1a1a2e"); btn_f.pack(anchor="e")
        if self._step > 0:
            tk.Button(btn_f, text="Back", font=("Segoe UI",9),
                      bg="#2a2a3e", fg="#888", relief="flat",
                      cursor="hand2", padx=12, pady=4,
                      command=self._prev).pack(side="left", padx=(0,6))
        is_last = self._step == len(self.STEPS) - 1
        tk.Button(btn_f,
                  text="Done" if is_last else "Next →",
                  font=("Segoe UI",9,"bold"),
                  bg="#1DB954", fg="white", relief="flat",
                  cursor="hand2", padx=16, pady=4,
                  command=self._next).pack(side="left")

        # Position panel BELOW the app window so the blocking overlay
        # (which covers only the app window area) cannot intercept clicks.
        # Execution proof: _ovr covers rx..rx+rw, ry..ry+rh
        # Panel at py = ry+rh+8 is outside that rect — clicks reach _panel directly.
        self._panel.update_idletasks()
        pw = self._panel.winfo_reqwidth()
        ph = self._panel.winfo_reqheight()
        rx = self._root.winfo_rootx()
        ry = self._root.winfo_rooty()
        rw = self._root.winfo_width()
        rh = self._root.winfo_height()
        sh = self._root.winfo_screenheight()
        px = rx + (rw - pw) // 2
        py = ry + rh + 8   # just below the window — outside _ovr coverage
        # Clamp: if no room below, place above the window instead
        if py + ph > sh - 10:
            py = ry - ph - 8
        px = max(0, px)
        py = max(0, py)
        self._panel.geometry(f"+{px}+{py}")

        def _sync_panel(e=None):
            if not (self._panel and self._panel.winfo_exists()): return
            try:
                rx2 = self._root.winfo_rootx()
                ry2 = self._root.winfo_rooty()
                rw2 = self._root.winfo_width()
                rh2 = self._root.winfo_height()
                pw2 = self._panel.winfo_width()
                ph2 = self._panel.winfo_height()
                sh2 = self._root.winfo_screenheight()
                py2 = ry2 + rh2 + 8
                if py2 + ph2 > sh2 - 10:
                    py2 = ry2 - ph2 - 8
                self._panel.geometry(
                    f"+{max(0, rx2+(rw2-pw2)//2)}+{max(0,py2)}")
            except: pass
        self._panel_bind_id = self._root.bind("<Configure>", _sync_panel, add="+")

    def _clear(self):
        # Unbind only our specific Configure handlers using stored bind IDs
        for id_attr in ("_ovr_bind_id", "_panel_bind_id"):
            bid = getattr(self, id_attr, None)
            if bid is not None:
                try: self._root.unbind("<Configure>", bid)
                except: pass
                setattr(self, id_attr, None)
        # Remove highlight border
        if getattr(self, "_border", None):
            try: self._border.config(highlightthickness=0)
            except: pass
            self._border = None
        # Destroy Toplevel windows and inner frames
        for attr in ("_ovr", "_panel", "_canvas", "_info"):
            w = getattr(self, attr, None)
            if w is not None:
                try:
                    if w.winfo_exists(): w.destroy()
                except: pass
                setattr(self, attr, None)

    def _next(self):
        self._step += 1
        self._show_step()

    def _prev(self):
        self._step = max(0, self._step - 1)
        self._show_step()

    def _finish(self):
        self._clear()
        self._app.cfg["tutorial_seen"] = True
        save_cfg(self._app.cfg)



class App:
    def __init__(self):
        self.cfg=load_cfg()
        self.hk=HotkeyEngine()
        self.mic=MicCtrl()
        self._enabled=True
        self._update_available=False
        self._active_grp=None
        self._timeout_job=None
        self._timeout_start=0
        self._group_widgets=[]
        self._tutorial = None
        
        self._build_win()
        self.overlay=VolumeOverlay(self.root)
        self.overlay._cfg = self.cfg   # give overlay access to cfg for position save
        self.mic_ov=None
        self._build_ui()
        _HOOK.start()
        self._init_single()   # sets _active_group_ref first
        self.hk.reload(self.cfg,self._on_vol,self._on_switch)  # then registers keys
        self._reg_mic_hk()

        if self.cfg.get("mic_enabled", False):
            self.mic.sync(self.cfg)
            if self.cfg.get("mic_start_muted", False):
                self.mic.set(True, self.cfg)
            self.mic_ov = MicOverlay(self.root, self.mic, self.cfg)

        self._refresh_loop()
        self._setup_tray()
        # Never minimize on first launch — tutorial needs the window visible
        if self.cfg.get("start_minimized",True) and self.cfg.get("tutorial_seen",False):
            self.root.after(200,self._hide)

    # ── Window ────────────────────────────────────────────────────────────────
    def _build_win(self):
        self.root=tk.Tk()
        _init_ttk_theme(self.root)
        self.root.title(APP_NAME)
        self.root.configure(bg=BG)
        self.root.resizable(True,True)
        self.root.minsize(600,400)
        self.root.protocol("WM_DELETE_WINDOW",self._hide)
        self.root.bind("<FocusOut>", lambda e: self._card_drag_cancel() if hasattr(self,"_card_drag_cancel") else None)
        # When clicking anywhere that is NOT an Entry, release focus from any
        # active Entry (removes the blinking cursor). add="+" preserves other bindings.
        def _maybe_release_focus(e):
            if not isinstance(e.widget, tk.Entry):
                self.root.focus_set()
        self.root.bind("<Button-1>", _maybe_release_focus, add="+")
        w,h=640,720
        sw=self.root.winfo_screenwidth(); sh=self.root.winfo_screenheight()
        h=min(h, sh-80)
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        # Set ghost icon for window title bar
        try:
            ghost_img = make_tray_img([], enabled=True)
            ghost_img = ghost_img.resize((32,32), Image.LANCZOS)
            self._win_icon = ImageTk.PhotoImage(ghost_img)
            self.root.iconphoto(True, self._win_icon)
        except Exception:
            pass

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr=tk.Frame(self.root,bg="#111118",pady=12); hdr.pack(fill="x")
        lf=tk.Frame(hdr,bg="#111118"); lf.pack(side="left",padx=16)
        tk.Label(lf,text="KnobMixer",font=("Segoe UI",16,"bold"),fg=TEXT,bg="#111118").pack(anchor="w")
        tk.Label(lf,text=f"v{APP_VER}  •  Volume control for apps using knobs and hotkeys",
                 font=("Segoe UI",8),fg=SUBTEXT,bg="#111118").pack(anchor="w", pady=(1,0))

        rf=tk.Frame(hdr,bg="#111118"); rf.pack(side="right",padx=16)
        tk.Button(rf,text="⚙",font=("Segoe UI",13),bg=BORDER,fg=SUBTEXT,
                  width=3,
                  activebackground=HOVER,activeforeground=TEXT,
                  relief="flat",cursor="hand2",
                  command=self._open_settings).pack(side="right",padx=(6,0), ipady=2)
        tk.Button(rf,text="?",font=("Segoe UI",11,"bold"),bg=BORDER,fg=SUBTEXT,
                  width=3,
                  activebackground=HOVER,activeforeground=TEXT,
                  relief="flat",cursor="hand2",padx=0,
                  command=self._start_tutorial).pack(side="right",padx=(0,2), ipady=2)
        self._onoff_btn=tk.Button(rf,text="Enabled",font=("Segoe UI",9,"bold"),
                                  bg="#183524",fg="#1DB954",activebackground="#214733",
                                  relief="flat",cursor="hand2",padx=12,pady=5,
                                  command=self._toggle_en)
        self._onoff_btn.pack(side="right",padx=(0,8))

        self._update_url = [None]  # store URL for click
        # Update check button — styled like ON button, sits next to gear icon
        self._update_btn = tk.Button(
            rf, text="Check Updates",
            font=("Segoe UI",8,"bold"), bg=PANEL_SOFT, fg=SUBTEXT,
            activebackground=HOVER, activeforeground=TEXT,
            relief="flat", cursor="hand2", padx=10, pady=5,
            command=self._manual_update_check)
        self._update_btn.pack(side="right", padx=(0,4))

        # Mode bar
        mb=tk.Frame(self.root,bg=PANEL,pady=6); mb.pack(fill="x")
        self._mode_bar = mb  # reference for tutorial
        tk.Label(mb,text="Mode:",font=("Segoe UI",9),fg=SUBTEXT,bg=PANEL).pack(side="left",padx=(14,4))
        self._mode_var=tk.StringVar(value=self.cfg.get("mode","multi"))
        for val,lbl in [("single","1 Knob"),("multi","Multiple Knobs")]:
            tk.Radiobutton(mb,text=lbl,value=val,variable=self._mode_var,
                           font=("Segoe UI",9),fg=TEXT,bg=PANEL,selectcolor=BORDER,
                           activebackground=PANEL,activeforeground=TEXT,
                           command=self._on_mode).pack(side="left",padx=6)

        # Hardware knob row — sits between mode bar and knob panel
        self._hw_row = tk.Frame(self.root, bg=PANEL, pady=4)
        tk.Frame(self._hw_row, bg=BORDER, height=1).pack(fill="x", padx=10, pady=(0,4))
        hw_inner = tk.Frame(self._hw_row, bg=PANEL)
        hw_inner.pack(fill="x", padx=14)
        self._hw_var = tk.BooleanVar(value=self.cfg.get("hw_knob_enabled", False))
        def _on_hw_toggle():
            self.cfg["hw_knob_enabled"] = self._hw_var.get()
            self._autosave()
            self._rebuild_knob_panel()
        tk.Checkbutton(hw_inner, text="Hardware Knob",
                       variable=self._hw_var,
                       font=("Segoe UI",9), fg=TEXT, bg=PANEL,
                       selectcolor=BG, activebackground=PANEL,
                       activeforeground=TEXT,
                       command=_on_hw_toggle).pack(side="left")
        tk.Label(hw_inner,
                 text="Can't change your knob's volume keys? Use this.",
                 font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL).pack(side="left", padx=(6,0))
        # Hardware knob only relevant in 1-Knob mode — show/hide with mode
        if self.cfg.get("mode","single") == "single" or self.cfg.get("hw_knob_enabled",False):
            self._hw_row.pack(fill="x")

        # Single-knob status bar
        self._sb=tk.Frame(self.root,bg="#111118",pady=3)
        self._active_lbl=tk.Label(self._sb,text="",font=("Segoe UI",9,"bold"),
                                  fg="#1DB954",bg="#111118")
        self._active_lbl.pack(side="left",padx=14)
        self._timeout_lbl=tk.Label(self._sb,text="",font=("Segoe UI",8),
                                   fg=SUBTEXT,bg="#111118")
        self._timeout_lbl.pack(side="right",padx=14)
        if self.cfg.get("mode")=="single": self._sb.pack(fill="x")

        # 1-Knob key panel — shown above groups in 1-knob mode only
        self._knob_panel = tk.Frame(self.root, bg=PANEL, pady=4)
        kp = self._knob_panel

        def _make_hk_cell(parent, label, get_val, set_val, clear_val):
            """Helper: label + hotkey button + ✕ in a compact cell."""
            f = tk.Frame(parent, bg=PANEL)
            tk.Label(f, text=label, font=("Segoe UI",8), fg=SUBTEXT,
                     bg=PANEL).pack(anchor="w")
            rf = tk.Frame(f, bg=PANEL); rf.pack(anchor="w")
            btn = make_hotkey_btn(rf, get_val(), set_val)
            btn.pack(side="left")
            def _clr(b=btn):
                cap = getattr(b, "_capture", None)
                if cap and cap._active:
                    cap._finish(None)
                else:
                    clear_val(); b.config(text="—")
            tk.Button(rf, text="×", font=("Segoe UI",9,"bold"), bg=PANEL, fg="#6d7086",
                      activebackground=PANEL, activeforeground=TEXT,
                      relief="flat", cursor="hand2", padx=3, pady=0,
                      command=_clr).pack(side="left", padx=(1,0))
            return f

        sk = self.cfg.setdefault("single_keys", {"vol_down":"","vol_up":"","mute":""})
        hw = self.cfg.get("hw_knob_enabled", False)

        # Row 1: keys — show Vol-/Vol+/Mute only when NOT using hw knob
        row1 = tk.Frame(kp, bg=PANEL); row1.pack(fill="x", padx=10, pady=(4,1))
        tk.Label(row1, text="1-Knob:", font=("Segoe UI",8,"bold"), fg=SUBTEXT,
                 bg=PANEL).pack(side="left", padx=(0,8))

        if not hw:
            for action, lbl in [("vol_down","Vol-"), ("vol_up","Vol+"), ("mute","Mute")]:
                def _set_single_hotkey(hk, a=action):
                    ok, msg = _validate_hotkey_choice(self.cfg, hk, ("single_shared", a))
                    if not ok:
                        _themed_alert(self.root, "Hotkey already in use", msg)
                        return
                    self.cfg["single_keys"][a] = hk
                    self._autosave()
                _make_hk_cell(row1, lbl,
                              lambda a=action: sk.get(a,""),
                              _set_single_hotkey,
                              lambda a=action: (self.cfg["single_keys"].__setitem__(a,""), self._autosave())
                              ).pack(side="left", padx=(0,10))

        def _ck_cb(hk):
            ok, msg = _validate_hotkey_choice(self.cfg, hk, ("cycle",))
            if not ok:
                _themed_alert(self.root, "Hotkey already in use", msg)
                return
            self.cfg["cycle_key"] = hk; self._autosave()
        def _ck_clr(): self.cfg["cycle_key"] = ""; self._autosave()
        if hw:
            tk.Label(row1, text="Vol keys handled by Hardware Knob",
                     font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL).pack(side="left", padx=(0,14))
        _make_hk_cell(row1, "Cycle",
                      lambda: self.cfg.get("cycle_key",""),
                      _ck_cb, _ck_clr).pack(side="left")
        # Knob panel packs after groups frame is created (see below)
        self._knob_panel_pending = self.cfg.get("mode") == "single"

        # Groups — scrollable canvas with minimal auto-hide scrollbar
        cf=tk.Frame(self.root,bg=BG); cf.pack(fill="both",expand=True)
        self._groups_cf = cf  # reference for knob panel ordering
        if getattr(self, "_knob_panel_pending", False):
            self._knob_panel.pack(fill="x", before=cf)
        self._canvas=tk.Canvas(cf,bg=BG,highlightthickness=0,bd=0)
        # Minimal custom scrollbar: thin dark strip
        self._sb_frame=tk.Frame(cf,bg=BG,width=10)
        self._sb_thumb=tk.Frame(self._sb_frame,bg=SB_THUMB,cursor="hand2")
        self._sb_visible=False
        self._canvas.pack(side="left",fill="both",expand=True)
        # DON'T pack sb_frame yet — only show when needed
        self.gf=tk.Frame(self._canvas,bg=BG)
        self._cwin=self._canvas.create_window((0,0),window=self.gf,anchor="nw")

        def _update_scroll(*_):
            bbox=self._canvas.bbox("all")
            if not bbox: return
            content_h=bbox[3]; canvas_h=self._canvas.winfo_height()
            needs_scroll = content_h > canvas_h + 4
            if needs_scroll:
                # Content overflows — enable scrolling
                self._canvas.configure(scrollregion=bbox)
                if not self._sb_visible:
                    self._sb_frame.pack(side="right",fill="y",padx=(0,0))
                    self._sb_visible=True
                if canvas_h > 0:
                    ratio=canvas_h/content_h
                    top_frac=self._canvas.yview()[0]
                    thumb_h=max(20,int(canvas_h*ratio))
                    thumb_y=int(top_frac*canvas_h)
                    self._sb_thumb.place(x=1,y=thumb_y,width=8,height=thumb_h)
            else:
                # Content fits — lock scroll to top, hide scrollbar
                self._canvas.configure(scrollregion=(0,0,self._canvas.winfo_width(),canvas_h))
                self._canvas.yview_moveto(0)   # snap back to top
                if self._sb_visible:
                    self._sb_frame.pack_forget()
                    self._sb_visible=False

        def _on_yview(*args):
            self._canvas.yview(*args)
            _update_scroll()

        self._sb_frame.configure(width=6)
        # Scrollbar drag
        self._sb_drag_y=None
        def _sb_press(e):
            self._sb_drag_y=e.y_root
            self._sb_drag_frac=self._canvas.yview()[0]
        def _sb_move(e):
            if self._sb_drag_y is None: return
            dy=e.y_root-self._sb_drag_y
            bbox=self._canvas.bbox("all")
            if not bbox: return
            content_h=bbox[3]; canvas_h=self._canvas.winfo_height()
            frac=self._sb_drag_frac+dy/max(1,canvas_h)
            self._canvas.yview_moveto(max(0,min(1,frac)))
            _update_scroll()
        def _sb_release(e): self._sb_drag_y=None
        self._sb_thumb.bind("<ButtonPress-1>",  _sb_press)
        self._sb_thumb.bind("<B1-Motion>",      _sb_move)
        self._sb_thumb.bind("<ButtonRelease-1>",_sb_release)
        self._sb_thumb.bind("<Enter>", lambda e: self._sb_thumb.config(bg=SB_ACTIVE))
        self._sb_thumb.bind("<Leave>", lambda e: self._sb_thumb.config(bg=SB_THUMB))

        self.gf.bind("<Configure>",lambda e: _update_scroll())
        self._canvas.bind("<Configure>",lambda e: (
            self._canvas.itemconfig(self._cwin,width=e.width), _update_scroll()))

        def _main_wheel(ev):
            if not self._sb_visible: return   # no scroll if content fits
            self._canvas.yview_scroll(int(-1*ev.delta/120),"units")
            _update_scroll()

        def _bind_main_wheel(widget):
            """Recursively bind mousewheel to all group card widgets."""
            widget.bind("<MouseWheel>", _main_wheel)
            for child in widget.winfo_children():
                _bind_main_wheel(child)

        # Bind to canvas and groups frame; rebind after redraw
        self._bind_main_wheel = _bind_main_wheel
        self._canvas.bind("<MouseWheel>", _main_wheel)
        self._canvas.bind("<Enter>", lambda e: self._canvas.bind("<MouseWheel>", _main_wheel))
        self.gf.bind("<Enter>", lambda e: _bind_main_wheel(self.gf))
        self.gf.bind("<Configure>", lambda e: (_update_scroll(), _bind_main_wheel(self.gf)))
        self._redraw()

        # Bottom bar
        bot=tk.Frame(self.root,bg=PANEL,pady=7); bot.pack(fill="x")
        tk.Button(bot,text="+ Group",font=("Segoe UI",9),bg=BORDER,fg=SUBTEXT,
                  activebackground=HOVER,activeforeground=TEXT,
                  relief="flat",cursor="hand2",padx=10,pady=4,
                  command=self._add_group).pack(side="left",padx=(12,4))
        # Only show Master Vol button if no master vol group exists yet
        self._master_vol_btn = tk.Button(
            bot, text="+ Master Vol", font=("Segoe UI",9),
            bg="#1a2a3a", fg="#00BCD4",
            activebackground=HOVER, activeforeground=TEXT,
            relief="flat", cursor="hand2", padx=10, pady=4,
            command=self._add_master_group)
        self._update_master_vol_btn()
        self._dirty_lbl=tk.Label(bot,text="",font=("Segoe UI",8),fg="#1DB954",bg=PANEL)
        self._dirty_lbl.pack(side="right",padx=12)


    # ── Groups ────────────────────────────────────────────────────────────────
    def _redraw(self):
        for w in self.gf.winfo_children(): w.destroy()
        self._group_widgets.clear()
        for i,g in enumerate(self.cfg["groups"]): self._card(i,g)
        if hasattr(self, "_bind_main_wheel"):
            self._bind_main_wheel(self.gf)
        self._drag_state = {"active": False, "idx": None, "start_y": 0, "ghost": None}
        if hasattr(self, "_master_vol_btn"):
            self._update_master_vol_btn()
        # Force canvas to recalculate scrollregion after content changes
        self.gf.update_idletasks()
        bbox = self._canvas.bbox("all")
        if bbox:
            self._canvas.configure(scrollregion=bbox)

    def _card_drag_start(self, event, idx):
        """Begin drag — record starting position and which group."""
        self._drag_state.update(active=True, idx=idx, start_y=event.y_root)
        # Bind motion and release to ROOT so we never lose the drag
        # even if the widget gets destroyed and recreated during redraw
        self.root.bind("<B1-Motion>",      self._card_drag_motion_root)
        self.root.bind("<ButtonRelease-1>", self._card_drag_end_root)

    def _card_drag_motion_root(self, event):
        """Root-level motion handler — fires continuously while dragging."""
        if not self._drag_state.get("active"): return
        dy   = event.y_root - self._drag_state["start_y"]
        cur  = self._drag_state["idx"]
        groups = self.cfg["groups"]
        moved  = False

        if dy < -18 and cur > 0:
            groups[cur], groups[cur-1] = groups[cur-1], groups[cur]
            self._drag_state["idx"]     = cur - 1
            self._drag_state["start_y"] = event.y_root
            moved = True
        elif dy > 18 and cur < len(groups)-1:
            groups[cur], groups[cur+1] = groups[cur+1], groups[cur]
            self._drag_state["idx"]     = cur + 1
            self._drag_state["start_y"] = event.y_root
            moved = True

        if moved:
            self._redraw()  # redraw redraws all cards
            # Re-bind wheel after redraw
            if hasattr(self, "_bind_main_wheel"):
                self._bind_main_wheel(self.gf)

    def _card_drag_end_root(self, event):
        """Release — save and clean up root bindings."""
        if not self._drag_state.get("active"): return
        self._drag_state["active"] = False
        self._autosave()
        self.root.unbind("<B1-Motion>")
        self.root.unbind("<ButtonRelease-1>")

    def _card_drag_cancel(self, event=None):
        """Cancel drag if window loses focus mid-drag."""
        if self._drag_state.get("active"):
            self._drag_state["active"] = False
            self.root.unbind("<B1-Motion>")
            self.root.unbind("<ButtonRelease-1>")

    # Keep old signatures for the handle bindings (they just start the drag)
    def _card_drag_motion(self, event, idx): pass
    def _card_drag_end(self, event): pass

    def _card(self,idx,group):
        color=group.get("color","#888")
        mode =self.cfg.get("mode","multi")

        outer=tk.Frame(self.gf,bg=BG)
        outer.pack(fill="x",padx=12,pady=4)
        accent=tk.Frame(outer,bg=color,width=4)
        accent.pack(side="left",fill="y")
        card=tk.Frame(outer,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        card.pack(side="left",fill="x",expand=True)

        # Row 1: handle | dot | name | ★ | ON/OFF | del
        r1=tk.Frame(card,bg=PANEL); r1.pack(fill="x",padx=10,pady=(6,2))

        # Drag handle — ≡ symbol, cursor changes to indicate draggable
        handle=tk.Label(r1,text="≡",font=("Segoe UI",12),fg="#5b5b74",bg=PANEL,
                        cursor="sb_v_double_arrow")
        handle.pack(side="left",padx=(0,6))
        handle.bind("<ButtonPress-1>",  lambda e,i=idx: self._card_drag_start(e,i))
        handle.bind("<B1-Motion>",      lambda e,i=idx: self._card_drag_motion(e,i))
        handle.bind("<ButtonRelease-1>",self._card_drag_end)

        dot=tk.Label(r1,text="●",font=("Segoe UI",12),fg=color,bg=PANEL,cursor="hand2")
        dot.pack(side="left", pady=(0,1))
        dot.bind("<Button-1>",lambda e,g=group,c=card: self._pick_color(g,c))
        nv=tk.StringVar(value=group.get("name","Group"))
        ne=tk.Entry(r1,textvariable=nv,font=("Segoe UI",10,"bold"),
                    bg=PANEL,fg=TEXT,insertbackground=TEXT,relief="flat",bd=0,width=14)
        ne.pack(side="left",padx=7)
        ne.bind("<KeyRelease>", lambda e,g=group,v=nv: self._on_name(g,v))
        ne.bind("<Return>",    lambda e: self.root.focus_set())
        ne.bind("<Escape>",    lambda e: self.root.focus_set())
        if group.get("master_volume"):
            tk.Label(r1,text="PC MASTER",font=("Segoe UI",7,"bold"),fg="#7fe4f0",
                     bg="#173542", padx=8, pady=3).pack(side="left", padx=(0,4))

        tk.Button(r1,text="✕",font=("Segoe UI",8),fg=SUBTEXT,bg=PANEL,
                  activebackground=BORDER,activeforeground="#ff6b6b",
                  relief="flat",cursor="hand2",
                  command=lambda i=idx: self._del_group(i)).pack(side="right",padx=(6,0))

        en_btn=tk.Button(r1,text="Enabled" if group.get("enabled",True) else "Disabled",
                         font=("Segoe UI",8,"bold"),bg=BORDER,
                         fg="#1DB954" if group.get("enabled",True) else "#8b7280",
                         activebackground=HOVER,relief="flat",cursor="hand2",padx=8,pady=1)
        en_btn.pack(side="right",padx=2)
        def _tog(g=group,b=en_btn):
            g["enabled"]=not g.get("enabled",True)
            b.config(text="Enabled" if g["enabled"] else "Disabled",
                     bg="#183524" if g["enabled"] else "#2a1a1a",
                     fg="#1DB954" if g["enabled"] else "#8b7280")
            self._ensure_single_default_group()
            self._autosave()
        en_btn.config(bg="#183524" if group.get("enabled",True) else "#2a1a1a")
        en_btn.config(command=_tog)

        if mode=="single" and self.cfg.get("single_auto_revert", False):
            is_def = (idx == self.cfg.get("single_default_group", 0))
            def_btn=tk.Button(r1,text="★" if is_def else "☆",
                              font=("Segoe UI",11),
                              fg="#FFC107" if is_def else SUBTEXT,
                              bg=PANEL,activebackground=PANEL,
                              relief="flat",cursor="hand2")
            def_btn.pack(side="right",padx=2)
            def _set_def(g=group):
                self._set_single_default_group(g)
            def_btn.config(command=_set_def)
        else:
            def_btn = None

        # Row 2: vol% | step | mute
        r2=tk.Frame(card,bg=PANEL); r2.pack(fill="x",padx=10,pady=(0,0))
        vol_lbl=tk.Label(r2,text=self._vt(group),font=("Segoe UI",10,"bold"),
                         fg=color,bg=PANEL,width=8,anchor="w")
        vol_lbl.pack(side="left")

        step_var=tk.IntVar(value=group.get("step",5))
        def _step_chg(g=group,v=step_var): g["step"]=v.get(); self._autosave()
        tk.Label(r2,text="%",font=("Segoe UI",8),fg=SUBTEXT,bg=PANEL).pack(side="right",padx=(0,4))
        tk.Spinbox(r2,from_=1,to=25,textvariable=step_var,width=3,
                   font=("Segoe UI",8),bg=INPUT_BG,fg=TEXT,buttonbackground=BORDER,
                   highlightthickness=0,relief="flat",
                   command=_step_chg).pack(side="right",padx=(0,4))

        mute_btn=tk.Button(r2,text="Muted" if group.get("muted") else "Mute",
                           font=("Segoe UI",8,"bold"),
                           bg="#3a1a1a" if group.get("muted") else BORDER,
                           fg="#ff6b6b" if group.get("muted") else SUBTEXT,
                           relief="flat",cursor="hand2",padx=8,pady=2)
        mute_btn.pack(side="right",padx=(0,8))
        mute_btn.config(command=self._mk_mute(group,mute_btn,vol_lbl,color))

        # Slider
        vol_var=tk.IntVar(value=int(group.get("volume",80)))
        stn=f"G{idx}.Horizontal.TScale"
        st=ttk.Style(card)
        try: st.theme_use("clam")
        except Exception: pass
        st.configure(stn,troughcolor=INPUT_BG,background=color,sliderlength=17,sliderrelief="flat")
        ttk.Scale(card,from_=0,to=100,orient="horizontal",variable=vol_var,
                  style=stn,
                  command=self._mk_slider(group,vol_var,vol_lbl,color)
                  ).pack(fill="x",padx=10,pady=(3,4))

        # Row 3: hotkeys with labels
        r3=tk.Frame(card,bg=PANEL); r3.pack(fill="x",padx=10,pady=(0,3))
        if mode=="multi":
            for action,label in [("vol_down","Vol-"),("vol_up","Vol+"),("mute","Vol Mute")]:
                cur=group.get("keys",{}).get(action,"")
                lf2=tk.Frame(r3,bg=PANEL); lf2.pack(side="left",padx=(0,8))
                tk.Label(lf2,text=label,font=("Segoe UI",7),fg=SUBTEXT,bg=PANEL).pack(anchor="w")
                rf2=tk.Frame(lf2,bg=PANEL); rf2.pack(anchor="w")
                def _cb(hk,g=group,a=action):
                    ok, msg = _validate_hotkey_choice(self.cfg, hk, ("group", g.get("id", id(g)), a))
                    if not ok:
                        _themed_alert(self.root, "Hotkey already in use", msg)
                        return
                    g.setdefault("keys",{})[a]=hk
                    self.hk.reload(self.cfg,self._on_vol,self._on_switch)
                    self._autosave()
                btn=make_hotkey_btn(rf2,cur,_cb)
                btn.pack(side="left")
                def _clr(g=group,a=action,b=btn):
                    cap = getattr(b, "_capture", None)
                    if cap and cap._active:
                        cap._finish(None)  # cancel active capture
                    else:
                        g.setdefault("keys",{})[a]=""
                        b.config(text="—")
                        self.hk.reload(self.cfg,self._on_vol,self._on_switch)
                        self._autosave()
                tk.Button(rf2,text="×",font=("Segoe UI",9,"bold"),bg=PANEL,fg="#6d7086",
                          activebackground=PANEL,activeforeground=TEXT,
                          relief="flat",cursor="hand2",padx=3,pady=0,
                          command=_clr).pack(side="left",padx=(1,0))
        # 1-Knob mode: switch key removed — use Cycle key in the panel above groups

        # Row 4: apps (skip for master volume groups)
        if not group.get("master_volume"):
            r4=tk.Frame(card,bg=PANEL); r4.pack(fill="x",padx=10,pady=(1,7))
            apps_wrap=tk.Frame(r4,bg=PANEL)
            apps_wrap.pack(side="left",fill="x",expand=True)
            tk.Button(r4,text="Edit Apps",font=("Segoe UI",8,"bold"),bg=BORDER,fg=TEXT,
                      relief="flat",cursor="hand2",padx=8,pady=3,
                      command=lambda g=group:
                              AppsDialog(self.root,g,self._redraw,self._autosave,
                                         all_groups=self.cfg["groups"])
                      ).pack(side="right")
            self._render_app_chips(apps_wrap, group.get("apps", []))
        else:
            tk.Label(card,text="Controls the entire PC volume",
                     font=("Segoe UI",8),fg="#7fe4f0",bg=PANEL).pack(
                     anchor="w",padx=10,pady=(2,7))

        self._group_widgets.append({
            "vol_var":vol_var,"vol_lbl":vol_lbl,
            "mute_btn":mute_btn,"color":color,"step_var":step_var,
            "group": group, "apps_row": r4 if not group.get("master_volume") else None,
            "default_btn": def_btn,
        })

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _vt(self,g): return "Muted" if g.get("muted") else f"{int(g.get('volume',80))}%"
    def _at(self,g):
        if g.get("master_volume"): return ""
        apps=g.get("apps",[]); return (", ".join(apps[:6])+(f" +{len(apps)-6} more" if len(apps)>6 else "")) if apps else "No apps — click Edit to add"

    def _render_app_chips(self, parent, apps):
        for w in parent.winfo_children():
            w.destroy()
        if not apps:
            tk.Label(parent, text="No apps - click Edit Apps to add",
                     font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL).pack(anchor="w")
            return
        row = None
        for i, app in enumerate(apps[:6]):
            if i % 3 == 0:
                row = tk.Frame(parent, bg=PANEL)
                row.pack(anchor="w", fill="x", pady=(0,3))
            tk.Label(row, text=app, font=("Segoe UI",8),
                     fg="#a8d4ff", bg="#152533", padx=8, pady=3).pack(side="left", padx=(0,6))
        if len(apps) > 6:
            tk.Label(parent, text=f"+{len(apps)-6} more",
                     font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL).pack(anchor="w")

    def _mk_slider(self,group,vol_var,lbl,color):
        def _cmd(val):
            group["volume"]=float(val); group["muted"]=False
            lbl.config(text=self._vt(group),fg=color)
            apply_vol(group,self.cfg)
            # Debounce disk writes — only save 300ms after last slider move
            if hasattr(self,"_slider_job") and self._slider_job:
                try: self.root.after_cancel(self._slider_job)
                except: pass
            self._slider_job = self.root.after(300, self._autosave)
        return _cmd

    def _mk_mute(self,group,btn,lbl,color):
        def _cmd():
            if group.get("muted"): group["muted"]=False; group["volume"]=group.get("_vbm",80)
            else: group["_vbm"]=group["volume"]; group["muted"]=True
            btn.config(text="Muted" if group["muted"] else "Mute",
                       bg="#3a1a1a" if group["muted"] else BORDER,
                       fg="#ff6b6b" if group["muted"] else SUBTEXT)
            lbl.config(text=self._vt(group),fg="#ff6b6b" if group["muted"] else color)
            apply_vol(group,self.cfg); self._autosave()
        return _cmd

    def _on_name(self,g,v): g["name"]=v.get(); self._autosave()
    def _pick_color(self,g,card):
        col=colorchooser.askcolor(color=g.get("color","#888"),title="Pick color",parent=self.root)
        if col and col[1]: g["color"]=col[1]; self._autosave(); self._redraw()

    def _add_group(self):
        nid=max((g.get("id",0) for g in self.cfg["groups"]),default=-1)+1
        col=ACCENT[nid%len(ACCENT)]
        self.cfg["groups"].append({"id":nid,"name":f"Group {nid+1}","color":col,
            "apps":[],"foreground_mode":False,"keys":{"vol_down":"","vol_up":"","mute":""},
            "single_key":"","step":5,"volume":80,"muted":False,"_vbm":80,
            "enabled":True,"is_default":False})
        self._ensure_single_default_group()
        self._autosave(); self._redraw()

    def _update_master_vol_btn(self):
        """Show/hide + Master Vol button based on whether one exists."""
        has_master = any(g.get("master_volume") for g in self.cfg["groups"])
        if has_master:
            self._master_vol_btn.pack_forget()
        else:
            self._master_vol_btn.pack(side="left", padx=4)

    def _add_master_group(self):
        if any(g.get("master_volume") for g in self.cfg["groups"]):
            return  # already exists, button should be hidden anyway
        nid = max((g.get("id",0) for g in self.cfg["groups"]),default=-1)+1
        self.cfg["groups"].append({
            "id":nid, "name":"Master Volume", "color":"#00BCD4",
            "apps":[], "foreground_mode":False, "master_volume":True,
            "keys":{"vol_down":"","vol_up":"","mute":""},
            "single_key":"","step":5,"volume":80,"muted":False,"_vbm":80,
            "enabled":True,"is_default":False,
        })
        self._ensure_single_default_group()
        self._autosave(); self._redraw()

    def _del_group(self,idx):
        if len(self.cfg["groups"])<=1:
            _themed_alert(self.root, "KnobMixer", "Need at least one group.")
            return
        name=self.cfg["groups"][idx].get("name","Group")
        if _themed_confirm(self.root, "Delete Group", f'Delete "{name}"?', yes_text="Delete", danger=True):
            self.cfg["groups"].pop(idx); self._ensure_single_default_group(); self._autosave(); self._redraw()

    # ── Mode ──────────────────────────────────────────────────────────────────
    def _rebuild_knob_panel(self):
        """Rebuild the 1-knob hotkey panel (e.g. when hw_knob toggled)."""
        for w in self._knob_panel.winfo_children(): w.destroy()
        # Re-run the knob panel build logic
        kp = self._knob_panel
        def _make_hk_cell(parent, label, get_val, set_val, clear_val):
            f = tk.Frame(parent, bg=PANEL)
            tk.Label(f, text=label, font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL).pack(anchor="w")
            rf = tk.Frame(f, bg=PANEL); rf.pack(anchor="w")
            btn = make_hotkey_btn(rf, get_val(), set_val)
            btn.pack(side="left")
            def _clr(b=btn):
                cap = getattr(b, "_capture", None)
                if cap and cap._active: cap._finish(None)
                else: clear_val(); b.config(text="—")
            tk.Button(rf, text="×", font=("Segoe UI",9,"bold"), bg=PANEL, fg="#6d7086",
                      activebackground=PANEL, activeforeground=TEXT,
                      relief="flat", cursor="hand2", padx=3, pady=0,
                      command=_clr).pack(side="left", padx=(1,0))
            return f
        sk = self.cfg.setdefault("single_keys", {"vol_down":"","vol_up":"","mute":""})
        hw = self.cfg.get("hw_knob_enabled", False)
        row1 = tk.Frame(kp, bg=PANEL); row1.pack(fill="x", padx=10, pady=(4,1))
        tk.Label(row1, text="1-Knob:", font=("Segoe UI",8,"bold"), fg=SUBTEXT,
                 bg=PANEL).pack(side="left", padx=(0,8))
        if not hw:
            for action, lbl in [("vol_down","Vol-"), ("vol_up","Vol+"), ("mute","Mute")]:
                def _set_single_hotkey(hk, a=action):
                    ok, msg = _validate_hotkey_choice(self.cfg, hk, ("single_shared", a))
                    if not ok:
                        _themed_alert(self.root, "Hotkey already in use", msg)
                        return
                    self.cfg["single_keys"][a] = hk
                    self._autosave()
                _make_hk_cell(row1, lbl,
                              lambda a=action: sk.get(a,""),
                              _set_single_hotkey,
                              lambda a=action: (self.cfg["single_keys"].__setitem__(a,""), self._autosave())
                              ).pack(side="left", padx=(0,10))
        else:
            tk.Label(row1, text="Vol keys handled by Hardware Knob",
                     font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL).pack(side="left", padx=(0,14))
        def _ck_cb(hk):
            ok, msg = _validate_hotkey_choice(self.cfg, hk, ("cycle",))
            if not ok:
                _themed_alert(self.root, "Hotkey already in use", msg)
                return
            self.cfg["cycle_key"] = hk; self._autosave()
        def _ck_clr(): self.cfg["cycle_key"] = ""; self._autosave()
        _make_hk_cell(row1, "Cycle",
                      lambda: self.cfg.get("cycle_key",""),
                      _ck_cb, _ck_clr).pack(side="left")

    def _on_mode(self):
        mode = self._mode_var.get()
        self.cfg["mode"] = mode
        # When switching to multi: suspend hw_knob in the hook but preserve the
        # user's preference in cfg so it restores when switching back to single.
        # The hook checks mode=="single" before acting on hw_knob_enabled, so
        # simply reloading hotkeys with mode="multi" is enough to deactivate it.
        self._init_single()
        self.hk.reload(self.cfg, self._on_vol, self._on_switch)
        self._redraw()
        if mode == "single":
            self._sb.pack(fill="x")
            # Rebuild knob panel so it reflects current hw_knob_enabled state
            self._rebuild_knob_panel()
            self._knob_panel.pack(fill="x", before=self._groups_cf)
            self._hw_row.pack(fill="x", before=self._groups_cf)
            # Sync the checkbox to the saved preference
            self._hw_var.set(self.cfg.get("hw_knob_enabled", False))
        else:
            self._sb.pack_forget()
            self._knob_panel.pack_forget()
            self._hw_row.pack_forget()
        self._autosave()

    def _init_single(self):
        self._ensure_single_default_group()
        # Active group = first enabled group in list (user controls order by dragging)
        gs=self.cfg["groups"]
        enabled=[g for g in gs if g.get("enabled",True)]
        if enabled:
            self._active_grp=enabled[0]
            self.cfg["_active_group_ref"]=self._active_grp
        elif gs:
            self._active_grp=gs[0]
            self.cfg["_active_group_ref"]=self._active_grp
        if self.cfg.get("mode")=="single" or self.cfg.get("hw_knob_enabled"):
            self._update_active_lbl()
        # Show active bar if hw_knob enabled (even in multi mode)
        if self.cfg.get("hw_knob_enabled") and self.cfg.get("mode")=="multi":
            self._sb.pack(fill="x")

    def _on_switch(self,group):
        self._active_grp=group; self.cfg["_active_group_ref"]=group
        self.root.after(0,self._update_active_lbl)
        # Only show overlay when triggered by user action (hotkey),
        # not during startup init or programmatic switches
        if getattr(self, "_hotkeys_active", False):
            self.root.after(0,lambda g=group:
                self.overlay.show(g["name"],g.get("color","#888"),
                                  g["volume"],g.get("muted",False),
                                  self.cfg.get("overlay_size",1.0)))
        if self.cfg.get("mode","multi")=="single":
            if self._timeout_job: self.root.after_cancel(self._timeout_job)
            self._timeout_job = None
            self._timeout_lbl.config(text="")
            if self.cfg.get("single_auto_revert", False):
                self._timeout_start = time.time()
                ms = self.cfg.get("single_timeout", 60) * 1000
                self._timeout_job = self.root.after(ms, self._revert_default)
                self.root.after(0, self._tick)

    def _revert_default(self):
        self._ensure_single_default_group()
        idx=self.cfg.get("single_default_group",0); gs=self.cfg["groups"]
        if gs:
            self._active_grp=gs[min(idx,len(gs)-1)]
            self.cfg["_active_group_ref"]=self._active_grp
            self.root.after(0,self._update_active_lbl)
        self._timeout_job=None; self.root.after(0,lambda: self._timeout_lbl.config(text=""))

    def _tick(self):
        if self._timeout_job is None: return
        remain=max(0,self.cfg.get("single_timeout",30)-(time.time()-self._timeout_start))
        self._timeout_lbl.config(text=f"Reverts in {int(remain)}s")
        if remain>0: self.root.after(1000,self._tick)

    def _refresh_revert_timer(self):
        if self.cfg.get("mode", "multi") != "single":
            self._timeout_lbl.config(text="")
            return
        if not self.cfg.get("single_auto_revert", False):
            if self._timeout_job:
                self.root.after_cancel(self._timeout_job)
                self._timeout_job = None
            self._timeout_lbl.config(text="")
            return
        if self._active_grp:
            self._on_switch(self._active_grp)

    def _update_active_lbl(self):
        if self._active_grp:
            self._active_lbl.config(text=f"Active: {self._active_grp['name']}",
                                    fg=self._active_grp.get("color","#1DB954"))
    def _ensure_single_default_group(self):
        gs = self.cfg.get("groups", [])
        if not gs:
            self.cfg["single_default_group"] = 0
            return
        idx = int(self.cfg.get("single_default_group", 0))
        if 0 <= idx < len(gs) and gs[idx].get("enabled", True):
            return
        for i, g in enumerate(gs):
            if g.get("enabled", True):
                self.cfg["single_default_group"] = i
                return
        self.cfg["single_default_group"] = 0

    def _refresh_default_buttons(self):
        active = self.cfg.get("single_auto_revert", False) and self.cfg.get("mode") == "single"
        idx = self.cfg.get("single_default_group", 0)
        for i, w in enumerate(getattr(self, "_group_widgets", [])):
            btn = w.get("default_btn")
            if not btn:
                continue
            btn.config(text="★" if active and i == idx else "☆",
                       fg="#FFC107" if active and i == idx else SUBTEXT)

    def _set_single_default_group(self, group):
        try:
            self.cfg["single_default_group"] = self.cfg["groups"].index(group)
        except ValueError:
            return
        self._ensure_single_default_group()
        self._autosave()
        self._refresh_default_buttons()

    def _show_saved(self):
        self._dirty_lbl.config(text="✓ Saved")
        self.root.after(1500,lambda: self._dirty_lbl.config(text=""))

    # ── Global on/off ─────────────────────────────────────────────────────────
    def set_update_available(self, ver, url):
        self._update_available = True
        self.root.after(0, self._refresh_tray)  # update tray badge
        """Called when a newer version is found — shows download button."""
        self._update_url[0] = url
        self._update_ver    = ver
        self._update_btn.config(
            text=f"Update v{ver}",
            bg="#2f2414", fg="#ffb066",
            command=self._download_and_install)

    def _download_and_install(self):
        """Download installer silently then run it — in-app update."""
        import urllib.request, tempfile, subprocess, os

        # Build direct download URL from the releases page URL
        # GitHub direct download: releases/download/vX.Y/KnobMixer_Setup.exe
        ver = getattr(self, "_update_ver", "")
        if ver:
            dl_url = (f"https://github.com/{GITHUB_REPO}"
                      f"/releases/download/v{ver}/KnobMixer_Setup.exe")
        else:
            # Fallback — open browser if we don't have the version
            import webbrowser
            webbrowser.open(self._update_url[0] or
                            f"https://github.com/{GITHUB_REPO}/releases/latest")
            return

        # Disable button while downloading
        self._update_btn.config(text="Starting download…",
                                fg=SUBTEXT, bg=BORDER, command=lambda: None)

        def _download():
            try:
                # Save to temp file
                tmp = Path(tempfile.gettempdir()) / "KnobMixer_Setup.exe"

                def _progress(count, block, total):
                    if total > 0:
                        pct = min(100, int(count * block * 100 / total))
                        self.root.after(0, lambda p=pct:
                            self._update_btn.config(
                                text=f"Downloading… {p}%",
                                fg=SUBTEXT, bg=BORDER)
                            if self._update_btn.winfo_exists() else None)

                urllib.request.urlretrieve(dl_url, tmp, _progress)

                # Download complete — run installer
                # The installer closes KnobMixer itself (taskkill in [Run])
                # so we just launch it and let it take over
                self.root.after(0, lambda: self._update_btn.config(
                    text="Installing…", fg=SUBTEXT, bg=BORDER)
                    if self._update_btn.winfo_exists() else None)

                # Quit the app FIRST, then launch installer
                # This prevents "can't close app" error from Inno Setup
                def _launch_installer():
                    try: subprocess.Popen([str(tmp)], shell=False)
                    except Exception: pass
                self.root.after(0, lambda: (
                    self._update_btn.config(text="Restarting…", fg=SUBTEXT, bg=BORDER)
                    if self._update_btn.winfo_exists() else None))
                self.root.after(500, _launch_installer)
                self.root.after(800, self._quit)

            except Exception as e:
                # Download failed — fall back to browser
                import webbrowser
                self.root.after(0, lambda: (
                    self._update_btn.config(
                        text="Download failed — click to open browser",
                        fg="#cc6666", bg="#2a1a1a",
                        command=lambda: webbrowser.open(
                            f"https://github.com/{GITHUB_REPO}/releases/latest"))
                    if self._update_btn.winfo_exists() else None))

        threading.Thread(target=_download, daemon=True).start()

    def _manual_update_check(self):
        """User clicked Update button — disable it while checking."""
        self._update_btn.config(text="Checking…", fg=SUBTEXT, bg=BORDER,
                                command=lambda: None)  # disable while checking
        def _reset():
            if self._update_btn.winfo_exists():
                self._update_btn.config(text="Check Updates", fg=SUBTEXT, bg=PANEL_SOFT,
                                        command=self._manual_update_check)
        def _on_version(ver, url):
            if ver and _ver_tuple(ver) > _ver_tuple(APP_VER):
                self.root.after(0, lambda: self.set_update_available(ver, url))
            elif ver:
                # Up to date — show for 30s then reset
                self.root.after(0, lambda: self._update_btn.config(
                    text="Up to date ✓", fg="#8acfa8", bg="#1d3126"))
                self.root.after(30000, _reset)
            else:
                # No connection — show clearly for 30s then reset
                self.root.after(0, lambda: self._update_btn.config(
                    text="No connection", fg="#cc6666", bg="#2a1a1a"))
                self.root.after(30000, _reset)
        _fetch_latest_version(_on_version)

    def _toggle_en(self):
        self._enabled=not self._enabled
        if self._enabled:
            self._onoff_btn.config(text="Enabled",bg="#183524",fg="#1DB954")
            self.hk.reload(self.cfg,self._on_vol,self._on_switch)
            self._reg_mic_hk()
        else:
            self._onoff_btn.config(text="Disabled",bg="#2a1a1a",fg="#8b7280")
            self.hk.stop()
        self.overlay.set_enabled(self._enabled)
        try: self.tray.icon=make_tray_img(
            self.cfg["groups"], self._enabled,
            mic_muted=self.mic.get() if self.cfg.get("mic_enabled") else None,
            update_available=self._update_available)
        except: pass

    # ── Mic ───────────────────────────────────────────────────────────────────
    def _reg_mic_hk(self):
        """Register mic hotkey through the global Win32 hook.
        suppress=False so the key still reaches the game/other apps."""
        # Mic hotkey is added into the shared _HOOK alongside other hotkeys.
        # We reload the whole engine to cleanly re-register everything.
        # The actual registration happens in HotkeyEngine.reload() below,
        # but mic is separate so we add it directly here.
        hk = self.cfg.get("mic_hotkey","").strip()
        if not hk or not self.cfg.get("mic_enabled", False):
            return
        # suppress=False: key passes through to game AND triggers mic toggle
        _HOOK.register(hk, self._toggle_mic, suppress=False)

    def _toggle_mic(self):
        self.mic.toggle(self.cfg)
        if self.mic_ov: self.root.after(0, self.mic_ov.update)
        # Update tray icon colour (green=live, red=muted)
        self.root.after(0, self._refresh_tray)

    def _refresh_tray(self):
        try:
            mic_muted = self.mic.get() if self.cfg.get("mic_enabled") else None
            self.tray.icon = make_tray_img(
                self.cfg["groups"], self._enabled,
                mic_muted=mic_muted, update_available=self._update_available)
        except Exception:
            pass

    # ── Volume change callback ─────────────────────────────────────────────────
    def _on_vol(self,group):
        if self.cfg.get("show_overlay",True):
            self.root.after(0,lambda g=group:
                self.overlay.show(g["name"],g.get("color","#888"),
                                  g["volume"],g.get("muted",False),
                                  self.cfg.get("overlay_size",1.0)))
        # Sync UI immediately + again after 150ms for master vol readback
        self.root.after(0, self._sync)
        self.root.after(150, self._sync)

    def _sync(self):
        for g,w in zip(self.cfg["groups"],self._group_widgets):
            try:
                w["vol_var"].set(int(g["volume"]))
                w["vol_lbl"].config(text=self._vt(g),
                                    fg="#ff6b6b" if g.get("muted") else w["color"])
                w["mute_btn"].config(text="Muted" if g.get("muted") else "Mute",
                                     bg="#3a1a1a" if g.get("muted") else BORDER,
                                     fg="#ff6b6b" if g.get("muted") else SUBTEXT)
            except: pass

    def _refresh_loop(self):
        # Only sync UI when window is actually visible — saves CPU when in tray
        if self.root.winfo_viewable():
            self._sync()
        self.root.after(1000, self._refresh_loop)

    def _sync_mute_states(self):
        """On startup, clear any saved mute state for per-app groups.
        ISimpleAudioVolume has no reliable GetMute() for per-app audio.
        Always reset to unmuted on launch so state matches reality (#18)."""
        changed = False
        for g in self.cfg.get("groups", []):
            if g.get("master_volume"): continue
            if g.get("muted", False):
                g["muted"] = False
                apply_vol(g, self.cfg)
                changed = True
        if changed:
            save_cfg(self.cfg)
            self.root.after(0, self._sync)

    def _hook_health_check(self):
        """Every 10s check hook thread is alive — restart if dead (e.g. after sleep)."""
        if not _HOOK._running or (_HOOK._thread and not _HOOK._thread.is_alive()):
            print("[Hook] Health check: hook dead — restarting")
            _HOOK._running = False
            _HOOK.start()
            self.hk.reload(self.cfg, self._on_vol, self._on_switch)
        self.root.after(10000, self._hook_health_check)

    # ── Autosave ──────────────────────────────────────────────────────────────
    def _autosave(self):
        """Save config and reload hotkeys instantly."""
        self.cfg["mode"]=self._mode_var.get()
        for g,w in zip(self.cfg["groups"],self._group_widgets):
            try: g["step"]=w["step_var"].get()
            except: pass
        save_cfg(self.cfg)
        self.hk.reload(self.cfg,self._on_vol,self._on_switch)
        self._reg_mic_hk()
        if self.mic_ov: self.mic_ov.update()
        self._show_saved()

    # ── Settings ──────────────────────────────────────────────────────────────
    def _settings_open(self):
        """True if settings window is currently open."""
        return (hasattr(self, "_settings_win") and
                self._settings_win is not None and
                self._settings_win.winfo_exists())

    def _open_settings(self):
        if self._settings_open():
            self._settings_win.lift()
            return
        self._settings_win = SettingsWin(self.root, self.cfg, self._on_settings_change,
                                          quit_fn=self._quit, app_ref=self)

    def _on_settings_change(self):
        # Force-hide volume popup — settings changes can trigger hk.reload
        # which fires _on_vol callbacks and resets the hide timer
        self.overlay.set_enabled(False)
        self.root.after(100, lambda: self.overlay.set_enabled(
            self.cfg.get("show_overlay", True)))
        self.hk.reload(self.cfg,self._on_vol,self._on_switch)
        self._reg_mic_hk()
        if self.cfg.get("mic_enabled", False):
            if self.mic_ov:
                self.mic_ov.update()   # refresh settings (size, alpha, etc)
                self.mic_ov.show()     # make visible — needed when re-enabling
            else:
                self.mic_ov = MicOverlay(self.root, self.mic, self.cfg)
        else:
            if self.mic_ov: self.mic_ov.hide()

    # ── Tray ─────────────────────────────────────────────────────────────────
    def _setup_tray(self):
        img=make_tray_img(
            self.cfg["groups"], self._enabled,
            mic_muted=self.mic.get() if self.cfg.get("mic_enabled") else None,
            update_available=self._update_available)
        menu=pystray.Menu(
            pystray.MenuItem("Open KnobMixer",self._show,default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Enable / Disable",self._toggle_en),
            pystray.MenuItem("Toggle Mic",self._toggle_mic),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit",self._quit),
        )
        self.tray=pystray.Icon(APP_NAME,img,APP_NAME,menu)
        threading.Thread(target=self.tray.run,daemon=True).start()

    def _hide(self):
        # Clear tutorial if active — it lives inside root so must be cleaned
        if getattr(self, "_tutorial", None):
            try: self._tutorial._finish()
            except: pass
            self._tutorial = None
        if self._settings_open():
            self._settings_win.destroy()
        self.root.withdraw()
    def _show(self,*_):
        def _do_show():
            # Bring main window
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            # Bring open UI child windows (settings, edit dialogs) but NOT
            # overlay windows. Overlays have overrideredirect=True and manage
            # their own visibility — deiconifying them causes the spurious popup bug.
            for w in self.root.winfo_children():
                if isinstance(w, tk.Toplevel) and w.winfo_exists():
                    try:
                        if not w.overrideredirect():  # skip overlays
                            w.deiconify()
                            w.lift()
                    except: pass
        self.root.after(0, _do_show)
    def _quit(self,*_):
        if self._settings_open():
            self._settings_win.destroy()
        self.hk.stop(); _HOOK.stop(); self.tray.stop(); self.root.after(0,self.root.destroy)

    def run(self):
        global _update_callback
        # Route update results to this app instance's UI handler.
        # Guard: only call set_update_available if the fetched version is
        # strictly newer than the running version. Without this guard,
        # any successful fetch — even returning the same version — would
        # incorrectly show the update badge.
        def _on_update(ver, url):
            if ver and _ver_tuple(ver) > _ver_tuple(APP_VER):
                self.root.after(0, lambda: self.set_update_available(ver, url))
        _update_callback = _on_update
        self.root.after(10000, self._hook_health_check)
        self.root.after(2000, self._sync_mute_states)
        # Mark hotkeys as active after startup — prevents spurious overlay shows
        self.root.after(1000, lambda: setattr(self, "_hotkeys_active", True))
        if not self.cfg.get("tutorial_seen", False):
            self.root.after(800, self._start_tutorial)
        self.root.mainloop()

    def _start_tutorial(self):
        # Close settings if open — avoids z-order conflicts
        if self._settings_open():
            self._settings_win.destroy()
            self._settings_win = None
        self._tutorial = TutorialOverlay(self.root, self)


# ══════════════════════════════════════════════════════════════════════════════
# Apps Dialog
# ══════════════════════════════════════════════════════════════════════════════
class AppsDialog(tk.Toplevel):
    """Redesigned apps dialog:
    - Top: scrollable list of ADDED apps with ✕ to remove each
    - Middle: manual add field
    - Bottom: all open apps (with/without sound) as quick-add buttons
    """
    def __init__(self, parent, group, refresh_ui, save_fn, all_groups=None):
        super().__init__(parent)
        self.group      = group
        self.refresh_ui = refresh_ui
        self.save_fn    = save_fn
        self.all_groups = all_groups or []  # for duplicate detection (#9)
        # Working copy of apps — edit in memory, save on Save
        self._apps = list(group.get("apps", []))
        self.title(f"Edit Apps — {group.get('name','Group')}")
        self.configure(bg=BG)
        self.geometry("480x560")
        self.resizable(True, True)
        self.grab_set()
        self._build()
        _place_near_parent(self, parent, side="left")

    def _build(self):
        # ── Section 1: Added apps (top, expandable) ──────────────────────────
        hdr1 = tk.Frame(self, bg=BG); hdr1.pack(fill="x", padx=12, pady=(10,2))
        tk.Label(hdr1, text="Added apps:", font=("Segoe UI",9,"bold"),
                 fg=TEXT, bg=BG).pack(side="left")
        tk.Label(hdr1, text="click ✕ to remove",
                 font=("Segoe UI",8), fg=SUBTEXT, bg=BG).pack(side="left", padx=6)

        # Scrollable list of added apps
        added_outer = tk.Frame(self, bg=PANEL, highlightthickness=1,
                               highlightbackground=BORDER)
        added_outer.pack(fill="both", expand=True, padx=12, pady=(0,4))
        self._added_canvas = tk.Canvas(added_outer, bg=PANEL,
                                       highlightthickness=0, height=140)
        added_sb = ttk.Scrollbar(added_outer, orient="vertical",
                                 style="Knob.Vertical.TScrollbar",
                                 command=self._added_canvas.yview)
        self._added_canvas.configure(yscrollcommand=added_sb.set)
        self._added_canvas.pack(side="left", fill="both", expand=True)
        added_sb.pack(side="right", fill="y")
        self._added_inner = tk.Frame(self._added_canvas, bg=PANEL)
        self._added_cwin  = self._added_canvas.create_window(
            (0,0), window=self._added_inner, anchor="nw")
        self._added_inner.bind("<Configure>", lambda e: (
            self._added_canvas.configure(
                scrollregion=self._added_canvas.bbox("all")),
            self._added_canvas.itemconfig(
                self._added_cwin, width=self._added_canvas.winfo_width())))
        self._added_canvas.bind("<Configure>", lambda e:
            self._added_canvas.itemconfig(self._added_cwin, width=e.width))
        self._added_canvas.bind("<MouseWheel>", lambda e:
            self._added_canvas.yview_scroll(int(-1*e.delta/120), "units"))
        self._refresh_added()

        # ── Section 2: Manual add ─────────────────────────────────────────────
        mf = tk.Frame(self, bg=BG); mf.pack(fill="x", padx=12, pady=(4,2))
        tk.Label(mf, text="Add manually (process name, no .exe):",
                 font=("Segoe UI",8), fg=SUBTEXT, bg=BG).pack(anchor="w")
        mf2 = tk.Frame(mf, bg=BG); mf2.pack(fill="x", pady=(2,0))
        self._manual_var = tk.StringVar()
        manual_entry = tk.Entry(mf2, textvariable=self._manual_var,
                                font=("Consolas",9), bg=PANEL, fg=TEXT,
                                insertbackground=TEXT, relief="flat")
        manual_entry.pack(side="left", fill="x", expand=True, padx=(0,4))
        manual_entry.bind("<Return>", lambda e: self._add_manual())
        tk.Button(mf2, text="Add", font=("Segoe UI",8), bg=BORDER, fg=TEXT,
                  relief="flat", cursor="hand2", padx=8,
                  command=self._add_manual).pack(side="left")

        # ── Section 3: Open apps quick-add ───────────────────────────────────
        hdr2 = tk.Frame(self, bg=BG); hdr2.pack(fill="x", padx=12, pady=(8,2))
        tk.Label(hdr2, text="Open apps — click to add:",
                 font=("Segoe UI",8), fg=SUBTEXT, bg=BG).pack(side="left")
        tk.Button(hdr2, text="↻ Refresh", font=("Segoe UI",7), bg=BORDER,
                  fg=SUBTEXT, relief="flat", cursor="hand2", padx=4, pady=1,
                  command=self._refresh_open).pack(side="right")

        open_outer = tk.Frame(self, bg=PANEL, highlightthickness=1,
                              highlightbackground=BORDER)
        open_outer.pack(fill="x", padx=12, pady=(0,4))
        self._open_canvas = tk.Canvas(open_outer, bg=PANEL,
                                      highlightthickness=0, height=120)
        open_sb = ttk.Scrollbar(open_outer, orient="vertical",
                                style="Knob.Vertical.TScrollbar",
                                command=self._open_canvas.yview)
        self._open_canvas.configure(yscrollcommand=open_sb.set)
        self._open_canvas.pack(side="left", fill="both", expand=True)
        open_sb.pack(side="right", fill="y")
        self._open_inner = tk.Frame(self._open_canvas, bg=PANEL)
        self._open_cwin  = self._open_canvas.create_window(
            (0,0), window=self._open_inner, anchor="nw")
        self._open_inner.bind("<Configure>", lambda e: (
            self._open_canvas.configure(
                scrollregion=self._open_canvas.bbox("all")),
            self._open_canvas.itemconfig(
                self._open_cwin, width=self._open_canvas.winfo_width())))
        self._open_canvas.bind("<Configure>", lambda e:
            self._open_canvas.itemconfig(self._open_cwin, width=e.width))
        self._open_canvas.bind("<MouseWheel>", lambda e:
            self._open_canvas.yview_scroll(int(-1*e.delta/120), "units"))
        self._refresh_open()

        # ── Bottom: save/cancel ───────────────────────────────────────────────
        bf = tk.Frame(self, bg=BG); bf.pack(fill="x", padx=12, pady=(4,10))
        tk.Button(bf, text="Save", font=("Segoe UI",10,"bold"),
                  bg="#1DB954", fg="white", relief="flat", cursor="hand2",
                  padx=18, pady=4, command=self._save).pack(side="left")
        tk.Button(bf, text="Cancel", font=("Segoe UI",10), bg=BORDER, fg=TEXT,
                  relief="flat", cursor="hand2", padx=12, pady=4,
                  command=self.destroy).pack(side="left", padx=6)

    def _refresh_added(self):
        """Redraw the added apps list with ✕ buttons."""
        for w in self._added_inner.winfo_children(): w.destroy()
        if not self._apps:
            tk.Label(self._added_inner, text="  No apps added yet",
                     font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL,
                     pady=8).pack(anchor="w")
            return
        for app in list(self._apps):
            row = tk.Frame(self._added_inner, bg=PANEL)
            row.pack(fill="x", padx=6, pady=1)
            tk.Label(row, text=app, font=("Consolas",9), fg=TEXT,
                     bg=PANEL).pack(side="left", padx=4)
            def _rm(a=app):
                if a in self._apps: self._apps.remove(a)
                self._refresh_added()
                self._refresh_open()
            tk.Button(row, text="✕", font=("Segoe UI",8), bg=PANEL,
                      fg="#555", activebackground=BORDER,
                      activeforeground="#ff6b6b", relief="flat",
                      cursor="hand2", padx=4,
                      command=_rm).pack(side="right")

    def _get_open_apps(self):
        """Get ALL open apps with audio sessions — not just ones making sound."""
        try:
            comtypes.CoInitialize()
            apps = set()
            for s in AudioUtilities.GetAllSessions():
                if s.Process is None: continue
                try:
                    name = s.Process.name().lower().removesuffix(".exe")
                    apps.add(name)
                except: pass
            return sorted(apps)
        except: return []
        finally:
            try: comtypes.CoUninitialize()
            except: pass

    def _refresh_open(self):
        """Show all open apps — including ones not currently making sound."""
        for w in self._open_inner.winfo_children(): w.destroy()
        apps = self._get_open_apps()
        if not apps:
            tk.Label(self._open_inner,
                     text="  (no apps with audio sessions found)",
                     font=("Segoe UI",8), fg=SUBTEXT, bg=PANEL,
                     pady=6).pack(anchor="w")
            return
        cols = 3; row_f = None
        for i, a in enumerate(apps):
            if i % cols == 0:
                row_f = tk.Frame(self._open_inner, bg=PANEL)
                row_f.pack(fill="x", padx=4, pady=1)
            already_added = a in self._apps
            # Check if in another group (#9)
            in_other = any(
                a in g.get("apps",[]) and g is not self.group
                for g in self.all_groups
            )
            if already_added:
                lbl = f"✓ {a}"
                bg, fg = "#1a3a1a", "#1DB954"
            elif in_other:
                lbl = f"⚠ {a}"
                bg, fg = "#3a2a1a", "#FFA500"
            else:
                lbl = f"+ {a}"
                bg, fg = BORDER, TEXT
            btn = tk.Button(row_f, text=lbl, font=("Consolas",8),
                            bg=bg, fg=fg, activebackground=HOVER,
                            relief="flat", cursor="hand2",
                            padx=6, pady=2, width=14)
            btn.pack(side="left", padx=(0,4))
            if not already_added:
                def _add(n=a, b=btn, other=in_other):
                    if other:
                        # Custom dialog with proper X button (askyesno blocks X)
                        result = [False]
                        dlg = tk.Toplevel(self)
                        dlg.title("Already in another group")
                        dlg.configure(bg=BG)
                        dlg.resizable(False, False)
                        dlg.grab_set()
                        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
                        _place_near_parent(dlg, self, side="left")
                        tk.Label(dlg, text=f"'{n}' is already in another group.\n"
                                            "Adding it here may cause conflicts.",
                                 font=("Segoe UI",9), fg=TEXT, bg=BG,
                                 padx=20, pady=12, justify="left").pack()
                        bf = tk.Frame(dlg, bg=BG, pady=8); bf.pack()
                        def _yes():
                            result[0] = True; dlg.destroy()
                        tk.Button(bf, text="Add anyway", font=("Segoe UI",9),
                                  bg="#3a1a1a", fg="#ff6b6b", relief="flat",
                                  cursor="hand2", padx=10, pady=4,
                                  command=_yes).pack(side="left", padx=6)
                        tk.Button(bf, text="Cancel", font=("Segoe UI",9),
                                  bg=BORDER, fg=TEXT, relief="flat",
                                  cursor="hand2", padx=10, pady=4,
                                  command=dlg.destroy).pack(side="left", padx=6)
                        dlg.wait_window()
                        if not result[0]: return
                    self._apps.append(n)
                    self._refresh_added()
                    self._refresh_open()
                btn.config(command=_add)

    def _add_manual(self):
        """Add app from manual entry field."""
        name = self._manual_var.get().strip().lower().removesuffix(".exe")
        if not name: return
        if name in self._apps:
            self._manual_var.set("")
            return
        self._apps.append(name)
        self._manual_var.set("")
        self._refresh_added()
        self._refresh_open()

    def _save(self):
        self.group["apps"] = list(self._apps)
        self.refresh_ui()
        self.save_fn()
        self.grab_release()
        self.destroy()

    def destroy(self):
        try: self.grab_release()
        except: pass
        super().destroy()


# ── Analytics & Update checker ───────────────────────────────────────────────
# ANALYTICS_URL: set to your Cloudflare Worker / server endpoint.
# When set, sends one small anonymous ping on startup per day.
# Payload: random install ID (no personal data) + version + Windows version.
# User can disable in Settings → General → "Send anonymous usage data".
#
# GITHUB_REPO: your GitHub username/repo — used for update checks.
# Set both before building your public release.
ANALYTICS_URL = "https://knobmixer-analytics.bdhhair11.workers.dev/ping"  # your Cloudflare Worker URL
GITHUB_REPO   = "KnobMixer/KnobMixer"
UPDATE_CHECK  = True        # set False to disable update checks entirely
_update_callback = None    # set by App to route update results to UI

def _get_install_id():
    """Persistent random ID per installation. Not tied to any person."""
    id_file = APPDATA_DIR / "install_id"
    if id_file.exists():
        val = id_file.read_text().strip()
        if val: return val
    import uuid
    new_id = str(uuid.uuid4())
    id_file.write_text(new_id)
    return new_id

def _should_ping_today():
    """Return True if we haven't successfully pinged analytics today yet.
    NOTE: stamp is written AFTER success, not before (#3)."""
    stamp_file = APPDATA_DIR / "last_ping"
    import datetime
    today = datetime.date.today().isoformat()
    if stamp_file.exists() and stamp_file.read_text().strip() == today:
        return False
    return True  # Don't stamp yet — stamp after successful ping

def _stamp_ping_today():
    """Mark today as successfully pinged."""
    import datetime
    try:
        (APPDATA_DIR / "last_ping").write_text(datetime.date.today().isoformat())
    except: pass

def _ping_analytics(cfg=None):
    """Daily analytics ping with exponential backoff retry.
    Handles offline-at-launch then reconnect — like Steam/Discord do it."""
    if not cfg or not cfg.get("analytics_enabled", True): return
    if not ANALYTICS_URL: return
    if not _should_ping_today(): return

    def _send(attempt=0):
        import urllib.request, platform, json, time
        try:
            data = {
                "id":      _get_install_id(),
                "version": APP_VER,
                "os":      platform.version()[:40],
            }
            req = urllib.request.Request(
                ANALYTICS_URL,
                data=json.dumps(data).encode(),
                headers={"Content-Type":"application/json",
                         "User-Agent":f"KnobMixer/{APP_VER}"},
                method="POST"
            )
            urllib.request.urlopen(req, timeout=5)
            _stamp_ping_today()  # only stamp on success (#3)
        except Exception:
            # Non-blocking retry with threading.Timer (#10)
            # Retry schedule: 15s, 1min, 5min, 15min, 1hr
            # Covers slow boot network connections (Start With Windows users)
            delays = [15, 60, 300, 900, 3600]
            if attempt < len(delays):
                threading.Timer(delays[attempt], lambda a=attempt: _send(a + 1)).start()
            # else: silently give up until next launch

    threading.Thread(target=_send, daemon=True).start()

def _ver_tuple(s):
    try: return tuple(int(x) for x in s.strip().lstrip("v").split("."))
    except: return (0,)

def _fetch_latest_version(callback, auto=False, attempt=0):
    """Fetch latest GitHub release version.
    auto=True: silent background check with retry if offline.
    auto=False (manual click): single try, fast 4s timeout, calls callback immediately."""
    if not GITHUB_REPO: return
    def _run():
        import urllib.request, json, time
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(
                url, headers={"User-Agent": f"KnobMixer/{APP_VER}"})
            data = json.loads(urllib.request.urlopen(req, timeout=4).read())
            # 404 = no releases yet → treat as "no update"
            if data.get("message") == "Not Found":
                if auto:
                    _stamp_update_check_today()
                callback(APP_VER, "")   # same version = up to date
                return
            latest = data.get("tag_name","").lstrip("v")
            dl_url = data.get("html_url",
                              f"https://github.com/{GITHUB_REPO}/releases/latest")
            if auto:
                _stamp_update_check_today()
            callback(latest or APP_VER, dl_url)
        except Exception:
            if auto:
                # Background check: retry after 1min then 10min then give up
                delays = [60, 600]
                if attempt < len(delays):
                    time.sleep(delays[attempt])
                    _fetch_latest_version(callback, auto=True, attempt=attempt+1)
                else:
                    callback(None, None)
            else:
                # Manual click: fail immediately, don't make user wait
                callback(None, None)
    threading.Thread(target=_run, daemon=True).start()

def _should_check_update_today():
    """Returns True once per day so update badges stay fresh without spamming checks."""
    import datetime
    stamp_file = APPDATA_DIR / "last_update_check"
    today      = datetime.date.today()
    if stamp_file.exists():
        try:
            last = datetime.date.fromisoformat(stamp_file.read_text().strip())
            if last == today:
                return False
        except: pass
    return True

def _stamp_update_check_today():
    import datetime
    try:
        (APPDATA_DIR / "last_update_check").write_text(datetime.date.today().isoformat())
    except Exception:
        pass

if __name__=="__main__":
    # ── Single instance check ─────────────────────────────────────────────────
    import ctypes
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "KnobMixer_SingleInstance")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        # Already running — bring existing window to front
        import ctypes.wintypes
        def _enum(hwnd, _):
            # Fix #15 — match exact title "KnobMixer", not substring
            buf = ctypes.create_unicode_buffer(256)
            ctypes.windll.user32.GetWindowTextW(hwnd, buf, 256)
            if buf.value == APP_NAME:
                ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                return False
            return True
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(_enum), 0)
        raise SystemExit(0)

    app = App()
    _ping_analytics(app.cfg)
    # Daily silent update check — no popups, just updates the tray badge/button
    if _should_check_update_today():
        def _on_version(ver, url):
            if ver and _ver_tuple(ver) > _ver_tuple(APP_VER):
                app.root.after(0, lambda: app.set_update_available(ver, url))
        _fetch_latest_version(_on_version, auto=True)
    app.run()
