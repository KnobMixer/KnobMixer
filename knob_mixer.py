"""
KnobMixer v2.1
All fixes applied:
- Instant autosave on every change
- Inline hotkey capture (no popup, supports combos, 2s listen window)
- Scrollable group cards so nothing is cut off
- Mic icon: click-through, transparency slider, mute-slash inside circle
- Mic sounds: short, two distinct tones, no queuing
- Sound picker with live preview, stays open
- Mic icon style chooser (multiple icons incl. fun ones)
- Labels on hotkey buttons (Vol-, Vol+, Vol Mute)
- Settings stays open on Apply
- BUILD fix: pyinstaller via python -m PyInstaller
"""

import sys, os, json, threading, winreg, ctypes, time, math, struct, wave, io, copy
from pathlib import Path

def _can_import(mod):
    try: __import__(mod); return True
    except ImportError: return False

def _ensure_deps():
    import subprocess
    needed = {"pycaw":"pycaw","comtypes":"comtypes","keyboard":"keyboard",
              "pystray":"pystray","PIL":"Pillow","psutil":"psutil"}
    missing = [pkg for mod,pkg in needed.items() if not _can_import(mod)]
    if missing:
        import tkinter as tk, tkinter.messagebox as mb
        r=tk.Tk(); r.withdraw()
        mb.showinfo("KnobMixer","Installing components (one-time ~30s)…"); r.destroy()
        for pkg in missing:
            subprocess.check_call([sys.executable,"-m","pip","install",pkg,"-q"],
                                  creationflags=subprocess.CREATE_NO_WINDOW)
_ensure_deps()

# ── Crash logging ─────────────────────────────────────────────────────────────
import traceback as _tb

def _setup_crash_log():
    """Redirect unhandled exceptions to a crash log in %APPDATA%\KnobMixer."""
    log_dir = Path(os.getenv("APPDATA",".")) / "KnobMixer"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "crash.log"
    _orig = sys.excepthook
    def _hook(exc_type, exc_val, exc_tb):
        try:
            import datetime
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"KnobMixer v{APP_VER if 'APP_VER' in dir() else '?'} crash — {datetime.datetime.now()}\n")
                f.write(_tb.format_exception(exc_type, exc_val, exc_tb)[-1])
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
    if not trigger: return None, None
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


# Single global hook instance shared by everything
_HOOK = GlobalHookManager()

# ── Constants ─────────────────────────────────────────────────────────────────
APP_NAME    = "KnobMixer"
APP_VER     = "2.2"
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

# Sound presets: (display_name, shape, mute_params, unmute_params)
# Params: (freq, dur_ms, vol_scale)
# Mute = lower/darker tone, Unmute = brighter/higher tone
SOUND_PRESETS = [
    ("Chime",      "bell",   (523, 95,  1.0), (784, 85,  1.0)),  # gentle warm bell
    ("Ding",       "bell",   (392, 75,  1.0), (880, 65,  1.0)),  # lower ding → bright ping
    ("Marimba",    "marimba",(330, 80,  1.0), (523, 75,  1.0)),  # wooden marimba bar hit
    ("Soft Ping",  "ping",   (660, 60,  1.0), (990, 55,  1.0)),  # light ping notification
    ("Glass",      "glass",  (880,100,  1.0),(1175,90,  1.0)),  # glass tap, crystal
    ("Synth",      "synth",  (220, 70,  1.0), (440, 65,  1.0)),  # soft synth pad hit
]

DEFAULT_CFG = {
    "version": 4,
    "mode": "multi",
    "start_minimized": True,
    "show_overlay": True,
    "analytics_enabled": True,
    "overlay_size": 1.0,
    "slowdown_enabled": True,
    "slowdown_threshold": 10,
    "slowdown_step": 0.5,
    "single_default_group": 0,
    "single_timeout": 60,
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
    "mic_sound_volume": 0.15,
    "mic_sound_preset": 0,
    "mic_icon_x": -1,
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
         "apps":["chrome","spotify","firefox","msedge","vlc","opera","brave"],
         "keys":{"vol_down":"","vol_up":"","mute":""},
         "single_key":"","step":5,"volume":80,"muted":False,"_vbm":80,
         "foreground_mode":False,"enabled":True,"is_default":False},
        {"id":2,"name":"Chat","color":"#5865F2",
         "apps":["discord","teams","slack","zoom","skype","telegram"],
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
            for k,v in [("overlay_size",1.0),("slowdown_enabled",True),
                        ("slowdown_threshold",10),("slowdown_step",0.5),
                        ("single_default_group",0),("single_timeout",60),("single_auto_revert",False),("hw_knob_enabled",False),("hw_knob_group",0),("cycle_key",""),
                        ("single_keys",{"vol_down":"","vol_up":"","mute":""}),
                        ("mic_enabled",True),("mic_device",""),("mic_hotkey","f9"),
                        ("mic_start_muted",False),("mic_sound_volume",0.6),
                        ("mic_sound_preset",0),("mic_icon_x",-1),("mic_icon_y",-1),
                        ("mic_icon_size",40),("mic_icon_alpha",0.85),
                        ("mic_icon_style","circle"),("mode","multi"),
                        ("start_minimized",True),("show_overlay",True),("analytics_enabled",True)]:
                d.setdefault(k,v)
            return d
        except: pass
    return copy.deepcopy(DEFAULT_CFG)

def save_cfg(cfg):
    cfg.pop("_active_group_ref", None)
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))

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
    """Play sound immediately, cancelling any queued sound."""
    global _sound_thread
    import winsound, tempfile
    data = _make_wav(freq, dur_ms, vol, shape)
    tmp  = Path(tempfile.mktemp(suffix=".wav"))
    tmp.write_bytes(data)
    def _play():
        with _sound_lock:
            try: winsound.PlaySound(str(tmp), winsound.SND_FILENAME)
            finally:
                try: tmp.unlink()
                except: pass
    # Cancel previous if still going (best-effort: just start new thread)
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
        if self._active: return
        self._active = True
        self._held.clear()
        self._btn.config(text="Press keys…", bg="#2a4a2a", fg="#1DB954")
        keyboard.hook(self._on_event, suppress=True)

    def _on_event(self, ev):
        if not self._active: return
        name = ev.name.lower()

        if ev.event_type == "down":
            if name == "escape":
                self._finish(None)
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
        keyboard.unhook_all()
        if combo:
            self._btn.after(0, lambda: self._btn.config(
                text=fmt_hotkey(combo), bg=BORDER, fg=TEXT))
            self._cb(combo)
        else:
            self._btn.after(0, lambda: self._btn.config(
                text=self._orig, bg=BORDER, fg=TEXT))

def fmt_hotkey(raw):
    if not raw: return "—"
    return "+".join(p.strip().upper() for p in raw.split("+") if p.strip())

def make_hotkey_btn(parent, current_key, callback, label_prefix=""):
    """Create a hotkey button with inline capture. Returns the button."""
    display = fmt_hotkey(current_key) if current_key else "—"
    text    = f"{label_prefix}{display}" if label_prefix else display
    btn = tk.Button(parent, text=text,
                    font=("Consolas",8), bg=BORDER, fg=TEXT,
                    activebackground=HOVER, relief="flat",
                    cursor="hand2", padx=8, pady=2)
    HotkeyCapture(btn, callback, text)
    return btn

# ── Audio ─────────────────────────────────────────────────────────────────────
_audio_lock = threading.Lock()

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

def _find_game_process():
    """Find a running game by checking process priority.
    Games typically run at HIGH or ABOVE_NORMAL priority."""
    try:
        import psutil as _ps
        _skip = {"knobmixer","explorer","svchost","system","csrss","winlogon",
                 "services","lsass","smss","wininit","fontdrvhost","dwm",
                 "searchhost","runtimebroker","discord","chrome","firefox",
                 "msedge","opera","brave","spotify","teams","slack","zoom",
                 "python","pythonw","conhost","cmd","powershell"}
        HIGH_PRIORITY  = 0x00000080
        ABOVE_NORMAL   = 0x00008000
        for proc in _ps.process_iter(["name","pid"]):
            try:
                name = proc.info["name"].lower().removesuffix(".exe")
                if name in _skip: continue
                # Check Windows process priority class
                handle = ctypes.windll.kernel32.OpenProcess(0x0400, False, proc.info["pid"])
                if handle:
                    pc = ctypes.windll.kernel32.GetPriorityClass(handle)
                    ctypes.windll.kernel32.CloseHandle(handle)
                    if pc in (HIGH_PRIORITY, ABOVE_NORMAL):
                        return name
            except: continue
    except: pass
    return None

def _calc_vol(current, delta, cfg):
    thr  = cfg.get("slowdown_threshold", 10)
    fine = cfg.get("slowdown_step", 0.5)
    if cfg.get("slowdown_enabled", True):
        if delta < 0 and current <= thr:
            return max(0.0, current - fine)
        if delta > 0 and current < thr:
            return min(100.0, current + fine)
    return max(0.0, min(100.0, current + delta))

def apply_vol(group, cfg):
    if not group.get("enabled", True): return
    def _w():
        with _audio_lock:
            scalar = 0.0 if group.get("muted") else group["volume"]/100.0
            # Master volume group — controls Windows master output volume
            if group.get("master_volume", False):
                try:
                    comtypes.CoInitialize()
                    # pycaw GetSpeakers() returns an AudioDevice wrapper.
                    # Access the raw COM IMMDevice via ._dev, then Activate
                    # to get IAudioEndpointVolume — the correct pycaw pattern.
                    dev = AudioUtilities.GetSpeakers()
                    # Try ._dev (raw COM ptr) first, fall back to direct call
                    raw = getattr(dev, "_dev", dev)
                    iface = raw.Activate(
                        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    ep = iface.QueryInterface(IAudioEndpointVolume)
                    if group.get("muted"):
                        ep.SetMute(1, None)
                    else:
                        ep.SetMute(0, None)
                        ep.SetMasterVolumeLevelScalar(float(scalar), None)
                    # Sync slider to actual Windows volume
                    group["volume"] = round(
                        ep.GetMasterVolumeLevelScalar() * 100)
                except Exception as e:
                    print(f"[Master vol error] {e}")
                return
            apps = list(group.get("apps",[]))
            # No foreground mode — apps must be added manually
            sess = _sessions()
            for app in apps:
                for vc in sess.get(app.lower(),[]):
                    try: vc.SetMasterVolume(scalar,None)
                    except: pass
    threading.Thread(target=_w, daemon=True).start()

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
                                    # Found it — build device ID and activate
                                    dev_id = "{0.0.1.00000000}." + "{" + guid_key + "}"
                                    import comtypes.client
                                    enumerator = comtypes.client.CreateObject(
                                        "{BCDE0395-E52F-467C-8E3D-C4579291692E}",
                                        interface=comtypes.IUnknown)
                                    # Use pycaw to activate by scanning sessions
                                    mic = AudioUtilities.GetMicrophone()
                                    if mic:
                                        raw = getattr(mic,"_dev",mic)
                                        iface = raw.Activate(
                                            IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
                                        return iface.QueryInterface(IAudioEndpointVolume)
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
    except Exception as e:
        print(f"Mic enum error: {e}")
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

    def sync(self):
        try:
            ep = self._ep()
            if ep: self._muted = bool(ep.GetMute())
        except: pass

    def get(self): return self._muted

    def set(self, muted, cfg):
        with self._lock:
            self._muted = muted
            def _w():
                try:
                    ep = self._ep(cfg)
                    if ep: ep.SetMute(1 if muted else 0, None)
                except: pass
            threading.Thread(target=_w, daemon=True).start()
            vol = cfg.get("mic_sound_volume", 0.6)
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
class VolumeOverlay:
    def __init__(self, root):
        self._root=root; self._win=None; self._job=None

    def _force_topmost(self):
        """Force window above fullscreen games. HWND_TOPMOST + NOACTIVATE
        so the game never loses focus. Do NOT touch WS_EX_LAYERED —
        tkinter already manages that for -alpha; touching it causes black square."""
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

    def show(self, name, color, volume, muted, scale=1.0):
        if self._win is None or not self._win.winfo_exists():
            self._win=tk.Toplevel(self._root)
            self._win.overrideredirect(True)
            self._win.attributes("-topmost", True)
            self._win.attributes("-alpha", 0.92)
            self._win.configure(bg="#12121e")
            # Apply Win32 flags so it shows over fullscreen games
            self._win.after(10, self._force_topmost)
        for w in self._win.winfo_children(): w.destroy()
        text = "MUTED" if muted else f"{int(volume)}%"
        fg   = "#ff6b6b" if muted else color
        pad  = max(8, int(16*scale))
        fsz  = max(10, int(28*scale))
        nsz  = max(8, int(9*scale))
        bw   = max(60, int(140*scale))
        f=tk.Frame(self._win,bg="#12121e",padx=pad,pady=max(8,int(10*scale))); f.pack()
        tk.Label(f,text=name,font=("Segoe UI",nsz),fg="#888",bg="#12121e").pack()
        tk.Label(f,text=text,font=("Segoe UI",fsz,"bold"),fg=fg,bg="#12121e").pack()
        if not muted:
            bg2=tk.Frame(f,bg="#2a2a3e",height=max(3,int(4*scale)),width=bw)
            bg2.pack(pady=(3,0)); bg2.pack_propagate(False)
            fw=max(2,int(bw*volume/100))
            tk.Frame(bg2,bg=fg,height=max(3,int(4*scale)),width=fw).place(x=0,y=0)
        self._win.update_idletasks()
        sw=self._win.winfo_screenwidth(); sh=self._win.winfo_screenheight()
        ww=self._win.winfo_width();       wh=self._win.winfo_height()
        self._win.geometry(f"+{sw-ww-20}+{sh-wh-60}")
        self._win.deiconify()
        # Re-assert topmost on every show so it stays above the game
        self._win.after(10, self._force_topmost)
        if self._job: self._root.after_cancel(self._job)
        self._job=self._root.after(2000, self._hide)

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

        x = self._cfg.get("mic_icon_x",-1)
        y = self._cfg.get("mic_icon_y",-1)
        if x<0 or y<0:
            sw=self._win.winfo_screenwidth(); x=sw-sz-24; y=120
        self._win.geometry(f"+{x}+{y}")

        self._canvas.bind("<ButtonPress-1>",   self._ds)
        self._canvas.bind("<B1-Motion>",       self._dm)
        self._canvas.bind("<ButtonRelease-1>", self._de)
        self._win.after(50, self._apply_clickthrough)
        self._win.after(60, self._force_topmost_mic)

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
            self._cfg["mic_icon_x"]=self._win.winfo_x()
            self._cfg["mic_icon_y"]=self._win.winfo_y()
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

# ── Hotkey engine ─────────────────────────────────────────────────────────────
class HotkeyEngine:
    """Routes all hotkeys through the single Win32 WH_KEYBOARD_LL hook.
    Works while any other key is held — same mechanism as Discord/OBS."""

    def __init__(self):
        pass  # _HOOK is global, shared

    def reload(self, cfg, on_vol, on_switch=None):
        _HOOK.clear()   # wipe previous bindings
        mode = cfg.get("mode","multi")

        if mode == "multi":
            for g in cfg["groups"]:
                keys = g.get("keys",{})

                def mk_vol(grp, delta):
                    def _():
                        if not grp.get("enabled",True): return
                        grp["volume"] = _calc_vol(grp["volume"], delta, cfg)
                        grp["muted"]  = False
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
                    ag["volume"]=_calc_vol(ag["volume"],ag.get("step",5),ref)
                    ag["muted"]=False; apply_vol(ag,ref); on_vol(ag)
            def _dn():
                ag = ref.get("_active_group_ref")
                if ag and ag.get("enabled",True):
                    ag["volume"]=_calc_vol(ag["volume"],-ag.get("step",5),ref)
                    ag["muted"]=False; apply_vol(ag,ref); on_vol(ag)
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
                g["volume"]=_calc_vol(g["volume"], g.get("step",5), cfg)
                g["muted"]=False; apply_vol(g,cfg); on_vol(g)

            def _hw_down():
                g = cfg.get("_active_group_ref")
                if not g or not g.get("enabled",True): return
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

        # Cycle key (both modes)
        ck = cfg.get("cycle_key","").strip()
        if ck and on_switch:
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
def make_tray_img(groups, enabled=True, mic_muted=None):
    """Ghost tray icon.
    enabled=False → grey ghost (app disabled)
    mic_muted=True → red ghost (mic muted)
    mic_muted=False → green ghost (mic live)
    mic_muted=None → default green (mic not in use)
    """
    sz  = 64
    img = Image.new("RGBA",(sz,sz),(0,0,0,0))
    d   = ImageDraw.Draw(img)
    if not enabled:
        col = "#444444"
    elif mic_muted is True:
        col = "#e74c3c"   # red = mic muted
    elif mic_muted is False:
        col = "#1DB954"   # green = mic live
    else:
        col = "#1DB954"   # default green
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
    return img

# ══════════════════════════════════════════════════════════════════════════════
# Settings Window
# ══════════════════════════════════════════════════════════════════════════════
class SettingsWin(tk.Toplevel):
    def __init__(self, parent, cfg, on_change):
        super().__init__(parent)
        self.cfg = cfg
        self.on_change = on_change
        self.title("Settings — KnobMixer")
        self.configure(bg=BG)
        self.geometry("540x580")
        self.resizable(True, True)
        self._build()

    # ── Simple scrollable tab helper ─────────────────────────────────────────
    def _make_tab(self, nb, title):
        """Add a notebook tab with auto-hiding minimal scrollbar."""
        outer = tk.Frame(nb, bg=BG)
        nb.add(outer, text=title)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0, bd=0)

        # Minimal scrollbar: thin strip, auto-hides
        sb_frame = tk.Frame(outer, bg=BG, width=10)
        sb_thumb  = tk.Frame(sb_frame, bg="#4a4a70", cursor="hand2")
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
        st = ttk.Style()
        st.theme_use("clam")
        st.configure("TNotebook", background=BG, borderwidth=0)
        st.configure("TNotebook.Tab", background=PANEL, foreground=TEXT,
                     padding=[14, 6])
        st.map("TNotebook.Tab", background=[("selected", BORDER)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        self._build_general(nb)
        self._build_slowdown(nb)
        self._build_single(nb)
        self._build_mic(nb)

        bar = tk.Frame(self, bg=PANEL, pady=8)
        bar.pack(fill="x", pady=(4, 0))
        tk.Label(bar, text="All changes apply instantly and are auto-saved",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=PANEL).pack(side="left", padx=14)
        tk.Button(bar, text="Close", font=("Segoe UI", 9),
                  bg=BORDER, fg=TEXT, relief="flat", cursor="hand2",
                  padx=14, pady=5, command=self.destroy).pack(side="right", padx=12)

    # ── General tab ──────────────────────────────────────────────────────────
    def _build_general(self, nb):
        sc = self._make_tab(nb, "General")

        self._v_startmin = tk.BooleanVar(value=self.cfg.get("start_minimized", True))
        self._v_startup  = tk.BooleanVar(value=get_startup())
        self._v_overlay  = tk.BooleanVar(value=self.cfg.get("show_overlay", True))
        self._v_ovsize   = tk.DoubleVar(value=self.cfg.get("overlay_size", 1.0))

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
        tk.Button(f_links, text="Report a bug",
                  font=("Segoe UI", 8), bg=BORDER, fg=TEXT,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=lambda: webbrowser.open(
                      f"https://github.com/{GITHUB_REPO}/issues" if GITHUB_REPO
                      else "mailto:your@email.com")).pack(side="left")

    # ── Slowdown tab ─────────────────────────────────────────────────────────
    def _build_slowdown(self, nb):
        sc = self._make_tab(nb, "Slowdown Zone")

        self._v_sden  = tk.BooleanVar(value=self.cfg.get("slowdown_enabled", True))
        self._v_sdthr = tk.DoubleVar(value=self.cfg.get("slowdown_threshold", 10))
        self._v_sdstp = tk.DoubleVar(value=self.cfg.get("slowdown_step", 0.5))

        tk.Label(sc, text="When volume drops below the threshold, steps switch to a finer amount for precise quiet control.",
                 font=("Segoe UI", 9), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(12, 4), anchor="w")

        self._sep(sc)
        self._row(sc, "Enable slowdown zone",
                  lambda p: self._chk(p, self._v_sden).pack(side="left"))
        self._row(sc, "Trigger below (%) threshold",
                  lambda p: self._sld(p, self._v_sdthr, 1, 40, 1).pack(side="left"))
        self._row(sc, "Fine step size (%)",
                  lambda p: self._sld(p, self._v_sdstp, 0.1, 5.0, 0.1).pack(side="left"))

    # ── 1-Knob tab ───────────────────────────────────────────────────────────
    def _build_single(self, nb):
        sc = self._make_tab(nb, "1-Knob Mode")

        self._v_sto         = tk.IntVar(value=self.cfg.get("single_timeout", 60))
        self._v_auto_revert = tk.BooleanVar(value=self.cfg.get("single_auto_revert", False))

        tk.Label(sc,
                 text="In 1-Knob mode, one set of keys controls all groups.\n"
                      "Press a group switch key to focus that group. By default\n"
                      "it stays on that group until you switch again.",
                 font=("Segoe UI", 9), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(12, 4), anchor="w")

        self._sep(sc, "Auto-Revert to Default Group")
        self._row(sc, "Enable auto-revert",
                  lambda p: self._chk(p, self._v_auto_revert).pack(side="left"))
        tk.Label(sc,
                 text="When ON: after the timeout, knob goes back to the starred (★) group.\n"
                      "When OFF: knob stays on whichever group you last switched to.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(0,6), anchor="w")
        self._row(sc, "Revert after (seconds)",
                  lambda p: tk.Spinbox(p, from_=5, to=600,
                                       textvariable=self._v_sto,
                                       width=5, font=("Segoe UI", 9),
                                       bg=PANEL, fg=TEXT,
                                       buttonbackground=BORDER,
                                       highlightthickness=0, relief="flat",
                                       command=self._apply).pack(side="left"))

        self._sep(sc, "Hardware Knob (e.g. AULAF75)")
        tk.Label(sc,
                 text="For keyboards with a physical volume knob that can't be remapped.\n"
                      "Intercepts the system volume keys and redirects them to control\n"
                      "your active group instead — system volume stays untouched.\n\n"
                      "The knob always controls whichever group is currently active.\n"
                      "Use the Cycle Key below to switch between groups.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(4,8), anchor="w")

        self._v_hw_en  = tk.BooleanVar(value=self.cfg.get("hw_knob_enabled", False))

        self._row(sc, "Enable hardware knob intercept",
                  lambda p: self._chk(p, self._v_hw_en).pack(side="left"))

        self._sep(sc, "Reset All Keybinds")
        tk.Label(sc,
                 text="Clears all hotkeys in every group, mic toggle, and shared keys. Useful for a fresh start.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(4,6), anchor="w")
        f_reset = tk.Frame(sc, bg=BG); f_reset.pack(fill="x", padx=16, pady=(0,8))
        def _reset_keybinds():
            import tkinter.messagebox as _mb
            if not _mb.askyesno("Reset Keybinds",
                    "Clear ALL hotkeys?\nThis cannot be undone.",
                    parent=self): return
            for g in self.cfg.get("groups",[]):
                g["keys"] = {"vol_down":"","vol_up":"","mute":""}
                g["single_key"] = ""
            self.cfg["single_keys"] = {"vol_down":"","vol_up":"","mute":""}
            self.cfg["cycle_key"]   = ""
            self.cfg["mic_hotkey"]  = ""
            self._apply()
            _mb.showinfo("KnobMixer","All keybinds cleared.", parent=self)
        tk.Button(f_reset, text="Reset All Keybinds",
                  font=("Segoe UI", 9), bg="#3a1a1a", fg="#ff6b6b",
                  activebackground="#4a2020", activeforeground="#ff9999",
                  relief="flat", cursor="hand2", padx=12, pady=5,
                  command=_reset_keybinds).pack(side="left")

        self._sep(sc, "Group Cycle Key")
        tk.Label(sc, text="Press this key to cycle through all enabled groups. Works in both modes. Popup shows active group name.",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG,
                 justify="left").pack(padx=16, pady=(4, 6), anchor="w")

        cur_ck = self.cfg.get("cycle_key", "")
        def _ck_cb(hk):
            self.cfg["cycle_key"] = hk
            self._apply()
        f_ck = tk.Frame(sc, bg=BG)
        f_ck.pack(fill="x", padx=16, pady=4)
        tk.Label(f_ck, text="Cycle key", font=("Segoe UI", 9), fg=TEXT,
                 bg=BG, width=28, anchor="w").pack(side="left")
        make_hotkey_btn(f_ck, cur_ck, _ck_cb).pack(side="left")

        self._sep(sc, "Shared Volume Keys (1-Knob Mode)")
        tk.Label(sc, text="Click a button then press your key (combos supported).",
                 font=("Segoe UI", 8), fg=SUBTEXT, bg=BG).pack(padx=16, anchor="w")

        sk = self.cfg.setdefault("single_keys", {"vol_down": "", "vol_up": "", "mute": ""})
        for action, lbl in [("vol_down", "Vol-  (quieter)"),
                             ("vol_up",   "Vol+  (louder)"),
                             ("mute",     "Volume Mute toggle")]:
            cur = sk.get(action, "")
            def _cb(hk, a=action):
                self.cfg["single_keys"][a] = hk
                self._apply()
            f = tk.Frame(sc, bg=BG)
            f.pack(fill="x", padx=16, pady=4)
            tk.Label(f, text=lbl, font=("Segoe UI", 9), fg=TEXT,
                     bg=BG, width=28, anchor="w").pack(side="left")
            make_hotkey_btn(f, cur, _cb).pack(side="left")

    # ── Mic tab ──────────────────────────────────────────────────────────────
    def _build_mic(self, nb):
        sc = self._make_tab(nb, "Mic Toggle")

        self._v_micen    = tk.BooleanVar(value=self.cfg.get("mic_enabled", True))
        self._v_michk    = tk.StringVar(value=self.cfg.get("mic_hotkey", "f9"))
        self._v_micst    = tk.BooleanVar(value=self.cfg.get("mic_start_muted", False))
        self._v_micvol   = tk.DoubleVar(value=self.cfg.get("mic_sound_volume", 0.15))
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
                              font=("Segoe UI", 9))
        dev_cb.pack(side="left")
        dev_cb.bind("<<ComboboxSelected>>", lambda e: self._apply())

        # Mic hotkey
        cur_hk = self.cfg.get("mic_hotkey", "f9")
        def _hk_cb(hk):
            self._v_michk.set(hk)
            self._apply()
        f_hk = tk.Frame(sc, bg=BG)
        f_hk.pack(fill="x", padx=16, pady=7)
        tk.Label(f_hk, text="Mic mute hotkey", font=("Segoe UI", 9),
                 fg=TEXT, bg=BG, width=28, anchor="w").pack(side="left")
        make_hotkey_btn(f_hk, cur_hk, _hk_cb).pack(side="left")

        self._sep(sc, "Icon")
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
                          font=("Segoe UI", 9))
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
                  lambda p: self._sld(p, self._v_micvol, 0.001, 0.3, 0.005).pack(side="left"))

        tk.Label(sc, text="Click a sound to preview it — stays selected:",
                 font=("Segoe UI", 9), fg=TEXT, bg=BG).pack(padx=16, pady=(8, 4), anchor="w")

        # Grid of sound preset buttons
        sf = tk.Frame(sc, bg=PANEL, padx=6, pady=6)
        sf.pack(fill="x", padx=16, pady=(0, 10))
        self._preset_btns = []
        cur_p = self.cfg.get("mic_sound_preset", 0)
        self._preview_muted = True  # alternates on each click
        for i, (name, shape, mute_p, unmute_p) in enumerate(SOUND_PRESETS):
            row_i = i // 4
            col_i = i % 4
            is_sel = (i == cur_p)
            def _pick(idx=i, sh=shape, mp=mute_p, up=unmute_p):
                self._v_micpre.set(idx)
                self._preview_muted = not self._preview_muted
                fr, dur2, vs = mp if self._preview_muted else up
                play_sound(fr, dur2, self._v_micvol.get() * vs, sh)
                self._apply()
                self._refresh_preset_btns()
            btn = tk.Button(sf, text=name, font=("Segoe UI", 8),
                            bg="#1a3a1a" if is_sel else BORDER,
                            fg="#1DB954" if is_sel else TEXT,
                            relief="flat", cursor="hand2", padx=6, pady=3,
                            command=_pick)
            btn.grid(row=row_i, column=col_i, padx=3, pady=3, sticky="ew")
            self._preset_btns.append(btn)
        for col in range(4):
            sf.columnconfigure(col, weight=1)

    def _refresh_preset_btns(self):
        cur = self._v_micpre.get()
        for i, btn in enumerate(self._preset_btns):
            btn.config(bg="#1a3a1a" if i == cur else BORDER,
                       fg="#1DB954" if i == cur else TEXT)

    # ── Apply ────────────────────────────────────────────────────────────────
    def _get_mic_devices(self):
        """Get list of available microphone names."""
        return get_mic_devices()

    def _apply(self):
        c = self.cfg
        c["start_minimized"]   = self._v_startmin.get()
        c["show_overlay"]      = self._v_overlay.get()
        c["analytics_enabled"] = self._v_analytics.get()
        c["analytics_enabled"] = self._v_analytics.get()
        c["overlay_size"]      = round(self._v_ovsize.get(), 1)
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
        c["mic_sound_volume"]  = round(self._v_micvol.get(), 3)
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

# ══════════════════════════════════════════════════════════════════════════════
# Main UI
# ══════════════════════════════════════════════════════════════════════════════
class App:
    def __init__(self):
        self.cfg=load_cfg()
        self.hk=HotkeyEngine()
        self.mic=MicCtrl()
        self._enabled=True
        self._active_grp=None
        self._timeout_job=None
        self._timeout_start=0
        self._group_widgets=[]
        
        self._build_win()
        self.overlay=VolumeOverlay(self.root)
        self.mic_ov=None
        self._build_ui()
        self._init_single()
        _HOOK.start()
        self.hk.reload(self.cfg,self._on_vol,self._on_switch)
        self._reg_mic_hk()

        if self.cfg.get("mic_enabled",True):
            self.mic.sync()
            if self.cfg.get("mic_start_muted",False):
                self.mic.set(True,{"mic_sound_volume":0,"mic_sound_preset":0})
            self.mic_ov=MicOverlay(self.root,self.mic,self.cfg)

        self._refresh_loop()
        self._setup_tray()
        if self.cfg.get("start_minimized",True):
            self.root.after(200,self._hide)

    # ── Window ────────────────────────────────────────────────────────────────
    def _build_win(self):
        self.root=tk.Tk()
        self.root.title(APP_NAME)
        self.root.configure(bg=BG)
        self.root.resizable(True,True)
        self.root.minsize(520,400)
        self.root.protocol("WM_DELETE_WINDOW",self._hide)
        w,h=560,720
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
        hdr=tk.Frame(self.root,bg="#111118",pady=10); hdr.pack(fill="x")
        lf=tk.Frame(hdr,bg="#111118"); lf.pack(side="left",padx=14)
        tk.Label(lf,text="KnobMixer",font=("Segoe UI",15,"bold"),fg=TEXT,bg="#111118").pack(anchor="w")
        tk.Label(lf,text=f"v{APP_VER}  —  Per-app volume control for knobs & hotkeys",
                 font=("Segoe UI",8),fg=SUBTEXT,bg="#111118").pack(anchor="w")

        rf=tk.Frame(hdr,bg="#111118"); rf.pack(side="right",padx=14)
        tk.Button(rf,text="⚙",font=("Segoe UI",13),bg="#111118",fg=SUBTEXT,
                  activebackground="#111118",activeforeground=TEXT,
                  relief="flat",cursor="hand2",
                  command=self._open_settings).pack(side="right",padx=(4,0))
        self._onoff_btn=tk.Button(rf,text="ON",font=("Segoe UI",9,"bold"),
                                  bg="#1a3a1a",fg="#1DB954",activebackground="#2a4a2a",
                                  relief="flat",cursor="hand2",padx=10,pady=4,
                                  command=self._toggle_en)
        self._onoff_btn.pack(side="right",padx=(0,6))

        self._update_url = [None]  # store URL for click
        # Update check button — styled like ON button, sits next to gear icon
        self._update_btn = tk.Button(
            rf, text="Update",
            font=("Segoe UI",8), bg=BORDER, fg=SUBTEXT,
            activebackground=HOVER, activeforeground=TEXT,
            relief="flat", cursor="hand2", padx=8, pady=4,
            command=self._manual_update_check)
        self._update_btn.pack(side="right", padx=(0,4))

        # Mode bar
        mb=tk.Frame(self.root,bg=PANEL,pady=6); mb.pack(fill="x")
        tk.Label(mb,text="Mode:",font=("Segoe UI",9),fg=SUBTEXT,bg=PANEL).pack(side="left",padx=(14,4))
        self._mode_var=tk.StringVar(value=self.cfg.get("mode","multi"))
        for val,lbl in [("multi","Multiple Knobs"),("single","1 Knob")]:
            tk.Radiobutton(mb,text=lbl,value=val,variable=self._mode_var,
                           font=("Segoe UI",9),fg=TEXT,bg=PANEL,selectcolor=BORDER,
                           activebackground=PANEL,activeforeground=TEXT,
                           command=self._on_mode).pack(side="left",padx=6)

        # Single-knob status bar
        self._sb=tk.Frame(self.root,bg="#111118",pady=3)
        self._active_lbl=tk.Label(self._sb,text="",font=("Segoe UI",9,"bold"),
                                  fg="#1DB954",bg="#111118")
        self._active_lbl.pack(side="left",padx=14)
        self._timeout_lbl=tk.Label(self._sb,text="",font=("Segoe UI",8),
                                   fg=SUBTEXT,bg="#111118")
        self._timeout_lbl.pack(side="right",padx=14)
        if self.cfg.get("mode")=="single": self._sb.pack(fill="x")

        # Groups — scrollable canvas with minimal auto-hide scrollbar
        cf=tk.Frame(self.root,bg=BG); cf.pack(fill="both",expand=True)
        self._canvas=tk.Canvas(cf,bg=BG,highlightthickness=0,bd=0)
        # Minimal custom scrollbar: thin dark strip
        self._sb_frame=tk.Frame(cf,bg=BG,width=10)
        self._sb_thumb=tk.Frame(self._sb_frame,bg="#4a4a70",cursor="hand2")
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
        if hasattr(self, "_master_vol_btn"):
            self._update_master_vol_btn()

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

        if dy < -28 and cur > 0:
            groups[cur], groups[cur-1] = groups[cur-1], groups[cur]
            self._drag_state["idx"]     = cur - 1
            self._drag_state["start_y"] = event.y_root
            moved = True
        elif dy > 28 and cur < len(groups)-1:
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
        # Remove root-level bindings — back to normal
        self.root.unbind("<B1-Motion>")
        self.root.unbind("<ButtonRelease-1>")

    # Keep old signatures for the handle bindings (they just start the drag)
    def _card_drag_motion(self, event, idx): pass
    def _card_drag_end(self, event): pass

    def _card(self,idx,group):
        color=group.get("color","#888")
        mode =self.cfg.get("mode","multi")

        card=tk.Frame(self.gf,bg=PANEL,highlightbackground=color,highlightthickness=2)
        card.pack(fill="x",padx=12,pady=3)

        # Row 1: handle | dot | name | ★ | ON/OFF | del
        r1=tk.Frame(card,bg=PANEL); r1.pack(fill="x",padx=10,pady=(7,1))

        # Drag handle — ≡ symbol, cursor changes to indicate draggable
        handle=tk.Label(r1,text="≡",font=("Segoe UI",12),fg="#444",bg=PANEL,
                        cursor="sb_v_double_arrow")
        handle.pack(side="left",padx=(0,4))
        handle.bind("<ButtonPress-1>",  lambda e,i=idx: self._card_drag_start(e,i))
        handle.bind("<B1-Motion>",      lambda e,i=idx: self._card_drag_motion(e,i))
        handle.bind("<ButtonRelease-1>",self._card_drag_end)

        dot=tk.Label(r1,text="●",font=("Segoe UI",12),fg=color,bg=PANEL,cursor="hand2")
        dot.pack(side="left")
        dot.bind("<Button-1>",lambda e,g=group,c=card: self._pick_color(g,c))
        nv=tk.StringVar(value=group.get("name","Group"))
        ne=tk.Entry(r1,textvariable=nv,font=("Segoe UI",10,"bold"),
                    bg=PANEL,fg=TEXT,insertbackground=TEXT,relief="flat",bd=0,width=13)
        ne.pack(side="left",padx=6)
        ne.bind("<KeyRelease>",lambda e,g=group,v=nv: self._on_name(g,v))
        if group.get("master_volume"):
            tk.Label(r1,text="🔊 PC",font=("Segoe UI",7),fg="#00BCD4",
                     bg=PANEL).pack(side="left")

        tk.Button(r1,text="✕",font=("Segoe UI",8),fg=SUBTEXT,bg=PANEL,
                  activebackground=BORDER,activeforeground="#ff6b6b",
                  relief="flat",cursor="hand2",
                  command=lambda i=idx: self._del_group(i)).pack(side="right",padx=(2,0))

        en_btn=tk.Button(r1,text="ON" if group.get("enabled",True) else "OFF",
                         font=("Segoe UI",8,"bold"),bg=BORDER,
                         fg="#1DB954" if group.get("enabled",True) else "#666",
                         activebackground=HOVER,relief="flat",cursor="hand2",padx=8,pady=1)
        en_btn.pack(side="right",padx=2)
        def _tog(g=group,b=en_btn):
            g["enabled"]=not g.get("enabled",True)
            b.config(text="ON" if g["enabled"] else "OFF",
                     fg="#1DB954" if g["enabled"] else "#666")
            self._autosave()
        en_btn.config(command=_tog)

        if mode=="single":
            is_def=group.get("is_default",False)
            def_btn=tk.Button(r1,text="★" if is_def else "☆",
                              font=("Segoe UI",11),
                              fg="#FFC107" if is_def else SUBTEXT,
                              bg=PANEL,activebackground=PANEL,
                              relief="flat",cursor="hand2")
            def_btn.pack(side="right",padx=2)
            def _set_def(g=group):
                for gg in self.cfg["groups"]: gg["is_default"]=False
                g["is_default"]=True
                self.cfg["single_default_group"]=self.cfg["groups"].index(g)
                self._autosave(); self._redraw()
            def_btn.config(command=_set_def)

        # Row 2: vol% | step | mute
        r2=tk.Frame(card,bg=PANEL); r2.pack(fill="x",padx=10,pady=(1,0))
        vol_lbl=tk.Label(r2,text=self._vt(group),font=("Segoe UI",9,"bold"),
                         fg=color,bg=PANEL,width=7,anchor="w")
        vol_lbl.pack(side="left")

        step_var=tk.IntVar(value=group.get("step",5))
        def _step_chg(g=group,v=step_var): g["step"]=v.get(); self._autosave()
        tk.Label(r2,text="step:",font=("Segoe UI",8),fg=SUBTEXT,bg=PANEL).pack(side="right",padx=(0,2))
        tk.Spinbox(r2,from_=1,to=25,textvariable=step_var,width=3,
                   font=("Segoe UI",8),bg=BG,fg=TEXT,buttonbackground=BORDER,
                   highlightthickness=0,relief="flat",
                   command=_step_chg).pack(side="right",padx=(0,4))

        mute_btn=tk.Button(r2,text="Muted" if group.get("muted") else "Mute",
                           font=("Segoe UI",8),
                           bg="#3a1a1a" if group.get("muted") else BORDER,
                           fg="#ff6b6b" if group.get("muted") else SUBTEXT,
                           relief="flat",cursor="hand2",padx=6,pady=1)
        mute_btn.pack(side="right",padx=(0,6))
        mute_btn.config(command=self._mk_mute(group,mute_btn,vol_lbl,color))

        # Slider
        vol_var=tk.IntVar(value=int(group.get("volume",80)))
        stn=f"G{idx}.Horizontal.TScale"
        st=ttk.Style(); st.theme_use("clam")
        st.configure(stn,troughcolor=BORDER,background=color,sliderlength=15,sliderrelief="flat")
        ttk.Scale(card,from_=0,to=100,orient="horizontal",variable=vol_var,
                  style=stn,
                  command=self._mk_slider(group,vol_var,vol_lbl,color)
                  ).pack(fill="x",padx=10,pady=(2,4))

        # Row 3: hotkeys with labels
        r3=tk.Frame(card,bg=PANEL); r3.pack(fill="x",padx=10,pady=(0,3))
        if mode=="multi":
            for action,label in [("vol_down","Vol-"),("vol_up","Vol+"),("mute","Vol Mute")]:
                cur=group.get("keys",{}).get(action,"")
                col2=SUBTEXT
                lf2=tk.Frame(r3,bg=PANEL); lf2.pack(side="left",padx=(0,8))
                tk.Label(lf2,text=label,font=("Segoe UI",7),fg=SUBTEXT,bg=PANEL).pack(anchor="w")
                def _cb(hk,g=group,a=action):
                    g.setdefault("keys",{})[a]=hk
                    self.hk.reload(self.cfg,self._on_vol,self._on_switch)
                    self._autosave()
                btn=make_hotkey_btn(lf2,cur,_cb)
                btn.pack(anchor="w")
        else:
            sk=group.get("single_key","")
            lf2=tk.Frame(r3,bg=PANEL); lf2.pack(side="left")
            tk.Label(lf2,text="Switch key",font=("Segoe UI",7),fg=SUBTEXT,bg=PANEL).pack(anchor="w")
            def _cb_sk(hk,g=group):
                g["single_key"]=hk
                self.hk.reload(self.cfg,self._on_vol,self._on_switch)
                self._autosave()
            make_hotkey_btn(lf2,sk,_cb_sk).pack(anchor="w")

        # Row 4: apps (skip for master volume groups)
        if not group.get("master_volume"):
            r4=tk.Frame(card,bg=PANEL); r4.pack(fill="x",padx=10,pady=(0,7))
            apps_lbl=tk.Label(r4,text=self._at(group),font=("Segoe UI",8),
                              fg=SUBTEXT,bg=PANEL,anchor="w",wraplength=420)
            apps_lbl.pack(side="left",fill="x",expand=True)
            tk.Button(r4,text="Edit",font=("Segoe UI",8),bg=BORDER,fg=TEXT,
                      relief="flat",cursor="hand2",padx=6,pady=1,
                      command=lambda g=group,l=apps_lbl:
                              AppsDialog(self.root,g,l,self._at,self._autosave)
                      ).pack(side="right")
        else:
            tk.Label(card,text="Controls the entire PC volume",
                     font=("Segoe UI",8),fg="#00BCD4",bg=PANEL).pack(
                     anchor="w",padx=10,pady=(0,7))

        self._group_widgets.append({
            "vol_var":vol_var,"vol_lbl":vol_lbl,
            "mute_btn":mute_btn,"color":color,"step_var":step_var,
        })

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _vt(self,g): return "Muted" if g.get("muted") else f"{int(g.get('volume',80))}%"
    def _at(self,g):
        if g.get("master_volume"): return ""
        apps=g.get("apps",[]); return (", ".join(apps[:6])+(f" +{len(apps)-6} more" if len(apps)>6 else "")) if apps else "No apps — click Edit to add"

    def _mk_slider(self,group,vol_var,lbl,color):
        def _cmd(val):
            group["volume"]=float(val); group["muted"]=False
            lbl.config(text=self._vt(group),fg=color)
            apply_vol(group,self.cfg); self._autosave()
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
        self._autosave(); self._redraw()

    def _del_group(self,idx):
        if len(self.cfg["groups"])<=1:
            messagebox.showwarning("KnobMixer","Need at least one group.",parent=self.root); return
        name=self.cfg["groups"][idx].get("name","Group")
        if messagebox.askyesno("Delete",f'Delete "{name}"?',parent=self.root):
            self.cfg["groups"].pop(idx); self._autosave(); self._redraw()

    # ── Mode ──────────────────────────────────────────────────────────────────
    def _on_mode(self):
        self.cfg["mode"]=self._mode_var.get()
        self._init_single()
        self.hk.reload(self.cfg,self._on_vol,self._on_switch)
        self._redraw()
        if self.cfg["mode"]=="single": self._sb.pack(fill="x",after=self.root.children.get("!frame3",self._sb))
        else: self._sb.pack_forget()
        self._autosave()

    def _init_single(self):
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
        # Show overlay with group name when switching
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
        idx=self.cfg.get("single_default_group",0); gs=self.cfg["groups"]
        if gs:
            self._active_grp=gs[min(idx,len(gs)-1)]
            self.cfg["_active_group_ref"]=self._active_grp
            self.root.after(0,self._update_active_lbl)
        self._timeout_job=None; self.root.after(0,lambda: self._timeout_lbl.config(text=""))

    def _tick(self):
        if self._timeout_job is None: return
        remain=max(0,self.cfg.get("single_timeout",60)-(time.time()-self._timeout_start))
        self._timeout_lbl.config(text=f"Reverts in {int(remain)}s")
        if remain>0: self.root.after(1000,self._tick)

    def _update_active_lbl(self):
        if self._active_grp:
            self._active_lbl.config(text=f"Active: {self._active_grp['name']}",
                                    fg=self._active_grp.get("color","#1DB954"))

    # ── Global on/off ─────────────────────────────────────────────────────────
    def set_update_available(self, ver, url):
        """Called when a newer version is found. Updates button quietly."""
        self._update_url[0] = url
        self._update_btn.config(
            text=f"v{ver} available",
            bg="#1a3a1a", fg="#1DB954",
            command=self._open_update_url)

    def _open_update_url(self):
        import webbrowser
        if self._update_url[0]:
            webbrowser.open(self._update_url[0])

    def _manual_update_check(self):
        """User clicked Update button — disable it while checking."""
        self._update_btn.config(text="Checking…", fg=SUBTEXT, bg=BORDER,
                                command=lambda: None)  # disable while checking
        def _reset():
            if self._update_btn.winfo_exists():
                self._update_btn.config(text="Update", fg=SUBTEXT, bg=BORDER,
                                        command=self._manual_update_check)
        def _on_version(ver, url):
            if ver and _ver_tuple(ver) > _ver_tuple(APP_VER):
                self.root.after(0, lambda: self.set_update_available(ver, url))
            elif ver:
                # Up to date — show for 30s then reset
                self.root.after(0, lambda: self._update_btn.config(
                    text="Up to date ✓", fg=SUBTEXT, bg=BORDER))
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
            self._onoff_btn.config(text="ON",bg="#1a3a1a",fg="#1DB954")
            self.hk.reload(self.cfg,self._on_vol,self._on_switch)
            self._reg_mic_hk()
        else:
            self._onoff_btn.config(text="OFF",bg="#2a1a1a",fg="#555")
            self.hk.stop()
        self.overlay.set_enabled(self._enabled)
        try: self.tray.icon=make_tray_img(self.cfg["groups"],self._enabled,mic_muted=self.mic.get() if self.cfg.get("mic_enabled") else None)
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
        if not hk or not self.cfg.get("mic_enabled", True):
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
                self.cfg["groups"], self._enabled, mic_muted=mic_muted)
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
        self._sync()
        self.root.after(500,self._refresh_loop)

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
        self._dirty_lbl.config(text="✓ Saved")
        self.root.after(1500,lambda: self._dirty_lbl.config(text=""))

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
        self._settings_win = SettingsWin(self.root, self.cfg, self._on_settings_change)

    def _on_settings_change(self):
        self.hk.reload(self.cfg,self._on_vol,self._on_switch)
        self._reg_mic_hk()
        if self.cfg.get("mic_enabled",True):
            if self.mic_ov: self.mic_ov.update()
            else: self.mic_ov=MicOverlay(self.root,self.mic,self.cfg)
        else:
            if self.mic_ov: self.mic_ov.hide()

    # ── Tray ─────────────────────────────────────────────────────────────────
    def _setup_tray(self):
        img=make_tray_img(self.cfg["groups"],self._enabled,mic_muted=self.mic.get() if self.cfg.get("mic_enabled") else None)
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
        if self._settings_open():
            self._settings_win.destroy()
        self.root.withdraw()
    def _show(self,*_): self.root.after(0,self.root.deiconify); self.root.after(0,self.root.lift)
    def _quit(self,*_):
        if self._settings_open():
            self._settings_win.destroy()
        self.hk.stop(); _HOOK.stop(); self.tray.stop(); self.root.after(0,self.root.destroy)

    def run(self): self.root.mainloop()


# ══════════════════════════════════════════════════════════════════════════════
# Apps Dialog
# ══════════════════════════════════════════════════════════════════════════════
class AppsDialog(tk.Toplevel):
    def __init__(self,parent,group,apps_lbl,summary_fn,save_fn):
        super().__init__(parent)
        self.group=group; self.apps_lbl=apps_lbl
        self.summary_fn=summary_fn; self.save_fn=save_fn
        self.title(f"Apps — {group.get('name','Group')}")
        self.configure(bg=BG); self.geometry("460x520")
        self.resizable(True,True); self.grab_set(); self._build()

    def _build(self):
        # ── Top: manual text entry ────────────────────────────────────────────
        tk.Label(self,text="App process names, one per line (no .exe):",
                 font=("Segoe UI",9),fg=TEXT,bg=BG).pack(padx=12,pady=(12,4),anchor="w")
        self._txt=tk.Text(self,font=("Consolas",9),bg=PANEL,fg=TEXT,
                          insertbackground=TEXT,relief="flat",height=6,
                          padx=6,pady=6)
        self._txt.pack(fill="x",padx=12)
        self._txt.insert("1.0","\n".join(self.group.get("apps",[])))

        # ── Middle: all running audio apps in a scrollable area ───────────────
        hdr=tk.Frame(self,bg=BG); hdr.pack(fill="x",padx=12,pady=(8,2))
        tk.Label(hdr,text="Currently making sound — click to add:",
                 font=("Segoe UI",8),fg=SUBTEXT,bg=BG).pack(side="left")
        tk.Button(hdr,text="↻ Refresh",font=("Segoe UI",7),bg=BORDER,fg=SUBTEXT,
                  relief="flat",cursor="hand2",padx=4,pady=1,
                  command=self._refresh_r).pack(side="right")

        # Scrollable canvas for running apps — shows ALL of them
        rf_outer=tk.Frame(self,bg=PANEL,highlightthickness=1,
                          highlightbackground=BORDER)
        rf_outer.pack(fill="both",expand=True,padx=12,pady=(0,4))

        self._canvas=tk.Canvas(rf_outer,bg=PANEL,highlightthickness=0,
                               height=160)
        sb=tk.Scrollbar(rf_outer,orient="vertical",command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side="left",fill="both",expand=True)
        sb.pack(side="right",fill="y")

        self._rf=tk.Frame(self._canvas,bg=PANEL)
        self._cwin=self._canvas.create_window((0,0),window=self._rf,anchor="nw")
        self._rf.bind("<Configure>",lambda e:(
            self._canvas.configure(scrollregion=self._canvas.bbox("all")),
            self._canvas.itemconfig(self._cwin,width=self._canvas.winfo_width())))
        self._canvas.bind("<Configure>",lambda e:
            self._canvas.itemconfig(self._cwin,width=e.width))
        self._canvas.bind("<MouseWheel>",lambda e:
            self._canvas.yview_scroll(int(-1*e.delta/120),"units"))

        self._refresh_r()

        # ── Bottom: save/cancel ───────────────────────────────────────────────
        bf=tk.Frame(self,bg=BG); bf.pack(fill="x",padx=12,pady=(4,10))
        tk.Button(bf,text="Save",font=("Segoe UI",10,"bold"),bg="#1DB954",fg="white",
                  relief="flat",cursor="hand2",padx=18,pady=4,
                  command=self._save).pack(side="left")
        tk.Button(bf,text="Cancel",font=("Segoe UI",10),bg=BORDER,fg=TEXT,
                  relief="flat",cursor="hand2",padx=12,pady=4,
                  command=self.destroy).pack(side="left",padx=6)

    def _refresh_r(self):
        for w in self._rf.winfo_children(): w.destroy()
        apps=running_audio_apps()
        already=[l.strip() for l in self._txt.get("1.0","end").splitlines() if l.strip()]
        if not apps:
            tk.Label(self._rf,text="  (no apps making sound right now)",
                     font=("Segoe UI",8),fg=SUBTEXT,bg=PANEL,
                     pady=8).pack(anchor="w")
            return
        # Wrap buttons into rows of 3
        cols=3; row_f=None
        for i,a in enumerate(sorted(apps)):
            if i % cols == 0:
                row_f=tk.Frame(self._rf,bg=PANEL)
                row_f.pack(fill="x",padx=4,pady=1)
            already_added = a in already
            btn=tk.Button(row_f,
                text=f"✓ {a}" if already_added else f"+ {a}",
                font=("Consolas",8),
                bg="#1a3a1a" if already_added else BORDER,
                fg="#1DB954" if already_added else TEXT,
                activebackground=HOVER,relief="flat",cursor="hand2",
                padx=6,pady=2,width=14)
            btn.pack(side="left",padx=(0,4))
            if not already_added:
                btn.config(command=lambda n=a,b=btn: self._add(n,b))

    def _add(self,name,btn=None):
        cur=self._txt.get("1.0","end").strip()
        if name not in [l.strip() for l in cur.splitlines() if l.strip()]:
            self._txt.insert("end",f"\n{name}" if cur else name)
        if btn:
            btn.config(text=f"✓ {name}",bg="#1a3a1a",fg="#1DB954",
                       command=lambda:None)

    def _save(self):
        lines=[l.strip().lower() for l in
               self._txt.get("1.0","end").splitlines() if l.strip()]
        self.group["apps"]=lines
        self.apps_lbl.config(text=self.summary_fn(self.group))
        self.save_fn(); self.destroy()


# ── Analytics & Update checker ───────────────────────────────────────────────
# ANALYTICS_URL: set to your Cloudflare Worker / server endpoint.
# When set, sends one small anonymous ping on startup per day.
# Payload: random install ID (no personal data) + version + Windows version.
# User can disable in Settings → General → "Send anonymous usage data".
#
# GITHUB_REPO: your GitHub username/repo — used for update checks.
# Set both before building your public release.
ANALYTICS_URL = "https://knobmixer-analytics.bdhhair11.workers.dev/ping"
GITHUB_REPO   = "KnobMixer/KnobMixer"
UPDATE_CHECK  = True        # set False to disable update checks entirely

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
    """Return True if we haven't pinged analytics today yet."""
    stamp_file = APPDATA_DIR / "last_ping"
    import datetime
    today = datetime.date.today().isoformat()
    if stamp_file.exists() and stamp_file.read_text().strip() == today:
        return False
    stamp_file.write_text(today)
    return True

def _ping_analytics(cfg=None):
    """Daily analytics ping with exponential backoff retry.
    Handles offline-at-launch then reconnect — like Steam/Discord do it."""
    if cfg and not cfg.get("analytics_enabled", True): return
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
            # Success — stamp today so we don't retry again
        except Exception:
            # Retry with exponential backoff: 30s, 5min, 30min, give up
            delays = [30, 300, 1800]
            if attempt < len(delays):
                time.sleep(delays[attempt])
                _send(attempt + 1)
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
                callback(APP_VER, "")   # same version = up to date
                return
            latest = data.get("tag_name","").lstrip("v")
            dl_url = data.get("html_url",
                              f"https://github.com/{GITHUB_REPO}/releases/latest")
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

def _should_check_update_this_week():
    """Returns True once per week — so the auto check isn't every single launch."""
    import datetime
    stamp_file = APPDATA_DIR / "last_update_check"
    today      = datetime.date.today()
    if stamp_file.exists():
        try:
            last = datetime.date.fromisoformat(stamp_file.read_text().strip())
            if (today - last).days < 7:
                return False
        except: pass
    stamp_file.write_text(today.isoformat())
    return True

if __name__=="__main__":
    # ── Single instance check ─────────────────────────────────────────────────
    import ctypes
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "KnobMixer_SingleInstance")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        # Already running — bring existing window to front
        import ctypes.wintypes
        def _enum(hwnd, _):
            buf = ctypes.create_unicode_buffer(256)
            ctypes.windll.user32.GetWindowTextW(hwnd, buf, 256)
            if "KnobMixer" in buf.value:
                ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                return False
            return True
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(_enum), 0)
        raise SystemExit(0)

    _ping_analytics()
    app = App()
    # Weekly silent update check — no popups, just updates the button label
    if _should_check_update_this_week():
        def _on_version(ver, url):
            if ver and _ver_tuple(ver) > _ver_tuple(APP_VER):
                app.root.after(0, lambda: app.set_update_available(ver, url))
        _fetch_latest_version(_on_version, auto=True)
    app.run()
