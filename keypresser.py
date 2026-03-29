# ==============================================================================
# KeyPresser - Desktop Automation Tool for Silkroad Private Server
# ==============================================================================
#
# DEPENDENCIES - Install with:
#   pip install pywin32 keyboard psutil
#
# Or use the included requirements.txt:
#   pip install -r requirements.txt
#
# USAGE:
#   python keypresser.py
# ==============================================================================

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import ctypes
import time
import threading
import datetime
import tempfile
import subprocess
import urllib.request
import win32gui
import win32process
import win32con
import win32api
import psutil
import keyboard

try:
    import cv2
    from PIL import ImageGrab, ImageEnhance, ImageTk
    import numpy as np
    _IMAGE_MATCHING = True
except ImportError:
    _IMAGE_MATCHING = False

try:
    import pytesseract
    _OCR_AVAILABLE = True
except ImportError:
    _OCR_AVAILABLE = False

# ==============================================================================
# Version & update config
# ==============================================================================
APP_VERSION  = 4                          # bump this with every release
GITHUB_REPO  = "LinKxFr/InvictusSROKP"   # used for update checks


# ==============================================================================
# Admin elevation check — must run before any GUI is created
# ==============================================================================
def _is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def _relaunch_as_admin():
    """Re-launch this script elevated via the UAC 'runas' verb."""
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas",
        sys.executable,
        " ".join(f'"{a}"' for a in sys.argv),
        None, 1
    )

if not _is_admin():
    try:
        _relaunch_as_admin()
        sys.exit(0)
    except Exception:
        _root = tk.Tk()
        _root.withdraw()
        messagebox.showerror(
            "Administrator required",
            "KeyPresser was NOT launched as Administrator.\n\n"
            "It will not be able to send keystrokes to elevated windows "
            "(e.g. the Silkroad game client).\n\n"
            "Please right-click the script and choose 'Run as administrator', "
            "or re-run from an elevated command prompt.",
        )
        _root.destroy()
        sys.exit(1)

# --------------------------------------------------------------------------
# Configuration file path — saved next to this script
# --------------------------------------------------------------------------
CONFIG_FILE           = os.path.join(os.path.dirname(os.path.abspath(__file__)), "keypresser_config.json")
ALCHEMY_TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "alchemy_template.png")

# --------------------------------------------------------------------------
# Default configuration
# --------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "sequence": [
        {"key": "1", "vk": 0x31, "delay_ms": 50},
        {"key": "2", "vk": 0x32, "delay_ms": 50},
        {"key": "3", "vk": 0x33, "delay_ms": 50},
        {"key": "4", "vk": 0x34, "delay_ms": 50},
        {"key": "5", "vk": 0x35, "delay_ms": 50},
    ],
    # Timed actions: press-and-hold a key every N minutes.
    # Each entry:
    #   key         — display label
    #   vk          — Windows VK code (layout-independent)
    #   hold_ms     — how long to hold the key down (ms)
    #   interval_min — repeat every N minutes
    #   initial_min  — first trigger delay in minutes (e.g. "48 min left on buff")
    "timed_actions": [],
    "hotkey_start": "F6",
    "hotkey_stop":  "F7",
    "target_window": "",
    "alchemy_delay_ms":   3500,
    "alchemy_hotkey":     "F8",
    "alchemy_text_region": None,   # [x, y, w, h] in screen coords
    "alchemy_stop_text":  "changed to [5].",
}

# --------------------------------------------------------------------------
# Colour palette — dark theme
# --------------------------------------------------------------------------
BG       = "#1e1e2e"
BG2      = "#2a2a3e"
BG3      = "#313145"
ACCENT   = "#89b4fa"
ACCENT2  = "#cba6f7"
GREEN    = "#a6e3a1"
RED      = "#f38ba8"
YELLOW   = "#f9e2af"
TEAL     = "#94e2d5"
FG       = "#cdd6f4"
FG2      = "#a6adc8"
BORDER   = "#45475a"

# --------------------------------------------------------------------------
# Virtual Key name → VK code mapping
# --------------------------------------------------------------------------
KEY_NAME_TO_VK = {
    "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34,
    "5": 0x35, "6": 0x36, "7": 0x37, "8": 0x38, "9": 0x39,
    **{chr(c): 0x41 + (c - ord("a")) for c in range(ord("a"), ord("z") + 1)},
    "f1":  0x70, "f2":  0x71, "f3":  0x72, "f4":  0x73,
    "f5":  0x74, "f6":  0x75, "f7":  0x76, "f8":  0x77,
    "f9":  0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
    "space": 0x20, "enter": 0x0D, "return": 0x0D,
    "escape": 0x1B, "esc": 0x1B,
    "tab": 0x09, "backspace": 0x08, "delete": 0x2E,
    "shift": 0x10, "ctrl": 0x11, "alt": 0x12,
    "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
    "insert": 0x2D,
    "numpad0": 0x60, "numpad1": 0x61, "numpad2": 0x62, "numpad3": 0x63,
    "numpad4": 0x64, "numpad5": 0x65, "numpad6": 0x66, "numpad7": 0x67,
    "numpad8": 0x68, "numpad9": 0x69,
}


# ==============================================================================
# Layout-independent key sender using Win32 Virtual Key codes
# ==============================================================================
def send_vk_key(vk_code: int, hold_ms: int = 20):
    """
    Press and hold a key for `hold_ms` milliseconds then release.
    Uses the VK code directly — bypasses keyboard layout translation so
    the game always receives the correct key regardless of AZERTY/QWERTY.
    """
    win32api.keybd_event(vk_code, 0, 0, 0)
    time.sleep(max(hold_ms, 20) / 1000.0)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)


# ==============================================================================
# Helper — enumerate visible top-level windows
# ==============================================================================
def get_open_windows():
    results = []
    def _enum_cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title.strip():
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            results.append((hwnd, title, pid, proc.name()))
        except Exception:
            results.append((hwnd, title, 0, "unknown"))
    win32gui.EnumWindows(_enum_cb, None)
    return results

def get_foreground_hwnd():
    return win32gui.GetForegroundWindow()


def force_focus(hwnd: int):
    """
    Aggressively bring a window to the foreground.
    Uses the AttachThreadInput trick so Windows doesn't block the focus request.
    Also restores the window if it was minimised.
    """
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        fg_hwnd = win32gui.GetForegroundWindow()
        if fg_hwnd == hwnd:
            return   # already focused

        # AttachThreadInput trick — only when there is a valid foreground window
        if fg_hwnd:
            fg_tid  = win32process.GetWindowThreadProcessId(fg_hwnd)[0]
            tgt_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
            attached = fg_tid != tgt_tid
            if attached:
                ctypes.windll.user32.AttachThreadInput(fg_tid, tgt_tid, True)
            try:
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
            finally:
                if attached:
                    ctypes.windll.user32.AttachThreadInput(fg_tid, tgt_tid, False)
        else:
            # No current foreground window — just set directly
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


# ==============================================================================
# KeyPresser engine — rapid sequence loop (existing feature)
# ==============================================================================
class KeyPressEngine:
    """Sends a key sequence repeatedly while the target window is focused."""

    def __init__(self, target_hwnd: int, sequence: list, log_cb):
        self.target_hwnd = target_hwnd
        self.sequence    = sequence   # [{"key", "vk", "delay_ms"}, ...]
        self.log_cb      = log_cb
        self.running     = False
        self._thread     = None

    def start(self):
        if self.running:
            return
        self.running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        self.running = False

    def _loop(self):
        self.log_cb("Engine started — sending keys to target window.")
        lost_focus_logged = False

        while self.running:
            fg = get_foreground_hwnd()
            if fg != self.target_hwnd:
                if not lost_focus_logged:
                    self.log_cb("Focus lost — keypresses paused.")
                    lost_focus_logged = True
                time.sleep(0.1)
                continue

            if lost_focus_logged:
                self.log_cb("Focus regained — resuming keypresses.")
                lost_focus_logged = False

            for step in self.sequence:
                if not self.running:
                    break
                if get_foreground_hwnd() != self.target_hwnd:
                    break

                label   = step["key"]
                vk_code = step.get("vk", 0)
                delay_s = step["delay_ms"] / 1000.0

                if vk_code:
                    send_vk_key(vk_code, hold_ms=20)
                    self.log_cb(f"Key '{label}' sent  (VK=0x{vk_code:02X}, {step['delay_ms']} ms)")
                else:
                    self.log_cb(f"Key '{label}' SKIPPED — no VK code. Use the Capture button.")

                time.sleep(delay_s)

        self.log_cb("Engine stopped.")


# ==============================================================================
# Timed Action engine — press-and-hold a key every N minutes
# ==============================================================================
class TimedActionEngine:
    """
    Presses a single key (held for `hold_ms` ms) every `interval_min` minutes.

    `initial_min` sets the delay before the FIRST press — use this when a buff
    is already active with time remaining (e.g. "60-min scroll, 48 min left"
    → set initial_min=48 so the first press fires in 48 min, then every 60 min).

    The engine waits for the target window to be focused before each press.
    If the window is not focused when the timer fires, it retries every 5 s
    until the window is focused, then presses immediately.
    """

    def __init__(self, target_hwnd: int, vk_code: int, hold_ms: int,
                 interval_min: float, initial_min: float, label: str, log_cb,
                 enter_after_ms: int = 0, pre_focus_sec: float = 3.0):
        self.target_hwnd   = target_hwnd
        self.vk_code       = vk_code
        self.hold_ms       = hold_ms
        self.interval_sec  = interval_min * 60.0
        self.label         = label
        self.log_cb        = log_cb
        self.enter_after_ms  = enter_after_ms   # 0 = disabled
        self.pre_focus_sec   = max(1.0, pre_focus_sec)
        self.running       = False
        self._thread       = None
        # next_trigger is a Unix timestamp; readable from the main thread
        # for the countdown display (single float write is atomic in CPython).
        self.next_trigger  = 0.0
        self._initial_sec  = initial_min * 60.0

    def start(self):
        if self.running:
            return
        self.next_trigger = time.time() + self._initial_sec
        self.running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        self.running = False
        self.next_trigger = 0.0

    def _loop(self):
        self.log_cb(
            f"[Timer] '{self.label}' started — first press in "
            f"{self._initial_sec/60:.1f} min, then every {self.interval_sec/60:.1f} min."
        )

        while self.running:
            remaining = self.next_trigger - time.time()

            if remaining > self.pre_focus_sec:
                # Still far from trigger — sleep in short bursts so stop() is responsive
                time.sleep(0.5)
                continue

            if remaining > 0:
                # ── Pre-focus window: aggressively force the game to the foreground ──
                # Called every 100 ms for `pre_focus_sec` seconds before the press.
                force_focus(self.target_hwnd)
                self.log_cb(
                    f"[Timer] '{self.label}' — forcing focus "
                    f"({remaining:.1f}s to press)…"
                )
                time.sleep(0.1)
                continue

            # ── Timer fired — press the key ───────────────────────────────
            if not self.running:
                break

            # One final focus push, then wait up to 500 ms for the OS to
            # confirm the switch before we send any keystrokes.
            force_focus(self.target_hwnd)
            deadline = time.time() + 0.5
            while time.time() < deadline:
                if win32gui.GetForegroundWindow() == self.target_hwnd:
                    break
                time.sleep(0.05)

            if win32gui.GetForegroundWindow() != self.target_hwnd:
                self.log_cb(f"[Timer] '{self.label}' WARNING — window not focused at press time, retrying focus…")
                force_focus(self.target_hwnd)
                time.sleep(0.2)

            # Press ESC first so any open chat box / dialog is dismissed
            # before the actual key lands (prevents accidental chat input).
            send_vk_key(0x1B, hold_ms=50)   # VK_ESCAPE
            time.sleep(0.15)

            send_vk_key(self.vk_code, hold_ms=self.hold_ms)
            self.log_cb(
                f"[Timer] '{self.label}' pressed  "
                f"(VK=0x{self.vk_code:02X}, held {self.hold_ms} ms)"
            )

            # Optional: press Enter N ms after the main key (e.g. to confirm dialogs)
            if self.enter_after_ms > 0 and self.running:
                time.sleep(self.enter_after_ms / 1000.0)
                if self.running:
                    send_vk_key(0x0D, hold_ms=50)   # VK_RETURN
                    self.log_cb(f"[Timer] '{self.label}' — Enter sent "
                                f"({self.enter_after_ms} ms after key)")

            # Schedule the next press
            self.next_trigger = time.time() + self.interval_sec

        self.log_cb(f"[Timer] '{self.label}' stopped.")


# ==============================================================================
# Image-based screen search (OpenCV template matching)
# ==============================================================================
def find_on_screen(template_path: str, threshold: float = 0.7):
    """
    Locate `template_path` on the current screen using normalized cross-correlation.
    Returns (center_x, center_y) in logical screen coordinates, or None if not found.
    """
    if not _IMAGE_MATCHING or not os.path.exists(template_path):
        return None
    try:
        screenshot = np.array(ImageGrab.grab())
        screen_gray = cv2.cvtColor(screenshot, cv2.COLOR_RGB2GRAY)
        template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
        if template is None:
            return None
        result = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            h, w = template.shape
            return (max_loc[0] + w // 2, max_loc[1] + h // 2)
    except Exception:
        pass
    return None


# ==============================================================================
# Alchemy auto-clicker engine
# ==============================================================================
class AlchemyEngine:
    """
    Sends left mouse clicks at a fixed interval for as long as `running` is True.
    Only clicks when the target window is in the foreground.
    """

    def __init__(self, target_hwnd: int, delay_ms: int, log_cb,
                 template_path: str = "",
                 text_region=None, stop_text: str = "",
                 stop_cb=None):
        self.target_hwnd  = target_hwnd
        self.delay_ms     = max(100, delay_ms)
        self.log_cb       = log_cb
        self.template_path = template_path
        self.text_region  = text_region   # [x, y, w, h] or None
        self.stop_text    = stop_text.strip().lower()
        self.stop_cb      = stop_cb       # called (no args) when stop condition met
        self.running      = False
        self._thread: threading.Thread | None = None

    def start(self):
        self.running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        self.running = False

    def _ocr_region(self) -> str:
        """Screenshot the configured text region and return OCR text, or ''."""
        if not _OCR_AVAILABLE or not _IMAGE_MATCHING or not self.text_region:
            return ""
        try:
            x, y, w, h = self.text_region
            img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
            # Upscale + boost contrast so Tesseract handles game fonts better
            img = img.resize((w * 2, h * 2))
            img = ImageEnhance.Contrast(img).enhance(2.5)
            img = img.convert("L")   # greyscale
            return pytesseract.image_to_string(img)
        except Exception as e:
            self.log_cb(f"[Alchemy] OCR error: {e}")
            return ""

    def _loop(self):
        use_template = bool(self.template_path and os.path.exists(self.template_path)
                            and _IMAGE_MATCHING)
        use_ocr_stop = bool(self.stop_text and self.text_region and _OCR_AVAILABLE)
        mode_parts   = []
        if use_template:  mode_parts.append("image targeting")
        if use_ocr_stop:  mode_parts.append(f"stop at '{self.stop_text}'")
        self.log_cb(f"[Alchemy] Started ({', '.join(mode_parts) or 'cursor position'}).")
        _miss_count = 0

        while self.running:
            # ── Click ─────────────────────────────────────────────────────
            if win32gui.GetForegroundWindow() == self.target_hwnd:
                if use_template:
                    pos = find_on_screen(self.template_path)
                    if pos:
                        _miss_count = 0
                        ctypes.windll.user32.SetCursorPos(pos[0], pos[1])
                        time.sleep(0.05)
                        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
                        time.sleep(0.05)
                        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
                    else:
                        _miss_count += 1
                        if _miss_count % 5 == 1:
                            self.log_cb("[Alchemy] Fuse button not found on screen…")
                else:
                    ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
                    time.sleep(0.05)
                    ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)

            # ── Stop-condition check (2 s after click — result text appears then) ─
            if use_ocr_stop and self.running:
                time.sleep(2.0)
                if not self.running:
                    break
                ocr = self._ocr_region()
                if ocr.strip():
                    self.log_cb(f"[Alchemy] OCR: {ocr.strip()[:80]!r}")
                if self.stop_text.lower() in ocr.lower():
                    self.log_cb("[Alchemy] Stop condition met — target reached!")
                    self.running = False
                    if self.stop_cb:
                        self.stop_cb()
                    break
                # Wait out the rest of the delay (delay_ms already includes the 2 s)
                remaining = (self.delay_ms / 1000.0) - 2.0
                if remaining > 0 and self.running:
                    time.sleep(remaining)
            else:
                time.sleep(self.delay_ms / 1000.0)

        self.log_cb("[Alchemy] Auto-clicker stopped.")


# ==============================================================================
# Main Application Window
# ==============================================================================
class KeyPresserApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("Invictus SRO KP")
        self.resizable(False, True)   # allow vertical resize but not horizontal
        self.configure(bg=BG)
        self.geometry("710x800")
        self.minsize(710, 500)

        self.config_data = self._load_config()
        self.engine: KeyPressEngine | None = None

        self._target_hwnd   = 0
        self._window_map    = {}
        self._seq_rows      = []   # (key_var, vk_var, delay_var, frame)
        self._timed_rows    = []   # list of row-dicts (see _add_timed_row)
        self._capturing     = False
        self._alchemy_engine: AlchemyEngine | None = None

        self._build_ui()
        self._populate_window_list()
        self._register_hotkeys()

        saved = self.config_data.get("target_window", "")
        if saved and saved in self._window_map:
            self.window_combo.set(saved)
            self._on_window_selected()

        # Start the per-second countdown ticker for timed actions
        self.after(500, self._tick_countdowns)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ======================================================================
    # UI Construction
    # ======================================================================
    def _build_ui(self):
        # ── Accent stripe (top) ───────────────────────────────────────────
        tk.Frame(self, bg=ACCENT2, height=4).pack(fill="x", side="top")

        # ── Footer (bottom, packed before canvas so it stays pinned) ──────
        footer = tk.Frame(self, bg=BG2, height=28)
        footer.pack(fill="x", side="bottom")

        # Update button — left side
        upd_btn = tk.Button(footer, text="🔄 Check for Update",
                            bg=BG2, fg=FG2, activebackground=BORDER,
                            activeforeground=FG, relief="flat",
                            font=("Segoe UI", 8), cursor="hand2",
                            padx=6, pady=0)
        upd_btn.pack(side="left", padx=8, pady=4)
        upd_btn.config(command=lambda b=upd_btn: self._check_for_updates(b))

        # Credits — right side
        _ft = tk.Frame(footer, bg=BG2)
        _ft.pack(side="right", padx=10, pady=4)
        tk.Label(_ft, text=f"v{APP_VERSION}  —  Made by LinKx with ",
                 font=("Segoe UI", 9), fg=FG2, bg=BG2).pack(side="left")
        tk.Label(_ft, text="\u2665",
                 font=("Segoe UI", 11, "bold"), fg=RED, bg=BG2).pack(side="left")

        # ── Scrollable canvas — all main content lives inside sf ──────────
        _canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
        _vsb = tk.Scrollbar(self, orient="vertical", command=_canvas.yview,
                            bg=BG3, troughcolor=BG2, relief="flat")
        _canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="right", fill="y")
        _canvas.pack(side="left", fill="both", expand=True)

        sf = tk.Frame(_canvas, bg=BG)
        self._sf = sf
        _cw = _canvas.create_window((0, 0), window=sf, anchor="nw")

        sf.bind("<Configure>",
                lambda e: _canvas.configure(scrollregion=_canvas.bbox("all")))
        _canvas.bind("<Configure>",
                     lambda e: _canvas.itemconfig(_cw, width=e.width))
        _canvas.bind_all("<MouseWheel>",
                         lambda e: _canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # ── Title ──────────────────────────────────────────────────────────
        header = tk.Frame(sf, bg=BG)
        header.pack(fill="x", padx=12, pady=(8, 2))
        tk.Label(header, text="Invictus SRO KP",
                 font=("Segoe UI", 16, "bold"), fg=ACCENT2, bg=BG).pack(side="left")

        # ── Target window selector ─────────────────────────────────────────
        sel_frame = tk.LabelFrame(sf, text=" Target Window ", font=("Segoe UI", 9),
                                  fg=ACCENT, bg=BG, bd=1, relief="solid")
        sel_frame.pack(fill="x", padx=12, pady=(6, 4))

        inner_sel = tk.Frame(sel_frame, bg=BG)
        inner_sel.pack(fill="x", padx=8, pady=6)

        self.window_combo = ttk.Combobox(inner_sel, state="readonly", width=48,
                                         font=("Consolas", 9))
        self.window_combo.pack(side="left", padx=(0, 6))
        self.window_combo.bind("<<ComboboxSelected>>", lambda _: self._on_window_selected())

        tk.Button(inner_sel, text="⟳ Refresh", command=self._populate_window_list,
                  bg=BG3, fg=ACCENT, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=3, font=("Segoe UI", 9),
                  cursor="hand2").pack(side="left")

        self._target_label = tk.Label(sel_frame, text="No window selected",
                                      font=("Consolas", 8), fg=FG2, bg=BG)
        self._target_label.pack(anchor="w", padx=10, pady=(0, 4))

        # ── Key sequence editor ────────────────────────────────────────────
        seq_outer = tk.LabelFrame(sf, text=" Key Sequence ", font=("Segoe UI", 9),
                                  fg=ACCENT, bg=BG, bd=1, relief="solid")
        seq_outer.pack(fill="x", padx=12, pady=4)

        # Header widths match the row widgets: idx=3, key entry=7, vk label=7, spinbox=7
        hdr = tk.Frame(seq_outer, bg=BG)
        hdr.pack(fill="x", padx=8, pady=(6, 2))
        for txt, w in [("#", 3), ("Key", 7), ("VK (hex)", 7), ("Delay (ms)", 7), ("Actions", 0)]:
            tk.Label(hdr, text=txt, width=w, anchor="center",
                     fg=FG2, bg=BG, font=("Segoe UI", 8, "bold")).pack(side="left", padx=4)

        self._seq_container = tk.Frame(seq_outer, bg=BG)
        self._seq_container.pack(fill="x", padx=8)

        self._build_sequence_rows()

        row_ctrl = tk.Frame(seq_outer, bg=BG)
        row_ctrl.pack(fill="x", padx=8, pady=6)
        tk.Button(row_ctrl, text="+ Add Key", command=self._add_seq_row,
                  bg=BG3, fg=GREEN, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=3, font=("Segoe UI", 9),
                  cursor="hand2").pack(side="left")
        tk.Button(row_ctrl, text="⏱ Set All Delays", command=self._set_all_delays,
                  bg=BG3, fg=YELLOW, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=3, font=("Segoe UI", 9),
                  cursor="hand2").pack(side="left", padx=(6, 0))
        tk.Label(row_ctrl, text="  ⊙ Capture = click then press the physical key",
                 fg=FG2, bg=BG, font=("Segoe UI", 7, "italic")).pack(side="left", padx=8)

        # ── Key Sequence Controls (start/stop + hotkeys merged) ───────────
        ctrl_frame = tk.LabelFrame(sf, text=" Key Sequence Controls ",
                                   font=("Segoe UI", 9), fg=ACCENT, bg=BG, bd=1, relief="solid")
        ctrl_frame.pack(fill="x", padx=12, pady=4)

        inner_ctrl = tk.Frame(ctrl_frame, bg=BG)
        inner_ctrl.pack(fill="x", padx=10, pady=6)

        # Start / Stop buttons
        self._btn_start = tk.Button(inner_ctrl, text="▶ Start",
                                    command=self.start_pressing,
                                    bg=GREEN, fg="#1e1e2e", activebackground="#80c880",
                                    activeforeground="#1e1e2e", relief="flat",
                                    font=("Segoe UI", 9, "bold"), padx=10, pady=4,
                                    cursor="hand2")
        self._btn_start.pack(side="left", padx=(0, 4))

        self._btn_stop = tk.Button(inner_ctrl, text="■ Stop",
                                   command=self.stop_pressing,
                                   bg=RED, fg="#1e1e2e", activebackground="#d06070",
                                   activeforeground="#1e1e2e", relief="flat",
                                   font=("Segoe UI", 9, "bold"), padx=10, pady=4,
                                   cursor="hand2", state="disabled")
        self._btn_stop.pack(side="left", padx=(0, 6))

        self._status_var = tk.StringVar(value="● Idle")
        self._status_lbl = tk.Label(inner_ctrl, textvariable=self._status_var,
                                    fg=FG2, bg=BG, font=("Segoe UI", 9))
        self._status_lbl.pack(side="left", padx=(0, 10))

        # Vertical divider
        tk.Label(inner_ctrl, text="│", fg=BORDER, bg=BG,
                 font=("Segoe UI", 12)).pack(side="left", padx=(0, 8))

        # Hotkeys inline
        tk.Label(inner_ctrl, text="Start:", fg=FG2, bg=BG,
                 font=("Segoe UI", 8)).pack(side="left")
        self._hk_start_var = tk.StringVar(value=self.config_data.get("hotkey_start", "F6"))
        tk.Entry(inner_ctrl, textvariable=self._hk_start_var, width=5, bg=BG3, fg=ACCENT,
                 insertbackground=FG, relief="flat",
                 font=("Consolas", 8)).pack(side="left", padx=(3, 8))

        tk.Label(inner_ctrl, text="Stop:", fg=FG2, bg=BG,
                 font=("Segoe UI", 8)).pack(side="left")
        self._hk_stop_var = tk.StringVar(value=self.config_data.get("hotkey_stop", "F7"))
        tk.Entry(inner_ctrl, textvariable=self._hk_stop_var, width=5, bg=BG3, fg=ACCENT,
                 insertbackground=FG, relief="flat",
                 font=("Consolas", 8)).pack(side="left", padx=(3, 8))

        tk.Button(inner_ctrl, text="Apply", command=self._register_hotkeys,
                  bg=BG3, fg=YELLOW, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=6, pady=4, font=("Segoe UI", 8),
                  cursor="hand2").pack(side="left", padx=(0, 4))

        # Save on the right
        tk.Button(inner_ctrl, text="💾 Save", command=self.save_config,
                  bg=BG3, fg=ACCENT, activebackground=BORDER, activeforeground=FG,
                  relief="flat", font=("Segoe UI", 8), padx=8, pady=4,
                  cursor="hand2").pack(side="right")

        # ── Timed Actions ──────────────────────────────────────────────────
        self._build_timed_section()

        # ── Alchemy (Auto-Clicker) ─────────────────────────────────────────
        self._build_alchemy_section()

        # ── Log console ────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(sf, text=" Event Log ", font=("Segoe UI", 9),
                                  fg=ACCENT, bg=BG, bd=1, relief="solid")
        log_frame.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        self._log_text = tk.Text(log_frame, height=8, bg=BG2, fg=FG,
                                 font=("Consolas", 8), state="disabled",
                                 relief="flat", wrap="word",
                                 insertbackground=FG, selectbackground=BORDER)
        self._log_text.pack(fill="both", expand=True, padx=6, pady=6, side="left")

        sb = tk.Scrollbar(log_frame, command=self._log_text.yview,
                          bg=BG3, troughcolor=BG2, relief="flat")
        sb.pack(side="right", fill="y", pady=6)
        self._log_text.config(yscrollcommand=sb.set)

        self._log_text.tag_config("ts",    foreground=FG2)
        self._log_text.tag_config("key",   foreground=GREEN)
        self._log_text.tag_config("timer", foreground=TEAL)
        self._log_text.tag_config("warn",  foreground=YELLOW)
        self._log_text.tag_config("error", foreground=RED)
        self._log_text.tag_config("info",  foreground=ACCENT)

    # ======================================================================
    # Timed Actions section
    # ======================================================================
    # Pixel minwidths per column — shared by header row (row 0) and every
    # data row so that grid geometry guarantees perfect alignment.
    _TA_COL_W = [28, 60, 68, 70, 70, 30, 90, 28, 52, 28, 28, 28]
    #             #  Key Hold Every Next  unit Cdown ↵?  ↵ms  ▶   ■   ✕

    def _build_timed_section(self):
        ta_outer = tk.LabelFrame(self._sf, text=" Timed Actions  (press & hold a key every N minutes) ",
                                 font=("Segoe UI", 9), fg=TEAL, bg=BG, bd=1, relief="solid")
        ta_outer.pack(fill="x", padx=12, pady=4)

        # ── Single shared grid — header is row 0, data rows are 1, 2, … ──
        # Because everything lives in the SAME Frame with grid(), column 0 in
        # the header label is always exactly as wide as column 0 in every row.
        self._timed_grid = tk.Frame(ta_outer, bg=BG)
        self._timed_grid.pack(fill="x", padx=8, pady=(4, 0))

        for col, minw in enumerate(self._TA_COL_W):
            self._timed_grid.columnconfigure(col, minsize=minw)

        headers = ["#", "Key", "Hold\n(ms)", "Every\n(min)", "Next in",
                   "", "Countdown", "↵?", "↵ ms", "▶", "■", "✕"]
        for col, txt in enumerate(headers):
            tk.Label(self._timed_grid, text=txt, anchor="center",
                     fg=FG2, bg=BG,
                     font=("Segoe UI", 8, "bold")).grid(
                row=0, column=col, padx=1, pady=(4, 2), sticky="ew")

        # Build saved rows
        for ta in self.config_data.get("timed_actions", []):
            self._add_timed_row(
                key=ta.get("key", ""),
                vk=ta.get("vk", 0),
                hold_ms=ta.get("hold_ms", 200),
                interval_min=ta.get("interval_min", 60),
                initial_min=ta.get("initial_min", 60),
                enter_after_ms=ta.get("enter_after_ms", 0),
                initial_unit=ta.get("initial_unit", "min"),
            )

        ctrl = tk.Frame(ta_outer, bg=BG)
        ctrl.pack(fill="x", padx=8, pady=6)
        tk.Button(ctrl, text="+ Add Timed Action", command=self._add_timed_row,
                  bg=BG3, fg=TEAL, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=3, font=("Segoe UI", 9),
                  cursor="hand2").pack(side="left")
        tk.Button(ctrl, text="▶ Start All", command=self._start_all_timed,
                  bg=BG3, fg=GREEN, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=3, font=("Segoe UI", 9),
                  cursor="hand2").pack(side="left", padx=(6, 0))
        tk.Button(ctrl, text="■ Stop All", command=self._stop_all_timed,
                  bg=BG3, fg=RED, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=3, font=("Segoe UI", 9),
                  cursor="hand2").pack(side="left", padx=(4, 0))

        # Pre-focus seconds input
        tk.Label(ctrl, text="  Pre-focus:", fg=FG2, bg=BG,
                 font=("Segoe UI", 8)).pack(side="left", padx=(8, 2))
        self._prefocus_var = tk.IntVar(value=self.config_data.get("prefocus_sec", 3))
        _vcmd_pf = (self.register(lambda s: s.isdigit() or s == ''), '%P')
        tk.Spinbox(ctrl, from_=1, to=30, textvariable=self._prefocus_var, width=3,
                   bg=BG3, fg=TEAL, insertbackground=FG, buttonbackground=BG3,
                   relief="flat", font=("Consolas", 9),
                   validate='key', validatecommand=_vcmd_pf
                   ).pack(side="left")
        tk.Label(ctrl, text="s", fg=FG2, bg=BG,
                 font=("Segoe UI", 8)).pack(side="left", padx=(2, 0))

        # Live/idle indicator — updated every 500 ms by _tick_countdowns()
        self._timed_status_lbl = tk.Label(ctrl, text="● Idle",
                                          fg=RED, bg=BG, font=("Segoe UI", 8, "bold"))
        self._timed_status_lbl.pack(side="right", padx=4)

    # ======================================================================
    # Alchemy section
    # ======================================================================
    def _build_alchemy_section(self):
        outer = tk.LabelFrame(self._sf, text=" Alchemy  (Auto Left-Clicker) ",
                              font=("Segoe UI", 9), fg="#fab387", bg=BG, bd=1, relief="solid")
        outer.pack(fill="x", padx=12, pady=4)

        row = tk.Frame(outer, bg=BG)
        row.pack(fill="x", padx=10, pady=8)

        # Delay
        tk.Label(row, text="Delay (ms):", fg=FG2, bg=BG,
                 font=("Segoe UI", 9)).pack(side="left")
        _vcmd = (self.register(lambda s: s.isdigit() or s == ''), '%P')
        self._alchemy_delay_var = tk.IntVar(
            value=self.config_data.get("alchemy_delay_ms", 500))
        tk.Spinbox(row, from_=100, to=99999, textvariable=self._alchemy_delay_var,
                   width=6, bg=BG3, fg=FG, insertbackground=FG, buttonbackground=BG3,
                   relief="flat", font=("Consolas", 9),
                   validate='key', validatecommand=_vcmd
                   ).pack(side="left", padx=(4, 16))

        # Toggle hotkey
        tk.Label(row, text="Toggle key:", fg=FG2, bg=BG,
                 font=("Segoe UI", 9)).pack(side="left")
        self._alchemy_hk_var = tk.StringVar(
            value=self.config_data.get("alchemy_hotkey", "F8"))
        tk.Entry(row, textvariable=self._alchemy_hk_var, width=5,
                 bg=BG3, fg=ACCENT, insertbackground=FG, relief="flat",
                 font=("Consolas", 9)).pack(side="left", padx=(4, 4))
        tk.Button(row, text="Apply", command=self._register_hotkeys,
                  bg=BG3, fg=YELLOW, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=6, pady=2, font=("Segoe UI", 8),
                  cursor="hand2").pack(side="left", padx=(0, 16))

        # Start / Stop toggle button
        self._alchemy_btn = tk.Button(row, text="▶ Start",
                                      command=self.toggle_alchemy,
                                      bg="#fab387", fg="#1e1e2e",
                                      activebackground="#fcd9c4",
                                      activeforeground="#1e1e2e",
                                      relief="flat", font=("Segoe UI", 9, "bold"),
                                      padx=10, pady=4, cursor="hand2")
        self._alchemy_btn.pack(side="left", padx=(0, 8))

        # Status label
        self._alchemy_status_var = tk.StringVar(value="● Idle")
        tk.Label(row, textvariable=self._alchemy_status_var,
                 fg=FG2, bg=BG, font=("Segoe UI", 9)).pack(side="left")

        # ── Row 2: Fuse button image targeting ────────────────────────────
        row2 = tk.Frame(outer, bg=BG)
        row2.pack(fill="x", padx=10, pady=(0, 4))

        tk.Button(row2, text="📷 Set Fuse",
                  command=self._alchemy_capture_template,
                  bg=BG3, fg="#fab387", activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=2, font=("Segoe UI", 8),
                  cursor="hand2").pack(side="left")
        tk.Label(row2, text="  Draw a rectangle around the Fuse button",
                 fg=FG2, bg=BG, font=("Segoe UI", 7, "italic")).pack(side="left", padx=4)

        has_tpl = os.path.exists(ALCHEMY_TEMPLATE_PATH)
        self._alchemy_template_lbl = tk.Label(
            row2,
            text="✓ Template set" if has_tpl else "No template — clicking at cursor",
            fg=GREEN if has_tpl else FG2, bg=BG, font=("Segoe UI", 8))
        self._alchemy_template_lbl.pack(side="right", padx=4)

        # ── Row 3: OCR stop condition ──────────────────────────────────────
        row3 = tk.Frame(outer, bg=BG)
        row3.pack(fill="x", padx=10, pady=(0, 8))

        tk.Button(row3, text="📷 Set Text Area",
                  command=self._alchemy_capture_text_region,
                  bg=BG3, fg=TEAL, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=8, pady=2, font=("Segoe UI", 8),
                  cursor="hand2").pack(side="left")
        tk.Label(row3, text="  Stop when:", fg=FG2, bg=BG,
                 font=("Segoe UI", 8)).pack(side="left", padx=(8, 2))
        self._alchemy_stop_text_var = tk.StringVar(
            value=self.config_data.get("alchemy_stop_text", "changed to [5]."))
        tk.Entry(row3, textvariable=self._alchemy_stop_text_var, width=28,
                 bg=BG3, fg=TEAL, insertbackground=FG, relief="flat",
                 font=("Consolas", 8)).pack(side="left", padx=(0, 8))

        has_rgn = bool(self.config_data.get("alchemy_text_region"))
        self._alchemy_region_lbl = tk.Label(
            row3,
            text="✓ Region set" if has_rgn else "No region",
            fg=GREEN if has_rgn else FG2, bg=BG, font=("Segoe UI", 8))
        self._alchemy_region_lbl.pack(side="left")

        tk.Button(row3, text="🔍 Test OCR",
                  command=self._alchemy_test_ocr,
                  bg=BG3, fg=FG2, activebackground=BORDER, activeforeground=FG,
                  relief="flat", padx=6, pady=2, font=("Segoe UI", 8),
                  cursor="hand2").pack(side="left", padx=(8, 0))

        if not _OCR_AVAILABLE:
            tk.Label(row3, text="  ⚠ pytesseract not installed",
                     fg=YELLOW, bg=BG, font=("Segoe UI", 7, "italic")).pack(side="left", padx=4)

    def toggle_alchemy(self):
        if self._alchemy_engine and self._alchemy_engine.running:
            self.stop_alchemy()
        else:
            self.start_alchemy()

    def start_alchemy(self):
        if not self._target_hwnd:
            messagebox.showwarning("No Target", "Please select a target window first.")
            return
        if self._alchemy_engine and self._alchemy_engine.running:
            return
        delay_ms = max(100, self._alchemy_delay_var.get())
        self._alchemy_engine = AlchemyEngine(
            target_hwnd   = self._target_hwnd,
            delay_ms      = delay_ms,
            log_cb        = lambda msg: self._log(msg, "info"),
            template_path = ALCHEMY_TEMPLATE_PATH,
            text_region   = self.config_data.get("alchemy_text_region"),
            stop_text     = self._alchemy_stop_text_var.get(),
            stop_cb       = lambda: self.after(0, self.stop_alchemy),
        )
        self._alchemy_engine.start()
        self._alchemy_btn.config(text="■ Stop")
        self._alchemy_status_var.set("● Running")

    def stop_alchemy(self):
        if self._alchemy_engine:
            self._alchemy_engine.stop()
            self._alchemy_engine = None
        self._alchemy_btn.config(text="▶ Start")
        self._alchemy_status_var.set("● Idle")

    # ── Shared snipping-tool overlay ──────────────────────────────────────
    def _show_region_picker(self, callback):
        """
        Minimise the app, grab a screenshot, then show a full-screen canvas
        where the user drags to select a rectangle.
        callback(x, y, w, h, screenshot) is called on the main thread.
        """
        if not _IMAGE_MATCHING:
            messagebox.showerror("Missing Libraries",
                "Install Pillow, numpy, and opencv-python:\n"
                "  pip install Pillow numpy opencv-python")
            return
        self.iconify()
        self.after(250, lambda: self._open_picker_overlay(callback))

    def _open_picker_overlay(self, callback):
        screenshot = ImageGrab.grab()
        sw, sh = screenshot.size

        top = tk.Toplevel(self)
        top.overrideredirect(True)
        top.geometry(f"{sw}x{sh}+0+0")
        top.attributes('-topmost', True)

        photo = ImageTk.PhotoImage(screenshot)
        canvas = tk.Canvas(top, width=sw, height=sh, cursor='crosshair',
                           highlightthickness=0, bd=0)
        canvas.pack(fill='both', expand=True)
        canvas.create_image(0, 0, anchor='nw', image=photo)
        canvas.image = photo   # keep reference

        DIM, SEL, HINT, SZ = 'dim', 'sel', 'hint', 'sz'

        def _redraw_dim(x0=0, y0=0, x1=sw, y1=sh):
            """Draw dim overlay as 4 rectangles with the selection area clear."""
            canvas.delete(DIM)
            for rx0, ry0, rx1, ry1 in [
                (0,  0,   sw,  y0),   # top
                (0,  y1,  sw,  sh),   # bottom
                (0,  y0,  x0,  y1),   # left
                (x1, y0,  sw,  y1),   # right
            ]:
                if rx1 > rx0 and ry1 > ry0:
                    canvas.create_rectangle(rx0, ry0, rx1, ry1,
                                            fill='black', stipple='gray50',
                                            outline='', tags=DIM)

        _redraw_dim()   # full-screen dim before first drag
        canvas.create_text(sw // 2, sh // 2,
                           text='Drag to select   •   ESC to cancel',
                           fill='white', font=('Segoe UI', 16, 'bold'), tags=HINT)

        start = [0, 0]

        def on_press(e):
            start[0], start[1] = e.x, e.y
            canvas.delete(HINT)
            canvas.delete(SEL)
            canvas.delete(SZ)

        def on_drag(e):
            x0, y0 = min(start[0], e.x), min(start[1], e.y)
            x1, y1 = max(start[0], e.x), max(start[1], e.y)
            _redraw_dim(x0, y0, x1, y1)
            canvas.delete(SEL)
            canvas.create_rectangle(x0, y0, x1, y1,
                                    outline='#00ff88', width=2, tags=SEL)
            canvas.delete(SZ)
            label_y = max(y0 - 16, 12)
            canvas.create_text(x0 + (x1 - x0) // 2, label_y,
                                text=f'{x1 - x0} × {y1 - y0} px',
                                fill='#00ff88', font=('Consolas', 9, 'bold'), tags=SZ)

        def on_release(e):
            x0 = min(start[0], e.x)
            y0 = min(start[1], e.y)
            x1 = max(start[0], e.x)
            y1 = max(start[1], e.y)
            top.destroy()
            self.deiconify()
            if x1 - x0 > 5 and y1 - y0 > 5:
                self.after(0, lambda: callback(x0, y0, x1 - x0, y1 - y0, screenshot))
            else:
                self._log("[Alchemy] Region selection too small — cancelled.", "warn")

        def on_escape(e):
            top.destroy()
            self.deiconify()
            self._log("[Alchemy] Region selection cancelled.", "info")

        canvas.bind('<ButtonPress-1>',   on_press)
        canvas.bind('<B1-Motion>',       on_drag)
        canvas.bind('<ButtonRelease-1>', on_release)
        top.bind('<Escape>',             on_escape)
        top.focus_force()

    def _alchemy_capture_template(self):
        if self._alchemy_engine and self._alchemy_engine.running:
            messagebox.showwarning("Alchemy Running",
                                   "Stop the auto-clicker before capturing a new template.")
            return

        def _on_selected(x, y, w, h, screenshot):
            cropped = screenshot.crop((x, y, x + w, y + h))
            cropped.save(ALCHEMY_TEMPLATE_PATH)
            self._alchemy_template_lbl.config(text="✓ Template set", fg=GREEN)
            self._log(f"[Alchemy] Fuse template saved ({w}×{h} px).", "info")

        self._show_region_picker(_on_selected)

    def _alchemy_capture_text_region(self):
        if self._alchemy_engine and self._alchemy_engine.running:
            messagebox.showwarning("Alchemy Running",
                                   "Stop the auto-clicker before changing the text region.")
            return
        if not _OCR_AVAILABLE:
            messagebox.showerror("pytesseract not installed",
                "Install pytesseract + Tesseract OCR to use the stop condition:\n"
                "  1. pip install pytesseract\n"
                "  2. https://github.com/UB-Mannheim/tesseract/wiki")
            return

        def _on_selected(x, y, w, h, screenshot):
            self.config_data["alchemy_text_region"] = [x, y, w, h]
            self._alchemy_region_lbl.config(text="✓ Region set", fg=GREEN)
            self._log(f"[Alchemy] Text region set ({w}×{h} px at {x},{y}).", "info")

        self._show_region_picker(_on_selected)

    def _alchemy_test_ocr(self):
        """Immediately screenshot the text region and log what OCR reads."""
        region = self.config_data.get("alchemy_text_region")
        if not region:
            self._log("[Alchemy] No text region set — use 📷 Set Text Area first.", "warn")
            return
        if not _OCR_AVAILABLE or not _IMAGE_MATCHING:
            self._log("[Alchemy] OCR not available (pytesseract / Pillow not installed).", "warn")
            return
        try:
            x, y, w, h = region
            img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
            img = img.resize((w * 2, h * 2))
            img = ImageEnhance.Contrast(img).enhance(2.5)
            img = img.convert("L")
            text = pytesseract.image_to_string(img)
            result = text.strip()
            if result:
                self._log(f"[Alchemy] Test OCR: {result[:120]!r}", "info")
            else:
                self._log("[Alchemy] Test OCR: (no text detected)", "warn")
        except Exception as e:
            self._log(f"[Alchemy] Test OCR error: {e}", "warn")

    def _add_timed_row(self, key: str = "", vk: int = 0,
                       hold_ms: int = 200, interval_min: int = 60,
                       initial_min: int = 60, enter_after_ms: int = 0,
                       initial_unit: str = "min"):
        # grid row = position + 1 (row 0 is the header)
        grid_row = len(self._timed_rows) + 1

        # ── Index label ───────────────────────────────────────────────────
        idx_lbl = tk.Label(self._timed_grid, text=str(grid_row), anchor="center",
                           fg=FG2, bg=BG, font=("Consolas", 9))
        idx_lbl.grid(row=grid_row, column=0, padx=1, pady=2, sticky="ew")

        # ── Key entry (read-only, click-to-capture) ───────────────────────
        # No free-text input allowed — clicking triggers physical key capture.
        key_var = tk.StringVar(value=key if key else "⊙ key")
        key_entry = tk.Entry(self._timed_grid, textvariable=key_var,
                             width=6,
                             bg=BG3, fg=YELLOW if not key else FG,
                             insertbackground=FG, relief="flat",
                             font=("Consolas", 9), justify="center",
                             state="readonly", cursor="hand2",
                             readonlybackground=BG3)
        key_entry.grid(row=grid_row, column=1, padx=1, pady=2, sticky="ew")

        # Digit-only validation command — shared by all numeric spinboxes in this row
        _vcmd = (self.register(lambda s: s.isdigit() or s == ''), '%P')

        # ── Hold spinbox ──────────────────────────────────────────────────
        hold_var = tk.IntVar(value=hold_ms)
        hold_spin = tk.Spinbox(self._timed_grid, from_=50, to=9999,
                               textvariable=hold_var, width=5,
                               bg=BG3, fg=FG, insertbackground=FG,
                               buttonbackground=BG3, relief="flat",
                               font=("Consolas", 9),
                               validate='key', validatecommand=_vcmd)
        hold_spin.grid(row=grid_row, column=2, padx=1, pady=2, sticky="ew")

        # ── Interval spinbox ──────────────────────────────────────────────
        interval_var = tk.IntVar(value=interval_min)
        interval_spin = tk.Spinbox(self._timed_grid, from_=1, to=9999,
                                   textvariable=interval_var, width=5,
                                   bg=BG3, fg=FG, insertbackground=FG,
                                   buttonbackground=BG3, relief="flat",
                                   font=("Consolas", 9),
                                   validate='key', validatecommand=_vcmd)
        interval_spin.grid(row=grid_row, column=3, padx=1, pady=2, sticky="ew")

        # ── "Next in" spinbox (col 4) ─────────────────────────────────────
        # Value is in whatever unit the toggle says (min or sec).
        initial_display = round(initial_min * 60) if initial_unit == "sec" else round(initial_min)
        initial_var = tk.IntVar(value=initial_display)
        initial_spin = tk.Spinbox(self._timed_grid, from_=0, to=9999,
                                  textvariable=initial_var, width=5,
                                  bg=BG3, fg=YELLOW, insertbackground=FG,
                                  buttonbackground=BG3, relief="flat",
                                  font=("Consolas", 9),
                                  validate='key', validatecommand=_vcmd)
        initial_spin.grid(row=grid_row, column=4, padx=1, pady=2, sticky="ew")

        # ── Unit toggle button (col 5) ────────────────────────────────────
        # Clicking switches "Next in" between minutes and seconds.
        initial_unit_var = tk.StringVar(value=initial_unit)
        unit_btn = tk.Button(self._timed_grid,
                             text=initial_unit,
                             fg=TEAL if initial_unit == "sec" else YELLOW,
                             bg=BG3, activebackground=BORDER, activeforeground=FG,
                             relief="flat", font=("Consolas", 8), cursor="hand2")
        unit_btn.grid(row=grid_row, column=5, padx=1, pady=2, sticky="ew")

        def _toggle_unit(uv=initial_unit_var, ub=unit_btn):
            nxt = "sec" if uv.get() == "min" else "min"
            uv.set(nxt)
            ub.config(text=nxt, fg=TEAL if nxt == "sec" else YELLOW)

        unit_btn.config(command=_toggle_unit)

        # ── Live countdown label (col 6) ──────────────────────────────────
        # Ticked every 500 ms by _tick_countdowns(); colour shifts
        # TEAL → YELLOW → RED as the next press approaches.
        countdown_var = tk.StringVar(value="——:——")
        cd_lbl = tk.Label(self._timed_grid, textvariable=countdown_var,
                          anchor="center", fg=TEAL, bg=BG,
                          font=("Consolas", 9, "bold"))
        cd_lbl.grid(row=grid_row, column=6, padx=1, pady=2, sticky="ew")

        # ── Enter-after checkbox (col 7) ──────────────────────────────────
        # When checked, the engine sends Enter N ms after the main key press.
        enter_var = tk.BooleanVar(value=bool(enter_after_ms))
        enter_chk = tk.Checkbutton(self._timed_grid, variable=enter_var,
                                   bg=BG, activebackground=BG,
                                   selectcolor=BG3, relief="flat",
                                   cursor="hand2")
        enter_chk.grid(row=grid_row, column=7, padx=1, pady=2)

        # ── Enter-delay spinbox (col 8) ───────────────────────────────────
        # Active value in ms; only used when enter_var is True.
        enter_ms_val = enter_after_ms if enter_after_ms > 0 else 1500
        enter_ms_var = tk.IntVar(value=enter_ms_val)
        enter_ms_spin = tk.Spinbox(self._timed_grid, from_=100, to=9999,
                                   textvariable=enter_ms_var, width=4,
                                   bg=BG3, fg=TEAL, insertbackground=FG,
                                   buttonbackground=BG3, relief="flat",
                                   font=("Consolas", 9),
                                   validate='key', validatecommand=_vcmd)
        enter_ms_spin.grid(row=grid_row, column=8, padx=1, pady=2, sticky="ew")

        # ── VK code (internal only, not displayed) ───────────────────────
        vk_var     = tk.IntVar(value=vk)
        vk_lbl_var = tk.StringVar(value=f"0x{vk:02X}" if vk else "—")

        # ── Start button (col 9) ──────────────────────────────────────────
        start_btn = tk.Button(self._timed_grid, text="▶",
                              bg=BG3, fg=GREEN, activebackground=BORDER,
                              activeforeground=GREEN, relief="flat",
                              font=("Segoe UI", 9, "bold"), cursor="hand2")
        start_btn.grid(row=grid_row, column=9, padx=1, pady=2, sticky="ew")

        # ── Stop button (col 10) ──────────────────────────────────────────
        stop_btn = tk.Button(self._timed_grid, text="■",
                             bg=BG3, fg=RED, activebackground=BORDER,
                             activeforeground=RED, relief="flat",
                             font=("Segoe UI", 9, "bold"), cursor="hand2",
                             state="disabled")
        stop_btn.grid(row=grid_row, column=10, padx=1, pady=2, sticky="ew")

        # ── Delete button (col 11) ────────────────────────────────────────
        del_btn = tk.Button(self._timed_grid, text="✕",
                            bg=BG, fg=RED, activebackground=BORDER,
                            activeforeground=RED, relief="flat",
                            font=("Segoe UI", 9), cursor="hand2")
        del_btn.grid(row=grid_row, column=11, padx=1, pady=2, sticky="ew")

        # ── Row data dict ──────────────────────────────────────────────────
        row = {
            "key_var":          key_var,
            "vk_var":           vk_var,
            "vk_lbl_var":       vk_lbl_var,
            "hold_var":         hold_var,
            "interval_var":     interval_var,
            "initial_var":      initial_var,
            "initial_unit_var": initial_unit_var,  # "min" or "sec"
            "countdown_var":    countdown_var,
            "cd_lbl":           cd_lbl,        # direct ref for colour changes
            "enter_var":        enter_var,      # BooleanVar — send Enter?
            "enter_ms_var":     enter_ms_var,   # IntVar — delay before Enter
            "start_btn":        start_btn,
            "stop_btn":         stop_btn,
            "engine":           None,
            "idx_lbl":          idx_lbl,
            # All grid widgets for this row — used during deletion
            "_grid_widgets": [idx_lbl, key_entry, hold_spin, interval_spin,
                              initial_spin, unit_btn, cd_lbl, enter_chk,
                              enter_ms_spin, start_btn, stop_btn, del_btn],
        }
        self._timed_rows.append(row)

        # Wire up buttons
        start_btn.config(command=lambda r=row: self._start_timed(r))
        stop_btn.config(command=lambda r=row: self._stop_timed(r))

        # Clicking the key field triggers capture (read-only — no free typing)
        key_entry.bind("<Button-1>", lambda e, kv=key_var, vv=vk_var,
                       vlv=vk_lbl_var, en=key_entry:
                       self._run_capture_entry(kv, vv, vlv, en))

        def _delete(r=row):
            self._stop_timed(r)
            for w in r["_grid_widgets"]:
                w.grid_forget()
                w.destroy()
            self._timed_rows.remove(r)
            self._regrid_timed_rows()   # close the gap, fix index labels

        del_btn.config(command=_delete)

    def _regrid_timed_rows(self):
        """Re-place remaining rows after a deletion (keeps indices sequential)."""
        for i, row in enumerate(self._timed_rows, 1):
            row["idx_lbl"].config(text=str(i))
            for w in row["_grid_widgets"]:
                w.grid_configure(row=i)

    # ── Per-row start / stop ───────────────────────────────────────────────
    def _start_timed(self, row: dict):
        if not self._target_hwnd:
            messagebox.showwarning("No Target", "Please select a target window first.")
            return
        if row["engine"] is not None:
            return   # already running

        raw_key = row["key_var"].get().strip()
        vk = row["vk_var"].get()
        if not vk and raw_key not in ("", "⊙ key"):
            vk = KEY_NAME_TO_VK.get(raw_key.lower(), 0)
        if not vk:
            messagebox.showwarning(
                "No Key Captured",
                "Click the key field (⊙ key) and press the physical key you want to assign first."
            )
            return

        label          = raw_key if raw_key not in ("", "⊙ key") else f"VK=0x{vk:02X}"
        hold_ms        = max(50, row["hold_var"].get())
        interval_min   = max(1, row["interval_var"].get())
        raw_initial    = max(0, row["initial_var"].get())
        initial_min    = raw_initial / 60.0 if row["initial_unit_var"].get() == "sec" else float(raw_initial)
        enter_after_ms = row["enter_ms_var"].get() if row["enter_var"].get() else 0
        pre_focus_sec  = max(1, self._prefocus_var.get())

        engine = TimedActionEngine(
            target_hwnd    = self._target_hwnd,
            vk_code        = vk,
            hold_ms        = hold_ms,
            interval_min   = interval_min,
            initial_min    = initial_min,
            label          = label,
            log_cb         = lambda msg: self._log(msg, "timer"),
            enter_after_ms = enter_after_ms,
            pre_focus_sec  = pre_focus_sec,
        )
        engine.start()
        row["engine"] = engine

        row["start_btn"].config(state="disabled")
        row["stop_btn"].config(state="normal")
        row["countdown_var"].set("starting…")
        self._log(
            f"[Timer] '{label}' armed — first press in {initial_min} min, "
            f"hold {hold_ms} ms, repeat every {interval_min} min.",
            "timer",
        )

    def _stop_timed(self, row: dict):
        if row["engine"] is not None:
            row["engine"].stop()
            row["engine"] = None
        row["start_btn"].config(state="normal")
        row["stop_btn"].config(state="disabled")
        row["countdown_var"].set("——:——")
        row["cd_lbl"].config(fg=TEAL)

    def _start_all_timed(self):
        for row in self._timed_rows:
            self._start_timed(row)

    def _stop_all_timed(self):
        for row in self._timed_rows:
            self._stop_timed(row)

    # ── Countdown ticker (runs every 500 ms on the main thread) ───────────
    def _tick_countdowns(self):
        any_live = False
        for row in self._timed_rows:
            eng = row["engine"]
            if eng and eng.running:
                any_live = True
            if eng and eng.running and eng.next_trigger > 0:
                remaining = max(0, int(eng.next_trigger - time.time()))
                h, rem = divmod(remaining, 3600)
                m, s   = divmod(rem, 60)
                if h > 0:
                    row["countdown_var"].set(f"{h}h {m:02d}m {s:02d}s")
                else:
                    row["countdown_var"].set(f"  {m:02d}:{s:02d}  ")

                # Colour: TEAL → YELLOW (< 10% left) → RED (≤ 30 s, about to fire)
                interval_sec = eng.interval_sec
                if remaining <= 30:
                    colour = RED
                elif interval_sec > 0 and remaining < interval_sec * 0.1:
                    colour = YELLOW
                else:
                    colour = TEAL
                row["cd_lbl"].config(fg=colour)

        # Update the timed-actions live indicator
        if any_live:
            self._timed_status_lbl.config(text="● Live", fg=GREEN)
        else:
            self._timed_status_lbl.config(text="● Idle", fg=RED)

        self.after(500, self._tick_countdowns)

    # ======================================================================
    # Key sequence row management (existing)
    # ======================================================================
    def _build_sequence_rows(self):
        for step in self.config_data["sequence"]:
            self._add_seq_row(
                key=step.get("key", ""),
                vk=step.get("vk", 0),
                delay_ms=step.get("delay_ms", 50),
            )

    def _add_seq_row(self, key: str = "", vk: int = 0, delay_ms: int = 50):
        idx = len(self._seq_rows) + 1

        frame = tk.Frame(self._seq_container, bg=BG2, pady=3)
        frame.pack(fill="x", pady=2)

        tk.Label(frame, text=str(idx), width=3, anchor="center",
                 fg=FG2, bg=BG2, font=("Consolas", 9)).pack(side="left")

        # Read-only — only the Capture button can populate this field
        key_var = tk.StringVar(value=key)
        key_entry = tk.Entry(frame, textvariable=key_var, width=7,
                             bg=BG3, fg=FG, insertbackground=FG, relief="flat",
                             font=("Consolas", 10), justify="center",
                             state="readonly", cursor="hand2",
                             readonlybackground=BG3)
        key_entry.pack(side="left", padx=4)

        vk_var     = tk.IntVar(value=vk)
        vk_lbl_var = tk.StringVar(value=f"0x{vk:02X}" if vk else "—")
        tk.Label(frame, textvariable=vk_lbl_var, width=7, anchor="center",
                 fg=ACCENT2, bg=BG2, font=("Consolas", 9)).pack(side="left", padx=4)

        _vcmd_seq = (self.register(lambda s: s.isdigit() or s == ''), '%P')
        delay_var = tk.IntVar(value=delay_ms)
        tk.Spinbox(frame, from_=1, to=9999, textvariable=delay_var, width=7,
                   bg=BG3, fg=FG, insertbackground=FG, buttonbackground=BG3,
                   relief="flat", font=("Consolas", 10),
                   validate='key', validatecommand=_vcmd_seq).pack(side="left", padx=4)

        capture_btn = tk.Button(frame, text="⊙ Capture",
                                bg=BG3, fg=YELLOW, activebackground=BORDER,
                                activeforeground=FG, relief="flat",
                                font=("Segoe UI", 8), cursor="hand2", padx=4)
        capture_btn.pack(side="left", padx=2)
        # Clicking the key field itself also triggers capture
        key_entry.bind("<Button-1>", lambda e, b=capture_btn: b.invoke())

        del_btn = tk.Button(frame, text="✕", bg=BG2, fg=RED, activebackground=BORDER,
                            activeforeground=RED, relief="flat",
                            font=("Segoe UI", 9), cursor="hand2", padx=4)
        del_btn.pack(side="left", padx=2)

        row_tuple = (key_var, vk_var, delay_var, frame)
        self._seq_rows.append(row_tuple)

        def make_delete(f, rt):
            def _d():
                f.pack_forget(); f.destroy()
                self._seq_rows.remove(rt)
                self._renumber_seq_rows()
            return _d

        del_btn.config(command=make_delete(frame, row_tuple))
        capture_btn.config(command=lambda kv=key_var, vv=vk_var, vlv=vk_lbl_var, b=capture_btn:
                           self._run_capture(kv, vv, vlv, b))

    def _renumber_seq_rows(self):
        for i, (_, _, _, frame) in enumerate(self._seq_rows, 1):
            for child in frame.winfo_children():
                if isinstance(child, tk.Label):
                    child.config(text=str(i))
                    break

    def _set_all_delays(self):
        from tkinter import simpledialog
        val = simpledialog.askinteger(
            "Set All Delays",
            "Enter delay (ms) to apply to every key in the sequence:",
            initialvalue=50, minvalue=1, maxvalue=9999, parent=self,
        )
        if val is not None:
            for _, _, delay_var, _ in self._seq_rows:
                delay_var.set(val)

    # ======================================================================
    # Key capture — shared by both seq rows and timed rows
    # ======================================================================
    def _run_capture(self, key_var, vk_var, vk_lbl_var, btn):
        if self._capturing:
            return
        self._capturing = True
        orig = btn.cget("text")
        btn.config(text="▸ Press key…", fg=GREEN, state="disabled")
        self._log("Capture: press the physical key you want to assign…", "info")

        def _thread():
            try:
                event = keyboard.read_event(suppress=True)
                while event.event_type != keyboard.KEY_DOWN:
                    event = keyboard.read_event(suppress=True)

                name      = event.name
                scan_code = event.scan_code
                vk_code   = ctypes.windll.user32.MapVirtualKeyW(scan_code, 3)
                if vk_code == 0:
                    vk_code = KEY_NAME_TO_VK.get(name.lower(), 0)

                self.after(0, _apply, name, vk_code)
            except Exception as e:
                self.after(0, lambda: self._log(f"Capture error: {e}", "error"))
                self.after(0, _restore)

        def _apply(name, vk_code):
            key_var.set(name)
            vk_var.set(vk_code)
            vk_lbl_var.set(f"0x{vk_code:02X}" if vk_code else "—")
            self._log(f"Captured: '{name}'  VK=0x{vk_code:02X}", "info")
            _restore()

        def _restore():
            btn.config(text=orig, fg=YELLOW, state="normal")
            self._capturing = False

        threading.Thread(target=_thread, daemon=True).start()

    def _run_capture_entry(self, key_var, vk_var, vk_lbl_var, entry):
        """Capture variant for read-only Entry widgets (timed action rows).
        The entry itself acts as the visual feedback element."""
        if self._capturing:
            return
        self._capturing = True
        orig_text = key_var.get()
        orig_fg   = entry.cget("readonlybackground") and FG  # keep ref
        key_var.set("▸ press…")
        entry.config(fg=GREEN)
        self._log("Capture: press the physical key you want to assign…", "info")

        def _thread():
            try:
                event = keyboard.read_event(suppress=True)
                while event.event_type != keyboard.KEY_DOWN:
                    event = keyboard.read_event(suppress=True)

                name      = event.name
                scan_code = event.scan_code
                vk_code   = ctypes.windll.user32.MapVirtualKeyW(scan_code, 3)
                if vk_code == 0:
                    vk_code = KEY_NAME_TO_VK.get(name.lower(), 0)

                self.after(0, _apply, name, vk_code)
            except Exception as e:
                self.after(0, lambda: self._log(f"Capture error: {e}", "error"))
                self.after(0, _restore, orig_text)

        def _apply(name, vk_code):
            key_var.set(name)
            vk_var.set(vk_code)
            vk_lbl_var.set(f"0x{vk_code:02X}" if vk_code else "—")
            entry.config(fg=FG)
            self._log(f"Captured: '{name}'  VK=0x{vk_code:02X}", "info")
            self._capturing = False

        def _restore(fallback="⊙ key"):
            key_var.set(fallback if fallback in ("⊙ key", "") else fallback)
            entry.config(fg=YELLOW if fallback == "⊙ key" else FG)
            self._capturing = False

        threading.Thread(target=_thread, daemon=True).start()

    # ======================================================================
    # Window picker
    # ======================================================================
    def _populate_window_list(self):
        windows = get_open_windows()
        self._window_map = {}
        labels = []
        for hwnd, title, pid, proc in windows:
            label = f"[{proc}]  {title[:60]}"
            self._window_map[label] = hwnd
            labels.append(label)
        self.window_combo["values"] = labels
        self._log(f"Window list refreshed — {len(labels)} windows found.", "info")

    def _on_window_selected(self):
        label = self.window_combo.get()
        hwnd  = self._window_map.get(label, 0)
        self._target_hwnd = hwnd
        self._target_label.config(text=f"HWND: {hwnd}   {label}")
        self._log(f"Target: {label}  (hwnd={hwnd})", "info")

    # ======================================================================
    # Sequence start / stop
    # ======================================================================
    def start_pressing(self):
        if not self._target_hwnd:
            messagebox.showwarning("No Target", "Please select a target window first.")
            return
        if self.engine and self.engine.running:
            return

        sequence = self._collect_sequence()
        if not sequence:
            messagebox.showwarning("Empty Sequence", "Add at least one key to the sequence.")
            return

        self.engine = KeyPressEngine(
            target_hwnd=self._target_hwnd,
            sequence=sequence,
            log_cb=lambda msg: self._log(msg, _classify(msg))
        )
        self.engine.start()
        self._btn_start.config(state="disabled")
        self._btn_stop.config(state="normal")
        self._status_var.set("● Running")
        self._status_lbl.config(fg=GREEN)
        self._log("Sequence STARTED.", "info")

    def stop_pressing(self):
        if self.engine:
            self.engine.stop()
            self.engine = None
        self._btn_start.config(state="normal")
        self._btn_stop.config(state="disabled")
        self._status_var.set("● Idle")
        self._status_lbl.config(fg=FG2)
        self._log("Sequence STOPPED.", "info")

    # ======================================================================
    # Hotkeys
    # ======================================================================
    def _register_hotkeys(self):
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        start_key   = self._hk_start_var.get().strip()
        stop_key    = self._hk_stop_var.get().strip()
        alchemy_key = self._alchemy_hk_var.get().strip()
        try:
            keyboard.add_hotkey(start_key,   self.start_pressing,  suppress=False)
            keyboard.add_hotkey(stop_key,    self.stop_pressing,   suppress=False)
            if alchemy_key:
                keyboard.add_hotkey(alchemy_key, self.toggle_alchemy, suppress=False)
            self._log(
                f"Hotkeys: Start={start_key}  Stop={stop_key}  Alchemy={alchemy_key or '—'}",
                "info")
        except Exception as e:
            self._log(f"Hotkey registration failed: {e}", "error")

    # ======================================================================
    # Config persistence
    # ======================================================================
    def _collect_sequence(self) -> list:
        seq = []
        for key_var, vk_var, delay_var, _ in self._seq_rows:
            k = key_var.get().strip()
            if not k:
                continue
            try:
                d = int(delay_var.get())
            except (ValueError, tk.TclError):
                d = 50
            vk = vk_var.get() or KEY_NAME_TO_VK.get(k.lower(), 0)
            seq.append({"key": k, "vk": vk, "delay_ms": max(1, d)})
        return seq

    def _collect_timed_actions(self) -> list:
        actions = []
        for row in self._timed_rows:
            k = row["key_var"].get().strip()
            if not k:
                continue
            vk = row["vk_var"].get() or KEY_NAME_TO_VK.get(k.lower(), 0)
            enter_ms     = row["enter_ms_var"].get() if row["enter_var"].get() else 0
            unit         = row["initial_unit_var"].get()
            raw_initial  = max(0, row["initial_var"].get())
            initial_min  = raw_initial / 60.0 if unit == "sec" else float(raw_initial)
            actions.append({
                "key":            k,
                "vk":             vk,
                "hold_ms":        max(50, row["hold_var"].get()),
                "interval_min":   max(1, row["interval_var"].get()),
                "initial_min":    initial_min,
                "initial_unit":   unit,
                "enter_after_ms": enter_ms,
            })
        return actions

    def save_config(self):
        cfg = {
            "sequence":          self._collect_sequence(),
            "timed_actions":     self._collect_timed_actions(),
            "hotkey_start":      self._hk_start_var.get().strip(),
            "hotkey_stop":       self._hk_stop_var.get().strip(),
            "target_window":     self.window_combo.get(),
            "prefocus_sec":      max(1, self._prefocus_var.get()),
            "alchemy_delay_ms":   max(100, self._alchemy_delay_var.get()),
            "alchemy_hotkey":     self._alchemy_hk_var.get().strip(),
            "alchemy_text_region": self.config_data.get("alchemy_text_region"),
            "alchemy_stop_text":  self._alchemy_stop_text_var.get().strip(),
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
            self.config_data = cfg
            self._log(f"Settings saved → {CONFIG_FILE}", "info")
        except Exception as e:
            self._log(f"Failed to save config: {e}", "error")

    @staticmethod
    def _load_config() -> dict:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                merged = dict(DEFAULT_CONFIG)
                merged.update(data)
                return merged
            except Exception:
                pass
        return dict(DEFAULT_CONFIG)

    # ======================================================================
    # Log helpers
    # ======================================================================
    def _log(self, message: str, level: str = "info"):
        self.after(0, self._log_main_thread, message, level)

    def _log_main_thread(self, message: str, level: str):
        ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        self._log_text.config(state="normal")
        self._log_text.insert("end", f"[{ts}] ", "ts")
        self._log_text.insert("end", message + "\n", level)
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    # ======================================================================
    # Auto-updater
    # ======================================================================
    def _check_for_updates(self, btn):
        btn.config(state="disabled", text="Checking…")

        def _thread():
            try:
                url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
                req = urllib.request.Request(
                    url, headers={"User-Agent": "InvictusSROKP"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.load(r)

                tag = data["tag_name"].lstrip("v")
                try:
                    latest = int(tag)
                except ValueError:
                    latest = 0

                download_url = None
                for asset in data.get("assets", []):
                    if asset["name"].lower().endswith(".exe"):
                        download_url = asset["browser_download_url"]
                        break

                self.after(0, _done, latest, download_url,
                           data.get("name", f"v{latest}"))
            except Exception as e:
                self.after(0, _error, str(e))

        def _done(latest, download_url, release_name):
            btn.config(state="normal", text="🔄 Check for Update")
            if latest <= APP_VERSION:
                messagebox.showinfo(
                    "Up to date",
                    f"You're already on the latest version (v{APP_VERSION}). ✓")
            elif not download_url:
                messagebox.showwarning(
                    "Update available",
                    f"{release_name} is available but no .exe was found in the release.\n"
                    f"Download manually from:\nhttps://github.com/{GITHUB_REPO}/releases")
            else:
                if messagebox.askyesno(
                    "Update available",
                    f"v{latest} is available  (you have v{APP_VERSION}).\n\n"
                    f"Download and install now?\nThe app will restart automatically."):
                    self._do_update(download_url)

        def _error(msg):
            btn.config(state="normal", text="🔄 Check for Update")
            messagebox.showerror(
                "Update check failed",
                f"Could not reach GitHub:\n{msg}")

        threading.Thread(target=_thread, daemon=True).start()

    def _do_update(self, download_url: str):
        """Download new exe to a temp file, write a bat that swaps it in, then exit."""
        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                "Running from source",
                "Auto-update only works with the .exe version.\n"
                "Pull the latest code from GitHub manually.")
            return

        # Progress window
        win = tk.Toplevel(self)
        win.title("Downloading update…")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="Downloading update, please wait…",
                 fg=FG, bg=BG, font=("Segoe UI", 10)).pack(padx=24, pady=(16, 6))
        prog_var = tk.StringVar(value="0 %")
        tk.Label(win, textvariable=prog_var, fg=TEAL, bg=BG,
                 font=("Consolas", 12, "bold")).pack(pady=(0, 20))

        tmp_exe = os.path.join(tempfile.gettempdir(), "InvictusSROKP_update.exe")

        def _reporthook(blocks, block_size, total):
            if total > 0:
                pct = min(100, blocks * block_size * 100 // total)
                self.after(0, prog_var.set, f"{pct} %")

        def _thread():
            try:
                urllib.request.urlretrieve(download_url, tmp_exe, _reporthook)
                self.after(0, _apply)
            except Exception as e:
                self.after(0, win.destroy)
                self.after(0, lambda: messagebox.showerror(
                    "Download failed", str(e)))

        def _apply():
            win.destroy()
            exe = sys.executable
            # Write a tiny bat: wait 2 s (so this process exits), swap files, relaunch
            bat = os.path.join(tempfile.gettempdir(), "invictus_updater.bat")
            with open(bat, "w") as f:
                f.write(
                    "@echo off\n"
                    "timeout /t 2 /nobreak >nul\n"
                    f'copy /y "{tmp_exe}" "{exe}"\n'
                    f'start "" "{exe}"\n'
                    f'del "{tmp_exe}"\n'
                    'del "%~f0"\n'
                )
            subprocess.Popen(
                ["cmd", "/c", bat],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            self._on_close()

        threading.Thread(target=_thread, daemon=True).start()

    # ======================================================================
    # Window close
    # ======================================================================
    def _on_close(self):
        self.stop_pressing()
        for row in list(self._timed_rows):
            self._stop_timed(row)
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        self.destroy()


# ==============================================================================
# Utility: classify a log message for colour tagging
# ==============================================================================
def _classify(msg: str) -> str:
    m = msg.lower()
    if "error" in m or "fail" in m or "skipped" in m:
        return "error"
    if "lost" in m or "warn" in m or "paused" in m:
        return "warn"
    if "[timer]" in m:
        return "timer"
    if "sent" in m or "captured" in m:
        return "key"
    return "info"


# ==============================================================================
# Entry point
# ==============================================================================
if __name__ == "__main__":
    app = KeyPresserApp()

    style = ttk.Style(app)
    style.theme_use("clam")
    style.configure("TCombobox",
                    fieldbackground=BG3, background=BG3,
                    foreground=FG, selectforeground=FG,
                    selectbackground=BORDER, bordercolor=BORDER,
                    arrowcolor=ACCENT)
    style.map("TCombobox",
              fieldbackground=[("readonly", BG3)],
              foreground     =[("readonly", FG)],
              selectbackground=[("readonly", BORDER)])

    app.mainloop()
