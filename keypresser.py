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
import win32gui
import win32process
import win32con
import win32api
import psutil
import keyboard


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
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "keypresser_config.json")

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
                 enter_after_ms: int = 0):
        self.target_hwnd   = target_hwnd
        self.vk_code       = vk_code
        self.hold_ms       = hold_ms
        self.interval_sec  = interval_min * 60.0
        self.label         = label
        self.log_cb        = log_cb
        self.enter_after_ms = enter_after_ms   # 0 = disabled
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
        # How many seconds before the trigger we start forcing the window to foreground.
        PRE_FOCUS_SEC = 2.0

        self.log_cb(
            f"[Timer] '{self.label}' started — first press in "
            f"{self._initial_sec/60:.1f} min, then every {self.interval_sec/60:.1f} min."
        )

        while self.running:
            remaining = self.next_trigger - time.time()

            if remaining > PRE_FOCUS_SEC:
                # Still far from trigger — sleep in short bursts so stop() is responsive
                time.sleep(0.5)
                continue

            if remaining > 0:
                # ── Pre-focus window: force the game to the foreground ────
                # We do this every 100 ms for the last 2 seconds so the OS
                # has plenty of time to complete the focus switch before the press.
                try:
                    win32gui.SetForegroundWindow(self.target_hwnd)
                    self.log_cb(
                        f"[Timer] '{self.label}' — forcing focus "
                        f"({remaining:.1f}s to press)…"
                    )
                except Exception:
                    pass
                time.sleep(0.1)
                continue

            # ── Timer fired — press the key ───────────────────────────────
            if not self.running:
                break

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

        self._target_hwnd  = 0
        self._window_map   = {}
        self._seq_rows     = []   # (key_var, vk_var, delay_var, frame)
        self._timed_rows   = []   # list of row-dicts (see _add_timed_row)
        self._capturing    = False

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
        _ft = tk.Frame(footer, bg=BG2)
        _ft.pack(side="right", padx=10, pady=4)
        tk.Label(_ft, text="v2  —  Made by LinKx with ",
                 font=("Segoe UI", 9), fg=FG2, bg=BG2).pack(side="left")
        tk.Label(_ft, text="\u2665",
                 font=("Segoe UI", 9, "bold"), fg=RED, bg=BG2).pack(side="left")

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
    _TA_COL_W = [28, 60, 68, 70, 70, 90, 28, 52, 28, 28, 28]
    #             #  Key Hold Every Next  Cdown ↵?  ↵ms  ▶   ■   ✕

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

        headers = ["#", "Key", "Hold\n(ms)", "Every\n(min)", "Next in\n(min)",
                   "Countdown", "↵?", "↵ ms", "▶", "■", "✕"]
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
        tk.Label(ctrl,
                 text='  "Next in" = minutes until first press  '
                      '(e.g. buff has 48 min left -> set 48, repeat every 60)',
                 fg=FG2, bg=BG, font=("Segoe UI", 7, "italic")).pack(side="left", padx=6)
        # Live/idle indicator — updated every 500 ms by _tick_countdowns()
        self._timed_status_lbl = tk.Label(ctrl, text="● Idle",
                                          fg=RED, bg=BG, font=("Segoe UI", 8, "bold"))
        self._timed_status_lbl.pack(side="right", padx=4)

    def _add_timed_row(self, key: str = "", vk: int = 0,
                       hold_ms: int = 200, interval_min: int = 60,
                       initial_min: int = 60, enter_after_ms: int = 0):
        # grid row = position + 1 (row 0 is the header)
        grid_row = len(self._timed_rows) + 1

        # ── Index label ───────────────────────────────────────────────────
        idx_lbl = tk.Label(self._timed_grid, text=str(grid_row), anchor="center",
                           fg=FG2, bg=BG, font=("Consolas", 9))
        idx_lbl.grid(row=grid_row, column=0, padx=1, pady=2, sticky="ew")

        # ── Key entry ─────────────────────────────────────────────────────
        # width=6 caps the column at ~42 px (Consolas 9) so the grid fits
        key_var = tk.StringVar(value=key)
        key_entry = tk.Entry(self._timed_grid, textvariable=key_var,
                             width=6,
                             bg=BG3, fg=FG, insertbackground=FG,
                             relief="flat", font=("Consolas", 9), justify="center")
        key_entry.grid(row=grid_row, column=1, padx=1, pady=2, sticky="ew")

        # ── Hold spinbox ──────────────────────────────────────────────────
        hold_var = tk.IntVar(value=hold_ms)
        hold_spin = tk.Spinbox(self._timed_grid, from_=50, to=9999,
                               textvariable=hold_var, width=5,
                               bg=BG3, fg=FG, insertbackground=FG,
                               buttonbackground=BG3, relief="flat",
                               font=("Consolas", 9))
        hold_spin.grid(row=grid_row, column=2, padx=1, pady=2, sticky="ew")

        # ── Interval spinbox ──────────────────────────────────────────────
        interval_var = tk.IntVar(value=interval_min)
        interval_spin = tk.Spinbox(self._timed_grid, from_=1, to=9999,
                                   textvariable=interval_var, width=5,
                                   bg=BG3, fg=FG, insertbackground=FG,
                                   buttonbackground=BG3, relief="flat",
                                   font=("Consolas", 9))
        interval_spin.grid(row=grid_row, column=3, padx=1, pady=2, sticky="ew")

        # ── "Next in" spinbox ─────────────────────────────────────────────
        # Highlighted yellow — the user sets this once to match remaining buff
        # time before starting. After the first press, only interval_var is used.
        initial_var = tk.IntVar(value=initial_min)
        initial_spin = tk.Spinbox(self._timed_grid, from_=0, to=9999,
                                  textvariable=initial_var, width=5,
                                  bg=BG3, fg=YELLOW, insertbackground=FG,
                                  buttonbackground=BG3, relief="flat",
                                  font=("Consolas", 9))
        initial_spin.grid(row=grid_row, column=4, padx=1, pady=2, sticky="ew")

        # ── Live countdown label ───────────────────────────────────────────
        # Ticked every 500 ms by _tick_countdowns(); colour shifts
        # TEAL → YELLOW → RED as the next press approaches.
        countdown_var = tk.StringVar(value="——:——")
        cd_lbl = tk.Label(self._timed_grid, textvariable=countdown_var,
                          anchor="center", fg=TEAL, bg=BG,
                          font=("Consolas", 9, "bold"))
        cd_lbl.grid(row=grid_row, column=5, padx=1, pady=2, sticky="ew")

        # ── Enter-after checkbox (col 6) ──────────────────────────────────
        # When checked, the engine sends Enter N ms after the main key press.
        enter_var = tk.BooleanVar(value=bool(enter_after_ms))
        enter_chk = tk.Checkbutton(self._timed_grid, variable=enter_var,
                                   bg=BG, activebackground=BG,
                                   selectcolor=BG3, relief="flat",
                                   cursor="hand2")
        enter_chk.grid(row=grid_row, column=6, padx=1, pady=2)

        # ── Enter-delay spinbox (col 7) ───────────────────────────────────
        # Active value in ms; only used when enter_var is True.
        enter_ms_val = enter_after_ms if enter_after_ms > 0 else 1500
        enter_ms_var = tk.IntVar(value=enter_ms_val)
        enter_ms_spin = tk.Spinbox(self._timed_grid, from_=100, to=9999,
                                   textvariable=enter_ms_var, width=4,
                                   bg=BG3, fg=TEAL, insertbackground=FG,
                                   buttonbackground=BG3, relief="flat",
                                   font=("Consolas", 9))
        enter_ms_spin.grid(row=grid_row, column=7, padx=1, pady=2, sticky="ew")

        # ── VK code (internal only, not displayed) ───────────────────────
        vk_var     = tk.IntVar(value=vk)
        vk_lbl_var = tk.StringVar(value=f"0x{vk:02X}" if vk else "—")

        # ── Start button (col 8) ──────────────────────────────────────────
        start_btn = tk.Button(self._timed_grid, text="▶",
                              bg=BG3, fg=GREEN, activebackground=BORDER,
                              activeforeground=GREEN, relief="flat",
                              font=("Segoe UI", 9, "bold"), cursor="hand2")
        start_btn.grid(row=grid_row, column=8, padx=1, pady=2, sticky="ew")

        # ── Stop button (col 9) ───────────────────────────────────────────
        stop_btn = tk.Button(self._timed_grid, text="■",
                             bg=BG3, fg=RED, activebackground=BORDER,
                             activeforeground=RED, relief="flat",
                             font=("Segoe UI", 9, "bold"), cursor="hand2",
                             state="disabled")
        stop_btn.grid(row=grid_row, column=9, padx=1, pady=2, sticky="ew")

        # ── Delete button (col 10) ────────────────────────────────────────
        del_btn = tk.Button(self._timed_grid, text="✕",
                            bg=BG, fg=RED, activebackground=BORDER,
                            activeforeground=RED, relief="flat",
                            font=("Segoe UI", 9), cursor="hand2")
        del_btn.grid(row=grid_row, column=10, padx=1, pady=2, sticky="ew")

        # ── Row data dict ──────────────────────────────────────────────────
        row = {
            "key_var":        key_var,
            "vk_var":         vk_var,
            "vk_lbl_var":     vk_lbl_var,
            "hold_var":       hold_var,
            "interval_var":   interval_var,
            "initial_var":    initial_var,
            "countdown_var":  countdown_var,
            "cd_lbl":         cd_lbl,        # direct ref for colour changes
            "enter_var":      enter_var,      # BooleanVar — send Enter?
            "enter_ms_var":   enter_ms_var,   # IntVar — delay before Enter
            "start_btn":      start_btn,
            "stop_btn":       stop_btn,
            "engine":         None,
            "idx_lbl":        idx_lbl,
            # All grid widgets for this row — used during deletion
            "_grid_widgets": [idx_lbl, key_entry, hold_spin, interval_spin,
                              initial_spin, cd_lbl, enter_chk, enter_ms_spin,
                              start_btn, stop_btn, del_btn],
        }
        self._timed_rows.append(row)

        # Wire up buttons
        start_btn.config(command=lambda r=row: self._start_timed(r))
        stop_btn.config(command=lambda r=row: self._stop_timed(r))

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

        vk = row["vk_var"].get()
        if not vk:
            vk = KEY_NAME_TO_VK.get(row["key_var"].get().strip().lower(), 0)
        if not vk:
            messagebox.showwarning(
                "No Key Captured",
                "Use the ⊙ Capture button to assign a key to this timed action first."
            )
            return

        label          = row["key_var"].get() or f"VK=0x{vk:02X}"
        hold_ms        = max(50, row["hold_var"].get())
        interval_min   = max(1, row["interval_var"].get())
        initial_min    = max(0, row["initial_var"].get())
        enter_after_ms = row["enter_ms_var"].get() if row["enter_var"].get() else 0

        engine = TimedActionEngine(
            target_hwnd    = self._target_hwnd,
            vk_code        = vk,
            hold_ms        = hold_ms,
            interval_min   = interval_min,
            initial_min    = initial_min,
            label          = label,
            log_cb         = lambda msg: self._log(msg, "timer"),
            enter_after_ms = enter_after_ms,
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

        key_var = tk.StringVar(value=key)
        tk.Entry(frame, textvariable=key_var, width=7, bg=BG3, fg=FG,
                 insertbackground=FG, relief="flat", font=("Consolas", 10),
                 justify="center").pack(side="left", padx=4)

        vk_var     = tk.IntVar(value=vk)
        vk_lbl_var = tk.StringVar(value=f"0x{vk:02X}" if vk else "—")
        tk.Label(frame, textvariable=vk_lbl_var, width=7, anchor="center",
                 fg=ACCENT2, bg=BG2, font=("Consolas", 9)).pack(side="left", padx=4)

        delay_var = tk.IntVar(value=delay_ms)
        tk.Spinbox(frame, from_=1, to=9999, textvariable=delay_var, width=7,
                   bg=BG3, fg=FG, insertbackground=FG, buttonbackground=BG3,
                   relief="flat", font=("Consolas", 10)).pack(side="left", padx=4)

        capture_btn = tk.Button(frame, text="⊙ Capture",
                                bg=BG3, fg=YELLOW, activebackground=BORDER,
                                activeforeground=FG, relief="flat",
                                font=("Segoe UI", 8), cursor="hand2", padx=4)
        capture_btn.pack(side="left", padx=2)

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
        start_key = self._hk_start_var.get().strip()
        stop_key  = self._hk_stop_var.get().strip()
        try:
            keyboard.add_hotkey(start_key, self.start_pressing, suppress=False)
            keyboard.add_hotkey(stop_key,  self.stop_pressing,  suppress=False)
            self._log(f"Hotkeys: Start={start_key}  Stop={stop_key}", "info")
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
            enter_ms = row["enter_ms_var"].get() if row["enter_var"].get() else 0
            actions.append({
                "key":            k,
                "vk":             vk,
                "hold_ms":        max(50, row["hold_var"].get()),
                "interval_min":   max(1, row["interval_var"].get()),
                "initial_min":    max(0, row["initial_var"].get()),
                "enter_after_ms": enter_ms,
            })
        return actions

    def save_config(self):
        cfg = {
            "sequence":       self._collect_sequence(),
            "timed_actions":  self._collect_timed_actions(),
            "hotkey_start":   self._hk_start_var.get().strip(),
            "hotkey_stop":    self._hk_stop_var.get().strip(),
            "target_window":  self.window_combo.get(),
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
