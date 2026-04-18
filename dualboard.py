#!/usr/bin/env python3
"""
DualBoard - Per-device keyboard macro manager for Windows.

Lets you designate a secondary keyboard (e.g. a Keychron Q0 numpad) as a
macropad, and remap its keys to actions WITHOUT affecting your main keyboard.

How it works:
  1. Windows' Raw Input API is used to identify WHICH physical keyboard
     sent a keypress.
  2. A low-level keyboard hook (WH_KEYBOARD_LL) intercepts every key before
     it reaches any application.
  3. When a key comes from the designated "macropad" device, it's blocked
     and the mapped action is executed instead. Keys from all other
     keyboards pass through untouched.

Requirements: Windows 10/11, Python 3.8+. Uses only stdlib (tkinter, ctypes).

Usage: python dualboard.py

    - Click "Refresh" to enumerate keyboards.
    - Pick your macropad from the dropdown.
    - Click "+ Add Mapping" to record a key and assign it an action.
    - Click "Start Monitoring" to activate.
"""

import ctypes
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import webbrowser
from ctypes import wintypes, POINTER, WINFUNCTYPE, byref, sizeof, Structure
from pathlib import Path
from tkinter import ttk, messagebox, filedialog

# =============================================================================
# Win32 constants
# =============================================================================

# Raw input
RIM_TYPEKEYBOARD = 1
RIDEV_INPUTSINK = 0x00000100
RIDEV_REMOVE = 0x00000001
RID_INPUT = 0x10000003
RIDI_DEVICENAME = 0x20000007

# Window message
WM_INPUT = 0x00FF
WM_QUIT = 0x0012
HWND_MESSAGE = -3

# Keyboard flags in RAWKEYBOARD
RI_KEY_MAKE = 0
RI_KEY_BREAK = 1
RI_KEY_E0 = 2
RI_KEY_E1 = 4

# Low-level hook
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105
HC_ACTION = 0

LLKHF_EXTENDED = 0x01
LLKHF_INJECTED = 0x10

# SendInput
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_SCANCODE = 0x0008
KEYEVENTF_EXTENDEDKEY = 0x0001

# =============================================================================
# Win32 structures
# =============================================================================


class RAWINPUTDEVICE(Structure):
    _fields_ = [
        ("usUsagePage", wintypes.USHORT),
        ("usUsage", wintypes.USHORT),
        ("dwFlags", wintypes.DWORD),
        ("hwndTarget", wintypes.HWND),
    ]


class RAWINPUTDEVICELIST(Structure):
    _fields_ = [
        ("hDevice", wintypes.HANDLE),
        ("dwType", wintypes.DWORD),
    ]


class RAWINPUTHEADER(Structure):
    _fields_ = [
        ("dwType", wintypes.DWORD),
        ("dwSize", wintypes.DWORD),
        ("hDevice", wintypes.HANDLE),
        ("wParam", wintypes.WPARAM),
    ]


class RAWKEYBOARD(Structure):
    _fields_ = [
        ("MakeCode", wintypes.USHORT),
        ("Flags", wintypes.USHORT),
        ("Reserved", wintypes.USHORT),
        ("VKey", wintypes.USHORT),
        ("Message", wintypes.UINT),
        ("ExtraInformation", wintypes.ULONG),
    ]


class RAWINPUT(Structure):
    _fields_ = [
        ("header", RAWINPUTHEADER),
        ("keyboard", RAWKEYBOARD),
    ]


class KBDLLHOOKSTRUCT(Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class WNDCLASS(Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class MSG(Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", wintypes.POINT),
    ]


class KEYBDINPUT(Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class _INPUTunion(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("padding", ctypes.c_byte * 32)]


class INPUT(Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", _INPUTunion)]


# =============================================================================
# Win32 API bindings
# =============================================================================

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

WNDPROC = WINFUNCTYPE(
    ctypes.c_long, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
)
HOOKPROC = WINFUNCTYPE(ctypes.c_long, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

user32.RegisterClassW.argtypes = [POINTER(WNDCLASS)]
user32.RegisterClassW.restype = wintypes.ATOM
user32.DefWindowProcW.argtypes = [
    wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
]
user32.DefWindowProcW.restype = ctypes.c_long
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID,
]
user32.CreateWindowExW.restype = wintypes.HWND
user32.DestroyWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [POINTER(MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.TranslateMessage.argtypes = [POINTER(MSG)]
user32.DispatchMessageW.argtypes = [POINTER(MSG)]
user32.PostThreadMessageW.argtypes = [
    wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
]
user32.PostThreadMessageW.restype = wintypes.BOOL
user32.PeekMessageW.argtypes = [
    POINTER(MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT
]
user32.PeekMessageW.restype = wintypes.BOOL
PM_REMOVE = 0x0001

user32.RegisterRawInputDevices.argtypes = [
    POINTER(RAWINPUTDEVICE), wintypes.UINT, wintypes.UINT
]
user32.RegisterRawInputDevices.restype = wintypes.BOOL
user32.GetRawInputData.argtypes = [
    wintypes.HANDLE, wintypes.UINT, wintypes.LPVOID,
    POINTER(wintypes.UINT), wintypes.UINT,
]
user32.GetRawInputData.restype = wintypes.UINT
user32.GetRawInputDeviceList.argtypes = [
    POINTER(RAWINPUTDEVICELIST), POINTER(wintypes.UINT), wintypes.UINT
]
user32.GetRawInputDeviceList.restype = wintypes.UINT
user32.GetRawInputDeviceInfoW.argtypes = [
    wintypes.HANDLE, wintypes.UINT, wintypes.LPVOID, POINTER(wintypes.UINT)
]
user32.GetRawInputDeviceInfoW.restype = wintypes.UINT

user32.SetWindowsHookExW.argtypes = [
    ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD
]
user32.SetWindowsHookExW.restype = wintypes.HHOOK
user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.CallNextHookEx.argtypes = [
    wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
]
user32.CallNextHookEx.restype = ctypes.c_long

user32.SendInput.argtypes = [wintypes.UINT, POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT

kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE
kernel32.GetCurrentThreadId.restype = wintypes.DWORD

# =============================================================================
# Virtual key code -> friendly name map (partial; the common ones)
# =============================================================================

VK_NAMES = {
    0x08: "Backspace", 0x09: "Tab", 0x0D: "Enter", 0x10: "Shift",
    0x11: "Ctrl", 0x12: "Alt", 0x13: "Pause", 0x14: "CapsLock",
    0x1B: "Esc", 0x20: "Space", 0x21: "PageUp", 0x22: "PageDown",
    0x23: "End", 0x24: "Home", 0x25: "Left", 0x26: "Up",
    0x27: "Right", 0x28: "Down", 0x2C: "PrintScreen", 0x2D: "Insert",
    0x2E: "Delete", 0x5B: "LWin", 0x5C: "RWin", 0x5D: "Menu",
    0x60: "Num0", 0x61: "Num1", 0x62: "Num2", 0x63: "Num3",
    0x64: "Num4", 0x65: "Num5", 0x66: "Num6", 0x67: "Num7",
    0x68: "Num8", 0x69: "Num9", 0x6A: "Num*", 0x6B: "Num+",
    0x6C: "NumEnter", 0x6D: "Num-", 0x6E: "Num.", 0x6F: "Num/",
    0x70: "F1", 0x71: "F2", 0x72: "F3", 0x73: "F4", 0x74: "F5",
    0x75: "F6", 0x76: "F7", 0x77: "F8", 0x78: "F9", 0x79: "F10",
    0x7A: "F11", 0x7B: "F12", 0x7C: "F13", 0x7D: "F14", 0x7E: "F15",
    0x7F: "F16", 0x80: "F17", 0x81: "F18", 0x82: "F19", 0x83: "F20",
    0x84: "F21", 0x85: "F22", 0x86: "F23", 0x87: "F24",
    0x90: "NumLock", 0x91: "ScrollLock",
    0xA0: "LShift", 0xA1: "RShift", 0xA2: "LCtrl", 0xA3: "RCtrl",
    0xA4: "LAlt", 0xA5: "RAlt",
    0xBA: ";", 0xBB: "=", 0xBC: ",", 0xBD: "-", 0xBE: ".", 0xBF: "/",
    0xC0: "`", 0xDB: "[", 0xDC: "\\", 0xDD: "]", 0xDE: "'",
}


def vk_name(vk: int) -> str:
    """Return a friendly name for a virtual-key code."""
    if vk in VK_NAMES:
        return VK_NAMES[vk]
    if 0x30 <= vk <= 0x39:  # digits
        return chr(vk)
    if 0x41 <= vk <= 0x5A:  # letters
        return chr(vk)
    return f"VK_{vk:02X}"


# =============================================================================
# Device enumeration
# =============================================================================


# Generic device descriptions we want to see through to find a real name.
_GENERIC_NAMES = {
    "hid keyboard device",
    "hid-compliant keyboard",
    "hid keyboard",
    "keyboard device",
    "usb composite device",
    "usb input device",
    "hid-compliant device",
    "hid-compliant consumer control device",
    "hid-compliant system controller",
    "hid-compliant vendor-defined device",
}


def _parse_vid_pid(raw_device_name: str):
    """Extract VID/PID from a raw device path, or (None, None)."""
    vm = re.search(r"VID_([0-9A-Fa-f]{4})", raw_device_name)
    pm = re.search(r"PID_([0-9A-Fa-f]{4})", raw_device_name)
    return (
        vm.group(1).upper() if vm else None,
        pm.group(1).upper() if pm else None,
    )


def _physical_device_key(raw_device_name: str):
    """
    A stable key that groups all HID collections belonging to the same
    physical keyboard. We use VID+PID plus the USB instance/serial portion
    of the device path (the middle `#` segment, e.g. "7&abc123&0&0000"),
    which is shared across a device's sibling interfaces.
    """
    vid, pid = _parse_vid_pid(raw_device_name)
    # Device path looks like: \\?\HID#VID_xxxx&PID_xxxx&MI_xx&Col0x#<instance>#{GUID}
    parts = raw_device_name.split("#")
    instance = parts[2] if len(parts) >= 3 else ""
    # Drop the trailing "&0&0000"-style collection suffix to merge siblings.
    instance_root = instance.rsplit("&", 2)[0] if "&" in instance else instance
    return (vid, pid, instance_root)


def enumerate_keyboards():
    """
    Return a list of (handles, friendly_name, raw_name) for all keyboards,
    deduplicated by physical device. `handles` is a tuple of all HID handles
    belonging to the same physical keyboard (a single keyboard often exposes
    several HID collections — keyboard, consumer, system control).
    """
    count = wintypes.UINT(0)
    user32.GetRawInputDeviceList(None, byref(count), sizeof(RAWINPUTDEVICELIST))
    if count.value == 0:
        return []

    devices = (RAWINPUTDEVICELIST * count.value)()
    got = user32.GetRawInputDeviceList(devices, byref(count), sizeof(RAWINPUTDEVICELIST))
    if got == 0xFFFFFFFF:
        return []

    # group_key -> { handles: [...], raw_names: [...] }
    groups = {}
    group_order = []

    for i in range(got):
        dev = devices[i]
        if dev.dwType != RIM_TYPEKEYBOARD:
            continue

        # Get the raw device path
        name_len = wintypes.UINT(0)
        user32.GetRawInputDeviceInfoW(dev.hDevice, RIDI_DEVICENAME, None, byref(name_len))
        if name_len.value == 0:
            continue
        name_buf = ctypes.create_unicode_buffer(name_len.value)
        user32.GetRawInputDeviceInfoW(
            dev.hDevice, RIDI_DEVICENAME, name_buf, byref(name_len)
        )
        raw_name = name_buf.value

        # Skip RDP / synthetic keyboards (no VID/PID)
        if "VID_" not in raw_name:
            continue

        key = _physical_device_key(raw_name)
        if key not in groups:
            groups[key] = {"handles": [], "raw_names": []}
            group_order.append(key)
        groups[key]["handles"].append(dev.hDevice)
        groups[key]["raw_names"].append(raw_name)

    results = []
    for key in group_order:
        g = groups[key]
        # Use the first raw_name of the group as the canonical identifier;
        # try to resolve a friendly name using any of the group's paths.
        friendly = None
        for rn in g["raw_names"]:
            friendly = _friendly_name(rn)
            if friendly and friendly.lower() not in _GENERIC_NAMES:
                break
        friendly = friendly or "Keyboard"
        results.append((tuple(g["handles"]), friendly, g["raw_names"][0]))

    return results


def _friendly_name(raw_device_name: str) -> str:
    """
    Produce a human-readable name for a raw HID device path. Tries, in order:
      1. The USB product string descriptor (HidD_GetProductString) — the
         actual marketing name the device advertises (e.g. "Keychron Q0").
      2. The HID node's own registry DeviceDesc/FriendlyName.
      3. The parent USB device's registry DeviceDesc/FriendlyName.
      4. A VID:PID fallback.
    """
    vid, pid = _parse_vid_pid(raw_device_name)

    candidates = []

    # 1) Ask the device directly.
    s = _hid_product_string(raw_device_name)
    if s:
        candidates.append(s)

    # 2) The HID node's own registry entry (often generic for keyboards).
    s = _registry_name_for_path(raw_device_name)
    if s:
        candidates.append(s)

    # 3) Walk up to the USB parent device, which usually has a real name.
    if vid and pid:
        s = _usb_parent_name(vid, pid)
        if s:
            candidates.append(s)

    # Pick the first specific (non-generic) candidate.
    for c in candidates:
        if c and c.strip().lower() not in _GENERIC_NAMES:
            return c.strip()

    # Otherwise, take the first non-empty candidate and append VID:PID for
    # disambiguation between multiple generic keyboards.
    base = next((c for c in candidates if c), "Keyboard")
    if vid and pid:
        return f"{base} (VID:{vid} PID:{pid})"
    return base


def _registry_name_for_path(raw_device_name: str):
    """Read DeviceDesc/FriendlyName from the HID device's registry key."""
    try:
        import winreg
    except ImportError:
        return None
    try:
        stripped = raw_device_name
        if stripped.startswith("\\\\?\\"):
            stripped = stripped[4:]
        if "#{" in stripped:
            stripped = stripped.split("#{")[0]
        reg_path = "SYSTEM\\CurrentControlSet\\Enum\\" + stripped.replace("#", "\\")
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            for value_name in ("FriendlyName", "DeviceDesc"):
                try:
                    val, _ = winreg.QueryValueEx(key, value_name)
                    # DeviceDesc often looks like "@oem.inf,%key%;Display Name"
                    if ";" in val:
                        val = val.split(";", 1)[-1]
                    return val
                except FileNotFoundError:
                    continue
    except OSError:
        pass
    return None


def _usb_parent_name(vid: str, pid: str):
    """
    Walk SYSTEM\\CurrentControlSet\\Enum\\USB\\VID_xxxx&PID_xxxx\\* and
    return the first specific DeviceDesc/FriendlyName we find.
    """
    try:
        import winreg
    except ImportError:
        return None
    usb_path = f"SYSTEM\\CurrentControlSet\\Enum\\USB\\VID_{vid}&PID_{pid}"
    first_generic = None
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, usb_path) as key:
            for i in range(256):
                try:
                    sub_name = winreg.EnumKey(key, i)
                except OSError:
                    break
                try:
                    with winreg.OpenKey(key, sub_name) as sub:
                        for value_name in ("FriendlyName", "DeviceDesc"):
                            try:
                                val, _ = winreg.QueryValueEx(sub, value_name)
                                if ";" in val:
                                    val = val.split(";", 1)[-1]
                                val = val.strip()
                                if not val:
                                    continue
                                if val.lower() not in _GENERIC_NAMES:
                                    return val
                                if first_generic is None:
                                    first_generic = val
                            except FileNotFoundError:
                                continue
                except OSError:
                    continue
    except OSError:
        pass
    return first_generic


def _hid_product_string(raw_device_name: str):
    """
    Open the HID device with zero access (metadata only — doesn't fight
    exclusive-access games/keyboards) and ask for its USB product string.
    This is the cleanest source of a human-readable name.
    """
    try:
        hid = ctypes.WinDLL("hid", use_last_error=True)
    except OSError:
        return None

    CreateFileW = kernel32.CreateFileW
    CreateFileW.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
        ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
    ]
    CreateFileW.restype = wintypes.HANDLE
    CloseHandle = kernel32.CloseHandle
    CloseHandle.argtypes = [wintypes.HANDLE]
    CloseHandle.restype = wintypes.BOOL

    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    OPEN_EXISTING = 3
    INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

    h = CreateFileW(
        raw_device_name, 0,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        None, OPEN_EXISTING, 0, None,
    )
    if not h or h == INVALID_HANDLE_VALUE:
        return None
    try:
        HidD_GetProductString = hid.HidD_GetProductString
        HidD_GetProductString.argtypes = [
            wintypes.HANDLE, ctypes.c_void_p, wintypes.ULONG
        ]
        HidD_GetProductString.restype = wintypes.BOOL

        HidD_GetManufacturerString = hid.HidD_GetManufacturerString
        HidD_GetManufacturerString.argtypes = [
            wintypes.HANDLE, ctypes.c_void_p, wintypes.ULONG
        ]
        HidD_GetManufacturerString.restype = wintypes.BOOL

        prod = ctypes.create_unicode_buffer(256)
        mfr = ctypes.create_unicode_buffer(256)
        got_prod = bool(HidD_GetProductString(h, prod, sizeof(prod)))
        got_mfr = bool(HidD_GetManufacturerString(h, mfr, sizeof(mfr)))

        p = prod.value.strip() if got_prod else ""
        m = mfr.value.strip() if got_mfr else ""

        if p and m and not p.lower().startswith(m.lower()):
            return f"{m} {p}"
        if p:
            return p
        if m:
            return m
        return None
    finally:
        CloseHandle(h)


# =============================================================================
# Keyboard monitor (runs in a dedicated worker thread)
# =============================================================================


class KeyboardMonitor:
    """
    Runs a Win32 message pump on a dedicated thread. Registers for Raw Input
    (to identify device) and installs a low-level keyboard hook (to intercept).

    When a key from the target device fires, it's blocked and a callback is
    invoked on the worker thread — the callback should queue work for the GUI.
    """

    def __init__(self, on_key_event, on_blocked_action):
        self.on_key_event = on_key_event          # called for EVERY key (for diagnostics)
        self.on_blocked_action = on_blocked_action  # called when a bound key fires

        # Set of HANDLEs belonging to the target physical keyboard. A single
        # keyboard often has several HID collections; any of them counts.
        self.target_handles = set()
        self.mappings = {}              # { vk_code: action_dict }
        self._recording = False         # True while the GUI is recording a key
        self._lock = threading.Lock()

        # --- Device-per-VK cache (accessed only from the worker thread) ---
        #
        # Blocking a key in the LL hook suppresses the corresponding WM_INPUT
        # message, so we can't "block first, identify device later". Instead,
        # we use WM_INPUT from *previous* events to learn which physical
        # device last sent each VK code, and query that cache inside the hook
        # to make the block-or-pass decision immediately.
        #
        # Before checking the cache the hook calls _drain_raw_input() which
        # uses PeekMessageW to process any WM_INPUT messages that are already
        # queued — this keeps the cache as fresh as possible.
        self._last_device_for_vk = {}   # vk (int) → hDevice (int)
        self._draining = False          # reentrancy guard for PeekMessageW
        # VK codes whose KEYDOWN we blocked — we must block the matching
        # KEYUP too, or the OS thinks the key is still held.
        self._blocked_downs = set()

        self._thread = None
        self._thread_id = None
        self._hwnd = None
        self._hook = None
        self._running = False

        # Keep strong refs so GC doesn't eat our callbacks
        self._wnd_proc = WNDPROC(self._wnd_proc_impl)
        self._hook_proc = HOOKPROC(self._hook_proc_impl)
        self._atom = None

    # ----- Public API (called from GUI thread) -----

    def set_target(self, handles):
        """Set the target device's HID handle(s). Accepts an iterable."""
        with self._lock:
            if handles is None:
                self.target_handles = set()
            elif isinstance(handles, (list, tuple, set, frozenset)):
                self.target_handles = set(handles)
            else:
                # Single handle, for backwards compatibility
                self.target_handles = {handles}

    def set_mappings(self, mappings: dict):
        with self._lock:
            self.mappings = dict(mappings)

    def set_recording(self, recording: bool):
        """Tell the monitor to block every key event (not just mapped ones)
        while the GUI is capturing a key to assign."""
        with self._lock:
            self._recording = bool(recording)

    def start(self):
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True, name="DualBoard-Monitor")
        self._thread.start()

    def stop(self):
        if not self._running:
            return
        self._running = False
        if self._thread_id is not None:
            user32.PostThreadMessageW(self._thread_id, WM_QUIT, 0, 0)
        if self._thread is not None:
            self._thread.join(timeout=2.0)
        self._thread = None
        self._thread_id = None

    # ----- Worker thread -----

    def _run(self):
        self._thread_id = kernel32.GetCurrentThreadId()

        # Create a message-only window to receive WM_INPUT
        hinst = kernel32.GetModuleHandleW(None)
        wc = WNDCLASS()
        wc.lpfnWndProc = ctypes.cast(self._wnd_proc, ctypes.c_void_p).value
        wc.hInstance = hinst
        wc.lpszClassName = "DualBoardMsgWindow"
        self._atom = user32.RegisterClassW(byref(wc))

        self._hwnd = user32.CreateWindowExW(
            0, "DualBoardMsgWindow", "DualBoard", 0,
            0, 0, 0, 0, HWND_MESSAGE, None, hinst, None,
        )
        if not self._hwnd:
            print("CreateWindowExW failed:", ctypes.get_last_error())
            return

        # Register for keyboard raw input
        rid = RAWINPUTDEVICE()
        rid.usUsagePage = 0x01
        rid.usUsage = 0x06
        rid.dwFlags = RIDEV_INPUTSINK
        rid.hwndTarget = self._hwnd
        if not user32.RegisterRawInputDevices(byref(rid), 1, sizeof(RAWINPUTDEVICE)):
            print("RegisterRawInputDevices failed:", ctypes.get_last_error())

        # Install low-level keyboard hook
        self._hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self._hook_proc, hinst, 0)
        if not self._hook:
            print("SetWindowsHookExW failed:", ctypes.get_last_error())

        # Message pump
        msg = MSG()
        while True:
            ret = user32.GetMessageW(byref(msg), None, 0, 0)
            if ret == 0 or ret == -1:
                break
            user32.TranslateMessage(byref(msg))
            user32.DispatchMessageW(byref(msg))

        # Cleanup
        if self._hook:
            user32.UnhookWindowsHookEx(self._hook)
            self._hook = None

        # Unregister raw input
        rid.dwFlags = RIDEV_REMOVE
        rid.hwndTarget = None
        user32.RegisterRawInputDevices(byref(rid), 1, sizeof(RAWINPUTDEVICE))

        if self._hwnd:
            user32.DestroyWindow(self._hwnd)
            self._hwnd = None

    def _wnd_proc_impl(self, hwnd, msg, wparam, lparam):
        if msg == WM_INPUT:
            self._handle_raw_input(lparam)
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _handle_raw_input(self, hrawinput):
        """Process a single WM_INPUT event. Updates the device-per-VK cache
        and fires the on_key_event callback (used for key recording)."""
        size = wintypes.UINT(0)
        user32.GetRawInputData(
            hrawinput, RID_INPUT, None, byref(size), sizeof(RAWINPUTHEADER)
        )
        if size.value == 0:
            return
        buf = (ctypes.c_byte * size.value)()
        got = user32.GetRawInputData(
            hrawinput, RID_INPUT, buf, byref(size), sizeof(RAWINPUTHEADER)
        )
        if got == 0xFFFFFFFF:
            return

        raw = ctypes.cast(buf, POINTER(RAWINPUT))[0]
        if raw.header.dwType != RIM_TYPEKEYBOARD:
            return

        vk = raw.keyboard.VKey
        is_up = bool(raw.keyboard.Flags & RI_KEY_BREAK)
        h_device = raw.header.hDevice

        # Update the device-per-VK cache (key-down only — a key-up doesn't
        # change which device "owns" the key).
        if not is_up:
            self._last_device_for_vk[vk] = h_device

        with self._lock:
            is_target = bool(self.target_handles) and h_device in self.target_handles

        # Notify GUI (recording, diagnostics)
        try:
            self.on_key_event(vk, is_up, is_target, h_device)
        except Exception as e:
            print("on_key_event error:", e)

    def _drain_raw_input(self):
        """Pump any pending WM_INPUT messages from the queue (via PeekMessageW)
        to bring the device cache up to date. Called inside the LL hook so we
        know the source device before deciding to block or pass.

        A reentrancy guard prevents infinite loops: PeekMessageW can dispatch
        sent-message callbacks (including another hook invocation), so the
        nested hook call will simply skip the drain and use the cache as-is."""
        if self._draining:
            return
        self._draining = True
        try:
            msg = MSG()
            # Drain ALL pending WM_INPUT messages for our message-only window.
            while user32.PeekMessageW(
                byref(msg), self._hwnd, WM_INPUT, WM_INPUT, PM_REMOVE
            ):
                self._handle_raw_input(msg.lParam)
        finally:
            self._draining = False

    def _hook_proc_impl(self, code, wparam, lparam):
        if code != HC_ACTION:
            return user32.CallNextHookEx(self._hook, code, wparam, lparam)

        kbd = ctypes.cast(lparam, POINTER(KBDLLHOOKSTRUCT))[0]

        # Ignore events we (or anyone else) synthesized.
        if kbd.flags & LLKHF_INJECTED:
            return user32.CallNextHookEx(self._hook, code, wparam, lparam)

        is_down = wparam in (WM_KEYDOWN, WM_SYSKEYDOWN)
        is_up = wparam in (WM_KEYUP, WM_SYSKEYUP)
        if not (is_down or is_up):
            return user32.CallNextHookEx(self._hook, code, wparam, lparam)

        vk = kbd.vkCode

        # Fast path: if we blocked this key's KEYDOWN, we *must* also block
        # its KEYUP — otherwise the OS thinks the key is stuck.
        if is_up and vk in self._blocked_downs:
            self._blocked_downs.discard(vk)
            return 1

        with self._lock:
            is_mapped = vk in self.mappings
            recording = self._recording

        # Only care about mapped keys and keys pressed while recording.
        if not (is_mapped or recording):
            return user32.CallNextHookEx(self._hook, code, wparam, lparam)

        # Drain any queued WM_INPUT messages so our device cache is fresh.
        self._drain_raw_input()

        # Which device last sent this VK?
        last_dev = self._last_device_for_vk.get(vk)
        with self._lock:
            is_target = (
                last_dev is not None
                and bool(self.target_handles)
                and last_dev in self.target_handles
            )

        if not is_target:
            # Not from the macropad (or no device info yet) — let it through.
            return user32.CallNextHookEx(self._hook, code, wparam, lparam)

        # It's from the macropad.  Block it.
        if is_down:
            self._blocked_downs.add(vk)
            if not recording:
                with self._lock:
                    action = self.mappings.get(vk)
                if action is not None:
                    try:
                        self.on_blocked_action(vk, action)
                    except Exception as e:
                        print("on_blocked_action error:", e)
        return 1


# =============================================================================
# Action executor
# =============================================================================


def execute_action(action: dict):
    """Run an action dict. Supported types: launch, open_url, shell, send_text, send_keys."""
    try:
        kind = action.get("type")
        if kind == "launch":
            target = action.get("path", "")
            args = action.get("args", "")
            if not target:
                return
            if args:
                subprocess.Popen(f'"{target}" {args}', shell=False)
            else:
                os.startfile(target)

        elif kind == "open_url":
            url = action.get("url", "")
            if url:
                webbrowser.open(url)

        elif kind == "shell":
            cmd = action.get("command", "")
            if cmd:
                subprocess.Popen(cmd, shell=True)

        elif kind == "send_text":
            text = action.get("text", "")
            _send_unicode_text(text)

        elif kind == "send_keys":
            combo = action.get("keys", "")
            _send_key_combo(combo)

        else:
            print(f"Unknown action type: {kind}")

    except Exception as e:
        print(f"Action execution failed: {e}")


def _send_unicode_text(text: str):
    """Type out text via SendInput, character by character."""
    inputs = []
    for ch in text:
        # Key down
        i = INPUT()
        i.type = INPUT_KEYBOARD
        i.union.ki.wVk = 0
        i.union.ki.wScan = ord(ch)
        i.union.ki.dwFlags = KEYEVENTF_UNICODE
        inputs.append(i)
        # Key up
        i2 = INPUT()
        i2.type = INPUT_KEYBOARD
        i2.union.ki.wVk = 0
        i2.union.ki.wScan = ord(ch)
        i2.union.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
        inputs.append(i2)
    if inputs:
        arr = (INPUT * len(inputs))(*inputs)
        user32.SendInput(len(inputs), arr, sizeof(INPUT))


# Names we accept in "send keys" strings
KEY_COMBO_ALIASES = {
    "ctrl": 0x11, "control": 0x11,
    "alt": 0x12, "shift": 0x10, "win": 0x5B, "super": 0x5B,
    "enter": 0x0D, "return": 0x0D, "tab": 0x09, "esc": 0x1B, "escape": 0x1B,
    "space": 0x20, "backspace": 0x08, "delete": 0x2E, "del": 0x2E,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
    "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
}


def _send_key_combo(combo: str):
    """Parse and send a combo like 'ctrl+shift+esc'."""
    if not combo:
        return
    parts = [p.strip().lower() for p in combo.split("+")]
    vks = []
    for p in parts:
        if p in KEY_COMBO_ALIASES:
            vks.append(KEY_COMBO_ALIASES[p])
        elif len(p) == 1:
            vks.append(ord(p.upper()))
        elif p.startswith("f") and p[1:].isdigit():
            n = int(p[1:])
            if 1 <= n <= 24:
                vks.append(0x6F + n)  # VK_F1 = 0x70
    if not vks:
        return
    # Press all down in order
    inputs = []
    for vk in vks:
        i = INPUT()
        i.type = INPUT_KEYBOARD
        i.union.ki.wVk = vk
        inputs.append(i)
    # Release in reverse
    for vk in reversed(vks):
        i = INPUT()
        i.type = INPUT_KEYBOARD
        i.union.ki.wVk = vk
        i.union.ki.dwFlags = KEYEVENTF_KEYUP
        inputs.append(i)
    arr = (INPUT * len(inputs))(*inputs)
    user32.SendInput(len(inputs), arr, sizeof(INPUT))


# =============================================================================
# Config persistence
# =============================================================================

CONFIG_PATH = Path.home() / ".dualboard.json"


def load_config():
    if not CONFIG_PATH.exists():
        return {"device_raw_name": "", "mappings": {}}
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("device_raw_name", "")
        data.setdefault("mappings", {})
        return data
    except Exception as e:
        print(f"Config load failed: {e}")
        return {"device_raw_name": "", "mappings": {}}


def save_config(config):
    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        print(f"Config save failed: {e}")


# =============================================================================
# GUI
# =============================================================================


class ActionDialog(tk.Toplevel):
    """Dialog for editing a single mapping: key + action."""

    RESULT = None

    def __init__(self, parent, app, existing=None):
        super().__init__(parent)
        self.title("Edit Mapping")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.app = app
        self.result = None
        self.recorded_vk = None

        existing = existing or {}
        self.recorded_vk = existing.get("vk")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        # --- Key section ---
        keyframe = ttk.LabelFrame(frm, text="Trigger Key", padding=8)
        keyframe.pack(fill="x")

        row = ttk.Frame(keyframe)
        row.pack(fill="x")
        ttk.Label(row, text="Key:").pack(side="left")
        self.key_label_var = tk.StringVar(
            value=vk_name(self.recorded_vk) if self.recorded_vk else "(none)"
        )
        ttk.Label(row, textvariable=self.key_label_var, width=16,
                  font=("Segoe UI", 10, "bold")).pack(side="left", padx=6)
        self.record_btn = ttk.Button(row, text="● Record Key", command=self._start_recording)
        self.record_btn.pack(side="right")

        ttk.Label(keyframe, text="(press a key on the macropad after clicking Record)",
                  foreground="gray").pack(anchor="w", pady=(4, 0))

        # --- Action section ---
        actframe = ttk.LabelFrame(frm, text="Action", padding=8)
        actframe.pack(fill="x", pady=(10, 0))

        row = ttk.Frame(actframe)
        row.pack(fill="x")
        ttk.Label(row, text="Type:").pack(side="left")

        self.kind_var = tk.StringVar(value=existing.get("type", "launch"))
        self.kind_combo = ttk.Combobox(
            row, textvariable=self.kind_var, state="readonly",
            values=["launch", "open_url", "shell", "send_text", "send_keys"],
            width=14,
        )
        self.kind_combo.pack(side="left", padx=6)
        self.kind_combo.bind("<<ComboboxSelected>>", lambda e: self._update_fields())

        self.fields_frame = ttk.Frame(actframe)
        self.fields_frame.pack(fill="x", pady=(8, 0))

        # Dynamic fields populated per action type
        self.field_widgets = {}
        self._existing = existing
        self._update_fields()

        # --- Buttons ---
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right")
        ttk.Button(btns, text="OK", command=self._ok).pack(side="right", padx=6)

        self._recording = False

    def _update_fields(self):
        # Clear existing
        for w in self.fields_frame.winfo_children():
            w.destroy()
        self.field_widgets = {}

        kind = self.kind_var.get()
        ex = self._existing if self._existing.get("type") == kind else {}

        def add_entry(label, key, default="", browse=False):
            row = ttk.Frame(self.fields_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=label, width=10).pack(side="left")
            var = tk.StringVar(value=ex.get(key, default))
            ent = ttk.Entry(row, textvariable=var)
            ent.pack(side="left", fill="x", expand=True)
            self.field_widgets[key] = var
            if browse:
                def _browse():
                    p = filedialog.askopenfilename(parent=self)
                    if p:
                        var.set(p)
                ttk.Button(row, text="Browse", command=_browse).pack(side="left", padx=(6, 0))

        if kind == "launch":
            add_entry("Path:", "path", browse=True)
            add_entry("Args:", "args")
        elif kind == "open_url":
            add_entry("URL:", "url", default="https://")
        elif kind == "shell":
            add_entry("Command:", "command")
        elif kind == "send_text":
            add_entry("Text:", "text")
        elif kind == "send_keys":
            add_entry("Keys:", "keys", default="ctrl+shift+esc")

    def _start_recording(self):
        self.record_btn.configure(text="Press a key...", state="disabled")
        self._recording = True
        self.app.set_recording_callback(self._key_recorded)

    def _key_recorded(self, vk):
        self._recording = False
        self.recorded_vk = vk
        self.key_label_var.set(vk_name(vk))
        self.record_btn.configure(text="● Record Key", state="normal")

    def _ok(self):
        if self.recorded_vk is None:
            messagebox.showwarning("No key", "Please record a key first.", parent=self)
            return
        action = {"vk": self.recorded_vk, "type": self.kind_var.get()}
        for k, var in self.field_widgets.items():
            action[k] = var.get()
        self.result = action
        self.app.set_recording_callback(None)
        self.destroy()

    def _cancel(self):
        self.app.set_recording_callback(None)
        self.destroy()


class DualBoardApp:

    def __init__(self):
        self.config = load_config()
        self.devices = []  # list of (handle, friendly, raw_name)

        # Recording callback (set when ActionDialog is waiting for a key)
        self._recording_cb = None
        self._recording_lock = threading.Lock()

        self.monitor = KeyboardMonitor(
            on_key_event=self._on_key_event,
            on_blocked_action=self._on_blocked_action,
        )
        self._running = False

        # --- Build GUI ---
        self.root = tk.Tk()
        self.root.title("DualBoard")
        self.root.geometry("620x480")
        self.root.minsize(560, 400)

        try:
            style = ttk.Style()
            if "vista" in style.theme_names():
                style.theme_use("vista")
        except Exception:
            pass

        self._build_gui()
        self._refresh_devices()
        self._push_mappings()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_gui(self):
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)

        # --- Top: device selection ---
        top = ttk.LabelFrame(outer, text="Macropad Device", padding=8)
        top.pack(fill="x")

        row1 = ttk.Frame(top)
        row1.pack(fill="x")
        ttk.Label(row1, text="Device:").pack(side="left")
        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(
            row1, textvariable=self.device_var, state="readonly"
        )
        self.device_combo.pack(side="left", fill="x", expand=True, padx=6)
        self.device_combo.bind("<<ComboboxSelected>>", lambda e: self._on_device_change())
        ttk.Button(row1, text="↻ Refresh", command=self._refresh_devices).pack(side="left")

        # --- Middle: mappings list ---
        mid = ttk.LabelFrame(outer, text="Key Mappings", padding=8)
        mid.pack(fill="both", expand=True, pady=(10, 0))

        listframe = ttk.Frame(mid)
        listframe.pack(fill="both", expand=True)

        cols = ("key", "type", "details")
        self.tree = ttk.Treeview(listframe, columns=cols, show="headings", height=10)
        self.tree.heading("key", text="Key")
        self.tree.heading("type", text="Action")
        self.tree.heading("details", text="Details")
        self.tree.column("key", width=90, anchor="w")
        self.tree.column("type", width=110, anchor="w")
        self.tree.column("details", width=320, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(listframe, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<Double-1>", lambda e: self._edit_mapping())

        btnrow = ttk.Frame(mid)
        btnrow.pack(fill="x", pady=(6, 0))
        ttk.Button(btnrow, text="+ Add", command=self._add_mapping).pack(side="left")
        ttk.Button(btnrow, text="Edit", command=self._edit_mapping).pack(side="left", padx=4)
        ttk.Button(btnrow, text="Remove", command=self._remove_mapping).pack(side="left")

        # --- Bottom: start/stop + status ---
        bot = ttk.Frame(outer)
        bot.pack(fill="x", pady=(10, 0))

        self.start_btn = ttk.Button(bot, text="▶ Start Monitoring", command=self._toggle_monitoring)
        self.start_btn.pack(side="left")

        self.status_var = tk.StringVar(value="Idle")
        ttk.Label(bot, textvariable=self.status_var, foreground="gray").pack(
            side="left", padx=10
        )

        self._refresh_mappings_list()

    # ----- Device handling -----

    def _refresh_devices(self):
        self.devices = enumerate_keyboards()
        names = [f"{friendly}" for _, friendly, _ in self.devices]
        self.device_combo["values"] = names

        # Restore previously selected device by raw_name
        saved_raw = self.config.get("device_raw_name", "")
        idx = None
        for i, (_, _, raw) in enumerate(self.devices):
            if raw == saved_raw:
                idx = i
                break
        if idx is not None:
            self.device_combo.current(idx)
            self._on_device_change()
        elif names:
            # Don't auto-select — force user to pick
            self.device_var.set("")

    def _on_device_change(self):
        idx = self.device_combo.current()
        if idx < 0 or idx >= len(self.devices):
            return
        handles, friendly, raw = self.devices[idx]
        self.monitor.set_target(handles)
        self.config["device_raw_name"] = raw
        save_config(self.config)
        self.status_var.set(f"Target: {friendly}")

    # ----- Mappings -----

    def _refresh_mappings_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for vk_str, action in self.config["mappings"].items():
            vk = int(vk_str)
            details = self._action_summary(action)
            self.tree.insert("", "end", iid=vk_str,
                             values=(vk_name(vk), action.get("type", ""), details))

    def _action_summary(self, action):
        t = action.get("type")
        if t == "launch":
            return action.get("path", "")
        if t == "open_url":
            return action.get("url", "")
        if t == "shell":
            return action.get("command", "")
        if t == "send_text":
            return repr(action.get("text", ""))
        if t == "send_keys":
            return action.get("keys", "")
        return ""

    def _add_mapping(self):
        dlg = ActionDialog(self.root, self)
        self.root.wait_window(dlg)
        if dlg.result:
            vk = dlg.result.pop("vk")
            self.config["mappings"][str(vk)] = dlg.result
            save_config(self.config)
            self._refresh_mappings_list()
            self._push_mappings()

    def _edit_mapping(self):
        sel = self.tree.selection()
        if not sel:
            return
        vk_str = sel[0]
        existing = dict(self.config["mappings"][vk_str])
        existing["vk"] = int(vk_str)
        dlg = ActionDialog(self.root, self, existing=existing)
        self.root.wait_window(dlg)
        if dlg.result:
            old_vk = int(vk_str)
            new_vk = dlg.result.pop("vk")
            if new_vk != old_vk:
                del self.config["mappings"][vk_str]
            self.config["mappings"][str(new_vk)] = dlg.result
            save_config(self.config)
            self._refresh_mappings_list()
            self._push_mappings()

    def _remove_mapping(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("Remove", "Remove this mapping?", parent=self.root):
            return
        del self.config["mappings"][sel[0]]
        save_config(self.config)
        self._refresh_mappings_list()
        self._push_mappings()

    def _push_mappings(self):
        mappings = {int(k): v for k, v in self.config["mappings"].items()}
        self.monitor.set_mappings(mappings)

    # ----- Monitoring -----

    def _toggle_monitoring(self):
        if self._running:
            self.monitor.stop()
            self._running = False
            self.start_btn.configure(text="▶ Start Monitoring")
            self.status_var.set("Stopped")
        else:
            if self.device_combo.current() < 0:
                messagebox.showwarning(
                    "No device", "Select a macropad device first.", parent=self.root
                )
                return
            self.monitor.start()
            self._running = True
            self.start_btn.configure(text="■ Stop Monitoring")
            idx = self.device_combo.current()
            _, friendly, _ = self.devices[idx]
            self.status_var.set(f"Monitoring: {friendly}")

    # ----- Key event callbacks (called from worker thread) -----

    def set_recording_callback(self, cb):
        """Called by ActionDialog to request the next keypress."""
        with self._recording_lock:
            self._recording_cb = cb
        # Tell the monitor to block every key while we're capturing one,
        # not just mapped keys — otherwise the key the user presses during
        # recording reaches applications (annoying) or, if it's a mapped
        # key being re-recorded, fires its old action.
        self.monitor.set_recording(cb is not None)
        # Ensure monitor is running so we can see keys during recording
        if cb is not None and not self._running:
            # Start in recording-only mode (empty mappings means nothing is blocked)
            self.monitor.start()
            self._running = True
            self.start_btn.configure(text="■ Stop Monitoring")
            self.status_var.set("Recording key...")

    def _on_key_event(self, vk, is_up, is_target, h_device):
        """From worker thread: any key from any device."""
        if is_up:
            return
        with self._recording_lock:
            cb = self._recording_cb
        if cb is not None and is_target:
            # Marshal to GUI thread
            self.root.after(0, lambda: cb(vk))

    def _on_blocked_action(self, vk, action):
        """From worker thread: a bound key fired. Run it off-thread."""
        self.root.after(0, lambda: self.status_var.set(f"Triggered: {vk_name(vk)}"))
        threading.Thread(target=execute_action, args=(action,), daemon=True).start()

    # ----- Lifecycle -----

    def _on_close(self):
        if self._running:
            self.monitor.stop()
        save_config(self.config)
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def main():
    if sys.platform != "win32":
        print("DualBoard requires Windows.")
        sys.exit(1)
    app = DualBoardApp()
    app.run()


if __name__ == "__main__":
    main()
