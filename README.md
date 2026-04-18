# DualBoard

A per-device keyboard macro manager for Windows. Designate a secondary keyboard
(like a Keychron Q0 numpad) as a macropad, and remap its keys to custom actions
*without* affecting the same keys on your main keyboard.

## How it works

The OS normally merges input from all your keyboards into a single stream —
which is why a `Num0` from a USB numpad is indistinguishable from a `Num0` on
your main keyboard. DualBoard solves this with two layered Win32 APIs:

1. **Raw Input API** (`WM_INPUT`) identifies *which* physical keyboard sent
   each keystroke by its device handle.
2. **Low-level keyboard hook** (`WH_KEYBOARD_LL`) intercepts every keypress
   system-wide before apps receive it.

When a key arrives, DualBoard correlates the two event streams: if the most
recent raw-input event for that key came from your chosen macropad, the key
is **blocked** and the mapped action runs instead. Keys from any other
keyboard pass through untouched.

Unlike tools that rely on a kernel driver (Interception), this uses only
standard user-mode APIs — no admin rights, no driver install.

## Requirements

- Windows 10 or 11
- Python 3.8 or newer (tkinter and ctypes are stdlib — nothing to `pip install`)

## Running

```
python dualboard.py
```

Steps:
1. Plug in both keyboards.
2. Click **Refresh** — you should see all connected keyboards listed.
3. Pick your macropad from the dropdown (device names come from the Windows
   registry; something like "HID Keyboard Device" or the vendor's name).
4. Click **+ Add** to create a mapping. Hit **● Record Key**, then press the
   key on your macropad. Choose an action type, fill in the fields, click OK.
5. Click **▶ Start Monitoring**. That key on the macropad now triggers your
   action; the same key on any other keyboard behaves normally.

Config is saved to `~/.dualboard.json` automatically.

## Action types

| Type | What it does | Fields |
|---|---|---|
| `launch` | Start an application | Path (Browse), optional Args |
| `open_url` | Open a URL in your default browser | URL |
| `shell` | Run an arbitrary shell command | Command |
| `send_text` | Type out a string | Text |
| `send_keys` | Send a hotkey combo | Keys (e.g. `ctrl+shift+esc`) |

For `send_keys`, accepted names include `ctrl`, `alt`, `shift`, `win`, `enter`,
`tab`, `esc`, `space`, `f1`–`f24`, arrow keys, `home`, `end`, `pageup`,
`pagedown`, and single letters/digits.

## Identifying your macropad

If several devices look alike in the dropdown, unplug your macropad, click
**Refresh**, note which one disappears, plug it back in, refresh again.
Pick the one that reappeared. DualBoard remembers the exact device (by its
raw device path) so it survives reboots even if USB ports change.

## Limitations

- **Windows only.** The approach is Win32-specific. Mac has Karabiner-Elements
  with native device filtering; Linux has `keyd` and `evdev`.
- **No blocking during a rare race window.** The correlation between Raw
  Input and the low-level hook has a ~100ms window. In practice, Raw Input
  is delivered first for the vast majority of keypresses, but in pathological
  cases (very high CPU load) a key might slip through unblocked. If you see
  this, the fallback is to put the Q0 into a VIA-remapped mode sending F13–F24
  so blocking doesn't matter.
- **Injected keys are ignored.** DualBoard deliberately ignores keypresses
  it or other programs synthesize (via `SendInput`), so `send_text` and
  `send_keys` actions won't recurse into themselves.

## Packaging as an .exe (optional)

If you don't want to keep Python around:

```
pip install pyinstaller
pyinstaller --onefile --noconsole --name DualBoard dualboard.py
```

The resulting `dist/DualBoard.exe` is self-contained.

## Troubleshooting

- **No keyboards listed** — Run `python dualboard.py` from a console to see
  error output. The Raw Input API should work for any standard user; if
  `RegisterRawInputDevices failed` appears, something unusual is blocking it.
- **Actions trigger but the original key ALSO fires** — The correlation
  didn't match in time. Try stopping and restarting monitoring. If it
  persists, the VIA-remap-to-F-keys fallback is more robust.
- **A bound key does nothing** — Make sure monitoring is started (the button
  should read "■ Stop Monitoring"), and that you selected the right device.
