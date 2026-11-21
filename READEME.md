# TackyZones

A PowerShell script that saves and restores the size and position of all open windows on Windows 10/11, using only built-in Windows APIs (`user32.dll`, `dwmapi.dll`, `shcore.dll`). No third-party software required.

## What it does

- **Save**: Records the position, size, title, owning process, and maximized state of every visible top-level window to a JSON file (default: `%USERPROFILE%\window-layout.json`).

- **Restore**: Re-enumerates the currently open windows, matches them to the saved entries by title (with process name as a sanity check), and moves each one back to its saved size and position.

Typical use cases: re-establishing your working layout after reconnecting external monitors, resuming a session after a reboot, or snapping back after a projector/TV hookup scrambles everything.

## Requirements

- Windows 10 or 11
- Windows PowerShell 5.1 or PowerShell 7.x
- Run at the same integrity level as the windows you want to capture (a normal, non-elevated session handles your normal windows fine)

## Usage

```powershell
# Save the current layout
.\Set-WindowLayout.ps1 -Action Save

# Restore a previously saved layout
.\Set-WindowLayout.ps1 -Action Restore

# Use a custom layout file
.\Set-WindowLayout.ps1 -Action Save -Path "D:\layouts\work.json"
.\Set-WindowLayout.ps1 -Action Restore -Path "D:\layouts\work.json"
```

### Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `Action`  | Yes | — | `Save` or `Restore` |
| `Path`    | No | `%USERPROFILE%\window-layout.json` | Path to the JSON layout file |

### One-keystroke save/restore (optional)

Create two shortcuts to the script, one per action, then assign a shortcut key in each shortcut's Properties:

```
powershell -ExecutionPolicy Bypass -File "C:\path\Set-WindowLayout.ps1" -Action Save
powershell -ExecutionPolicy Bypass -File "C:\path\Set-WindowLayout.ps1" -Action Restore
```

For example, assign `Ctrl+Alt+S` and `Ctrl+Alt+R`.

## How it works

- **Matching is by title, not window handle.** Window handles (HWND) change between sessions, so saved layouts store each window's title and owning process name. At restore time, currently open windows are matched by exact title, with the process name verified as a sanity check so an unrelated window sharing a title isn't grabbed.
- **Duplicate titles are handled.** Live windows are consumed in saved order, so several windows with the same title each get their own saved position.
- **Coordinates are clamped to the virtual desktop**, so a saved position never restores a window off-screen (e.g., after unplugging a monitor).
- **DPI awareness is declared at startup** (per-monitor v2, with automatic fallbacks). Save and restore operate on true physical pixels, so layouts are pixel-exact on mixed-DPI multi-monitor setups. If you change display scaling between save and restore, windows return to the same physical spot but occupy relatively more or less of the logical desktop.

## Limitations

These are inherent to how Windows manages windows, not bugs in the script:

- **Apps must be open when you restore.** The script moves existing windows; it does not relaunch applications.
- **Apps that manage their own placement** (some Electron/UWP apps) may reposition themselves after being moved.
- **Minimized windows are not saved** — Windows does not report their real position.
- **Windows without unique titles** may not match reliably.

## Troubleshooting

**`Restored 0 of 1 window(s)` with `System.Object[]` in the "not restored" list**
Older versions of the script wrapped `ConvertFrom-Json` output in `@(...)`. Under Windows PowerShell 5.1, `ConvertFrom-Json` emits a top-level JSON array as a *single* object, so the `@()` wrapper produced a one-element array containing the whole collection — every window appeared unmatched. Fixed in the current version by explicitly unwrapping the parsed result (PowerShell 7.x enumerates by default and is unaffected).

**DPI mode reports `System` instead of `PerMonitorV2`**
The per-monitor v2 call failed, usually because the hosting console already locked in a DPI-awareness context. On a single monitor or uniform scaling this is harmless; on a mixed-DPI setup, try launching with `powershell -NoProfile` to give the script the first shot at declaring awareness.

**Some windows don't restore**
Check the "Not restored" list printed after a restore. Entries there were either not open, or no live window matched the saved title/process pair. Open the app in question and restore again.

## Layout file format

The JSON file is human-readable and editable:

```json
[
  {
    "Process":  "msedge",
    "Title":    "Some page - Microsoft Edge",
    "Left":     -1920,
    "Top":      0,
    "Width":    1920,
    "Height":   1040,
    "Maximized": false
  }
]
```

## Changelog

- **1.4** — Fixed restore under Windows PowerShell 5.1: top-level JSON array from `ConvertFrom-Json` is now unwrapped explicitly instead of relying on `@()` enumeration, which fails on 5.1.
- **1.3** — Fixed parse error at startup caused by using a cmdlet call as a parameter default value; defaults must be constant expressions.
- **1.2** — Added per-monitor DPI awareness (v2) so save and restore operate on physical pixels; removes drift on mixed-DPI multi-monitor setups.
- **1.1** — Fixed crashes on duplicate titles, uninitialized handle tracking, desktop/shell windows being captured, pseudo-maximized restores, and repeated `Add-Type` failures; replaced O(n²) process lookups with a single lookup table.

## License

Provided as-is, free to use and modify.
