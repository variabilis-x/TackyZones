<#
.SYNOPSIS
  Saves and restores the size and position of all open windows on Windows 10/11.
  Uses only built-in Windows APIs (user32.dll / dwmapi.dll) - no third-party software.

.DESCRIPTION
  Save   : records the position/size/title/process/maximized-state of every
           visible top-level window to a JSON file
           (default: %USERPROFILE%\window-layout.json).
  Restore: re-enumerates currently open windows, matches them to the saved
           entries by window title (+ process as a sanity check), and moves
           them back to their saved rectangles, re-maximizing where needed.

  Matching is by TITLE because window handles (HWND) change between sessions,
  so apps without unique titles may not be matched. Restored coordinates are
  clamped to the current virtual desktop area so nothing lands off-screen.

.USAGE
  Save layout:
      .\Set-WindowLayout.ps1 -Action Save
  Restore layout:
      .\Set-WindowLayout.ps1 -Action Restore
  Custom file:
      .\Set-WindowLayout.ps1 -Action Save -Path "D:\layouts\work.json"

  Tip: create two shortcuts to this script, one with target
  `powershell -ExecutionPolicy Bypass -File "C:\path\Set-WindowLayout.ps1" -Action Save`
  and one with `-Action Restore`, then assign a Shortcut key (e.g. Ctrl+Alt+S /
  Ctrl+Alt+R) in each shortcut's Properties for one-keystroke save/restore.

.CHANGELOG
  1.1  Fixed: crash when several windows shared a title (invalid
       'Select-Object -Object' parameter); uninitialized $usedHandles array;
       desktop/shell windows (Progman/WorkerW/SHELLDLL_DefView) being captured;
       maximized windows restored as pseudo-maximized normal windows;
       Add-Type failing when the script ran twice in one session;
       O(n^2) Get-Process calls replaced with a single lookup table.

.NOTES
  Limitations (inherent to Windows, not the script):
   - The app must be OPEN when you restore; this moves existing windows, it
     does not relaunch apps.
   - Apps that manage their own window placement (some Electron/UWP apps)
     may fight back and reposition themselves.
   - With mixed-DPI multi-monitor setups, saved coordinates can be slightly
     off on restore because of DPI virtualization.
   - Minimized windows are not saved (Windows hides their real position).
   - Requires PowerShell run at the same integrity level as the target windows
     (a normal, non-elevated session handles your normal windows fine).
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Save', 'Restore')]
    [string]$Action,

    [string]$Path = Join-Path $env:USERPROFILE 'window-layout.json'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- Win32 helpers
# Guard so running the script twice in one session does not fail with
# "type already exists".
if (-not ('WinLayout.Helper' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace WinLayout
{
    public struct WinInfo
    {
        public IntPtr Handle;
        public int    Pid;
        public string Title;
        public bool   Zoomed;                  // currently maximized?
        public int    Left, Top, Right, Bottom;
    }

    public class Helper
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int L, T, R, B; }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr h);                 // minimized?
        [DllImport("user32.dll")]
        private static extern bool IsZoomed(IntPtr h);                  // maximized?
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int  GetWindowText(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int  GetWindowTextLength(IntPtr h);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int  GetClassName(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr after,
            int x, int y, int w, int h, uint flags);
        [DllImport("user32.dll")]
        private static extern int  GetWindowLong(IntPtr h, int index);
        [DllImport("user32.dll")]
        private static extern int  GetSystemMetrics(int index);        // virtual screen
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr h, int cmd);
        [DllImport("dwmapi.dll")]
        private static extern int  DwmGetWindowAttribute(IntPtr h, int attr,
            out int value, int size);

        private const int GWL_EXSTYLE       = -20;
        private const int WS_EX_TOOLWINDOW  = 0x00000080;
        private const int DWMWA_CLOAKED     = 14;   // hidden UWP/desktops windows
        private const uint SWP_NOZORDER      = 0x0004;
        private const uint SWP_NOACTIVATE    = 0x0010;
        private const int  SW_MAXIMIZE      = 3;

        // Window classes that belong to the shell / desktop, not real apps.
        private static readonly HashSet<string> ExcludedClasses =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "Progman", "WorkerW", "SHELLDLL_DefView" };

        private static List<WinInfo> _windows;

        private static bool Callback(IntPtr h, IntPtr lParam)
        {
            // Only real, visible, non-minimized windows with a title.
            if (!IsWindowVisible(h)) return true;
            if (IsIconic(h))         return true;

            int exStyle = GetWindowLong(h, GWL_EXSTYLE);
            if ((exStyle & WS_EX_TOOLWINDOW) != 0) return true;

            // Skip cloaked windows (UWP apps running in background, other desktops).
            int cloaked;
            if (DwmGetWindowAttribute(h, DWMWA_CLOAKED, out cloaked, sizeof(int)) == 0
                && cloaked != 0)
                return true;

            if (GetWindowTextLength(h) == 0) return true;

            // Skip the desktop / shell windows so they never enter the layout.
            var cls = new StringBuilder(256);
            GetClassName(h, cls, cls.Capacity);
            if (ExcludedClasses.Contains(cls.ToString())) return true;

            var sb = new StringBuilder(GetWindowTextLength(h) + 1);
            GetWindowText(h, sb, sb.Capacity);

            uint pid;
            GetWindowThreadProcessId(h, out pid);

            RECT r;
            GetWindowRect(h, out r);

            _windows.Add(new WinInfo {
                Handle = h,
                Pid    = (int)pid,
                Title  = sb.ToString(),
                Zoomed  = IsZoomed(h),
                Left   = r.L, Top = r.T, Right = r.R, Bottom = r.B
            });
            return true;
        }

        public static WinInfo[] Capture()
        {
            _windows = new List<WinInfo>();
            EnumWindows(Callback, IntPtr.Zero);
            return _windows.ToArray();
        }

        // Move a window to the given rectangle, clamped to the visible virtual
        // desktop so a saved position never ends up off-screen.
        public static bool Place(IntPtr h, int left, int top, int right, int bottom)
        {
            int vsX = GetSystemMetrics(76);   // SM_XVIRTUALSCREEN
            int vsY = GetSystemMetrics(77);   // SM_YVIRTUALSCREEN
            int vsW = GetSystemMetrics(78);   // SM_CXVIRTUALSCREEN
            int vsH = GetSystemMetrics(79);   // SM_CYVIRTUALSCREEN

            if (right - left > vsW) { right  = left + vsW; }
            if (bottom - top > vsH) { bottom = top + vsH; }
            if (left < vsX)       { right  += vsX - left;  left  = vsX; }
            if (top  < vsY)       { bottom += vsY - top;   top   = vsY; }
            if (right > vsX + vsW) { left   -= right - (vsX + vsW); right  = vsX + vsW; }
            if (bottom > vsY + vsH){ top    -= bottom - (vsY + vsH); bottom = vsY + vsH; }

            return SetWindowPos(h, IntPtr.Zero, left, top,
                right - left, bottom - top, SWP_NOZORDER | SWP_NOACTIVATE);
        }

        // Re-maximize a window after it has been placed (async so a hung
        // application cannot stall the script).
        public static bool Maximize(IntPtr h)
        {
            return ShowWindowAsync(h, SW_MAXIMIZE);
        }
    }
}
'@
}

# Build a single PID -> process-name lookup so we never call Get-Process in a
# loop (the old version did this per comparison, O(n^2)).
function Get-ProcessNameLookup {
    $lookup = @{}
    foreach ($p in Get-Process) { $lookup[$p.Id] = $p.ProcessName }
    return $lookup
}

switch ($Action) {

    # ------------------------------------------------------------- SAVE
    'Save' {
        $captured = [WinLayout.Helper]::Capture()
        $procName = Get-ProcessNameLookup

        $entries = foreach ($w in $captured) {
            if ([string]::IsNullOrWhiteSpace($w.Title)) { continue }
            if (-not $procName.ContainsKey($w.Pid))    { continue }

            [pscustomobject]@{
                Process  = $procName[$w.Pid]
                Title    = $w.Title
                Left     = $w.Left
                Top      = $w.Top
                Width    = $w.Right  - $w.Left
                Height   = $w.Bottom - $w.Top
                Maximized = $w.Zoomed
            }
        }

        if (-not $entries) {
            Write-Warning 'No windows found to save.'
            return
        }

        $entries | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $Path -Encoding Unicode
        Write-Host ("Saved {0} window(s) to '{1}'." -f @($entries).Count, $Path)
    }

    # ---------------------------------------------------------- RESTORE
    'Restore' {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Warning "No saved layout found at '$Path'. Run -Action Save first."
            return
        }

        $saved    = @(Get-Content -LiteralPath $Path -Raw -Encoding Unicode | ConvertFrom-Json)
        $current  = [WinLayout.Helper]::Capture()
        $procName = Get-ProcessNameLookup

        # Track which live windows we have already consumed, so several
        # same-titled windows each get their own position (in saved order).
        $usedHandles = @()

        $restored = 0
        $missed   = New-Object System.Collections.Generic.List[string]

        foreach ($s in $saved) {

            if ([string]::IsNullOrWhiteSpace($s.Title)) { continue }

            # Match on exact title; verify the process name agrees so we do not
            # grab an unrelated window that happens to share a title.
            $candidates = $current | Where-Object {
                $_.Title -eq $s.Title -and
                $procName.ContainsKey($_.Pid) -and
                $procName[$_.Pid] -eq $s.Process
            }

            $match = $candidates |
                Where-Object { $usedHandles -notcontains $_.Handle } |
                Select-Object -First 1        # <-- was: Select-Object -Object -First 1 (crashed)

            if (-not $match) {
                $missed.Add(('{0} - "{1}"' -f $s.Process, $s.Title))
                continue
            }

            $placed = [WinLayout.Helper]::Place($match.Handle,
                [int]$s.Left, [int]$s.Top,
                ([int]$s.Left + [int]$s.Width), ([int]$s.Top + [int]$s.Height))

            if ($placed) {
                $usedHandles += $match.Handle
                if ($s.Maximized) { [void][WinLayout.Helper]::Maximize($match.Handle) }
                $restored++
            }
            else {
                $missed.Add(('{0} - "{1}"' -f $s.Process, $s.Title))
            }
        }

        Write-Host ("Restored {0} of {1} window(s)." -f $restored, $saved.Count)
        if ($missed.Count -gt 0) {
            Write-Host 'Not restored (not open, or no matching title):'
            $missed | Sort-Object -Unique | ForEach-Object { Write-Host "  - $_" }
        }
    }
}

