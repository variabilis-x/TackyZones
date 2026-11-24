<#
.SYNOPSIS
  Saves and restores the size and position of all open windows on Windows 10/11.
  Uses only built-in Windows APIs (user32.dll / dwmapi.dll / shcore.dll) -
  no third-party software.

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

  PLACEMENT CORRECTION (convergence loop)
  On mixed-DPI systems (e.g. laptop panel at 150% + external monitor at
  100% via a dock), Windows converts coordinates between per-monitor spaces
  for windows whose DPI awareness does not match the monitor scaling. That
  conversion involves rounding, so a single SetWindowPos call can land a
  window a few pixels off the requested rectangle - deterministically, the
  same few pixels every time. Place() therefore closes the loop: it places
  the window, measures where it actually landed with GetWindowRect, adjusts
  the target by the observed error, and retries - up to four iterations or
  until the measured rectangle matches the target exactly. This converges
  on pixel-exact placement regardless of which conversion layer rounds.

  SNAP INTERFERENCE GUARD
  Windows 11's snap machinery (and some apps) can also re-adjust a window
  asynchronously AFTER SetWindowPos reports success. Restore therefore
  verifies every placed non-maximized window ~350 ms later and re-places
  any that moved, up to three rounds.

  DPI AWARENESS
  The script declares per-monitor-DPI-aware (v2) both at the process level
  (best effort) and at the THREAD level around every coordinate-sensitive
  operation (Capture, Place, and the convergence measurements inside
  Place), so all coordinates are TRUE PHYSICAL PIXELS on every
  Windows 10 1607+ system - including when the hosting console has already
  locked the process-level context.
  IMPORTANT: after upgrading from an older version, RE-SAVE your layout once
  before restoring; layouts saved by older versions may contain rounded
  (virtualized) coordinates.

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
  1.7  Fixed: deterministic few-pixel offset on restore, unaffected by
       disabling Windows snap and identical on repeated restores. Cause:
       on mixed-DPI systems, coordinate conversion between per-monitor
       spaces (for windows whose DPI awareness differs from the monitor
       scaling) rounds coordinates in both directions, so a single
       SetWindowPos lands the same few pixels off every time. Fix: Place()
       now converges - place, measure with GetWindowRect, adjust the target
       by the observed error, retry (max 4 iterations, stop when exact).
  1.6  Added: post-placement verification. Windows 11's snap machinery and
       some apps re-adjust windows asynchronously after SetWindowPos reports
       success; Restore re-measures placed windows ~350 ms later and
       re-asserts any that moved (up to three rounds), reporting the count.
  1.5  Fixed: restore landing windows a few pixels off from the saved
       position. Cause: when the process-level per-monitor DPI declaration
       fails (common, because the hosting console usually locks the context
       first), the script fell back to system-DPI-aware mode and all
       GetWindowRect/SetWindowPos calls went through DPI virtualization with
       its rounding on both directions. Fix: the Capture and Place methods
       now switch the CALLING THREAD to per-monitor-v2 DPI awareness via
       SetThreadDpiAwarenessContext (Windows 10 1607+) for the duration of
       the call and restore the previous context afterwards.
  1.4  Fixed: restore reported "Restored 0 of N" with System.Object[] names
       under Windows PowerShell 5.1 (ConvertFrom-Json top-level array
       emitted as a single object). Fix: parsed result is unwrapped
       explicitly (identical behavior on 5.1 and 7.x).
  1.3  Fixed: parse error at startup caused by using a cmdlet call as a
       parameter default value; defaults must be constant expressions.
  1.2  Added: per-monitor DPI awareness (v2) declared at startup.
  1.1  Fixed: crash when several windows shared a title; uninitialized
       $usedHandles array; desktop/shell windows being captured; maximized
       windows restored as pseudo-maximized normal windows; Add-Type failing
       when the script ran twice in one session; O(n^2) Get-Process calls
       replaced with a single lookup table.

.NOTES
  Limitations (inherent to Windows, not the script):
   - The app must be OPEN when you restore; this moves existing windows, it
     does not relaunch apps.
   - Apps that manage their own window placement (some Electron/UWP apps)
     may fight back and reposition themselves. The verification pass
     re-asserts the saved position, but a badly behaved app can win a
     tug-of-war over the long run.
   - Minimized windows are not saved (Windows hides their real position).
   - Requires PowerShell run at the same integrity level as the target windows
     (a normal, non-elevated session handles your normal windows fine).
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Save', 'Restore')]
    [string]$Action,

    # NOTE: parameter default values must be CONSTANT EXPRESSIONS. A cmdlet
    # call like 'Join-Path ...' is a parse error here; an expandable string
    # is evaluated at runtime and is legal.
    [string]$Path = "$env:USERPROFILE\window-layout.json"
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

        // ---- DPI awareness -------------------------------------------------
        // Best effort at process level (only succeeds if the host has not
        // already locked a context). Correctness does not depend on it:
        // Capture() and Place() switch the THREAD context instead.
        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);
        [DllImport("shcore.dll")]
        private static extern int  SetProcessDpiAwareness(int value);
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        // Thread-level context switching (Windows 10 1607+). Works even when
        // the process-level context is locked, which is common with console
        // hosts. Affects GetWindowRect / SetWindowPos on the calling thread.
        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr value);

        private static readonly IntPtr DPI_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
        private const int PROCESS_PER_MONITOR_DPI_AWARE = 2;

        public static string EnsureDpiAwareness()
        {
            // Best: per-monitor v2 (Windows 10 1703+).
            try
            {
                if (SetProcessDpiAwarenessContext(DPI_PER_MONITOR_AWARE_V2))
                    return "PerMonitorV2";
            }
            catch (EntryPointNotFoundException) { /* pre-1703 */ }

            // Second best: per-monitor v1.
            try
            {
                if (SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE) == 0)
                    return "PerMonitorV1";
            }
            catch (EntryPointNotFoundException) { /* pre-Win8.1 */ }

            // Last resort: system-DPI aware (exact on single-DPI systems).
            if (SetProcessDPIAware()) return "System";
            return "None";
        }

        // Enter per-monitor-v2 on the calling thread for the duration of a
        // coordinate-sensitive operation, restoring the previous context
        // afterwards. Returns the previous context, or IntPtr.Zero if the
        // API is unavailable (pre-1607 systems).
        public static IntPtr EnterThreadDpiAwareness()
        {
            try
            {
                IntPtr old = SetThreadDpiAwarenessContext(DPI_PER_MONITOR_AWARE_V2);
                return old; // IntPtr.Zero => API missing, caller must not restore
            }
            catch (EntryPointNotFoundException)
            {
                return IntPtr.Zero;
            }
        }

        public static void ExitThreadDpiAwareness(IntPtr previous)
        {
            if (previous != IntPtr.Zero)
                SetThreadDpiAwarenessContext(previous);
        }

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
            // Thread-level per-monitor-v2 context => GetWindowRect inside the
            // callback returns true physical pixels even when the process
            // context could not be upgraded (avoids DPI-virtualization drift).
            IntPtr oldCtx = EnterThreadDpiAwareness();
            try
            {
                _windows = new List<WinInfo>();
                EnumWindows(Callback, IntPtr.Zero);
                return _windows.ToArray();
            }
            finally
            {
                ExitThreadDpiAwareness(oldCtx);
            }
        }

        // Move a window to the given rectangle, clamped to the visible virtual
        // desktop so a saved position never ends up off-screen.
        //
        // CONVERGENCE LOOP: on mixed-DPI systems the coordinate conversion
        // between per-monitor spaces (for windows whose DPI awareness differs
        // from the monitor scaling) rounds coordinates in both directions, so
        // a single SetWindowPos can land the same few pixels off every time.
        // We therefore place, measure the actual result with GetWindowRect,
        // adjust the target by the observed error and retry - up to four
        // iterations, stopping as soon as the measured rectangle matches.
        public static bool Place(IntPtr h, int left, int top, int right, int bottom)
        {
            IntPtr oldCtx = EnterThreadDpiAwareness();
            try
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

                // Retry loop: place, measure, correct, repeat.
                for (int attempt = 0; attempt < 4; attempt++)
                {
                    if (!SetWindowPos(h, IntPtr.Zero, left, top,
                        right - left, bottom - top, SWP_NOZORDER | SWP_NOACTIVATE))
                        return false;

                    RECT r;
                    if (!GetWindowRect(h, out r))
                        return true;               // placed, but cannot verify

                    int dx = left       - r.L;     // position error (target - actual)
                    int dy = top        - r.T;
                    int dw = (right - left)  - (r.R - r.L);   // size error
                    int dh = (bottom - top)  - (r.B - r.T);

                    if (dx == 0 && dy == 0 && dw == 0 && dh == 0)
                        return true;               // pixel-exact; done

                    // Fold the observed error into the next attempt's target.
                    left   += dx;
                    top    += dy;
                    right  += dx + dw;
                    bottom += dy + dh;
                }
                return true;                       // closest achievable
            }
            finally
            {
                ExitThreadDpiAwareness(oldCtx);
            }
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

# Declare DPI awareness FIRST, before any window enumeration, so every
# coordinate below is in physical pixels. Best effort only: the thread-level
# switch inside Capture/Place guarantees correctness even if this falls back.
$script:DpiMode = [WinLayout.Helper]::EnsureDpiAwareness()

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
        Write-Host ("Saved {0} window(s) to '{1}' (DPI mode: {2})." -f @($entries).Count, $Path, $script:DpiMode)
    }

    # ---------------------------------------------------------- RESTORE
    'Restore' {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Warning "No saved layout found at '$Path'. Run -Action Save first."
            return
        }

        $parsed = Get-Content -LiteralPath $Path -Raw -Encoding Unicode | ConvertFrom-Json

        # Windows PowerShell 5.1 emits a top-level JSON array as a SINGLE
        # object, so wrapping it in @(...) yields a one-element array whose
        # only element is the whole collection (PS 7.x enumerates by default).
        # Unwrap explicitly so both versions behave identically.
        if ($parsed -is [array]) { $saved = $parsed } else { $saved = @($parsed) }

        $current  = [WinLayout.Helper]::Capture()
        $procName = Get-ProcessNameLookup

        # Track which live windows we have already consumed, so several
        # same-titled windows each get their own position (in saved order).
        $usedHandles = @()

        # Track every successful placement (handle -> saved entry) so the
        # verification pass below can re-check them.
        $placements = @{}

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
                Select-Object -First 1

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
                $placements[$match.Handle] = $s
                $restored++
            }
            else {
                $missed.Add(('{0} - "{1}"' -f $s.Process, $s.Title))
            }
        }

        # ------------------------------------------- VERIFICATION PASS
        # Windows 11's snap machinery (and some apps) can re-adjust a window
        # a moment AFTER SetWindowPos has reported success. Re-measure every
        # placed window ~350 ms later and put back any that moved; repeat up
        # to three rounds until everything sticks. Maximized windows are
        # skipped: their on-screen rectangle is owned by the maximize state.
        $reasserted = 0

        $toVerify = @{}
        foreach ($kv in $placements.GetEnumerator()) {
            if (-not $kv.Value.Maximized) { $toVerify[$kv.Key] = $kv.Value }
        }

        if ($toVerify.Count -gt 0) {
            for ($round = 1; $round -le 3; $round++) {
                Start-Sleep -Milliseconds 350
                $live    = [WinLayout.Helper]::Capture()
                $movedAny = $false

                foreach ($kv in $toVerify.GetEnumerator()) {
                    $h = $kv.Key
                    $s = $kv.Value

                    $w = $live | Where-Object { $_.Handle -eq $h } |
                        Select-Object -First 1
                    if (-not $w) { continue }   # window closed/minimized since

                    $offX = $w.Left - [int]$s.Left
                    $offY = $w.Top  - [int]$s.Top
                    $offW = ($w.Right  - $w.Left)   - [int]$s.Width
                    $offH = ($w.Bottom - $w.Top)    - [int]$s.Height

                    if ($offX -ne 0 -or $offY -ne 0 -or $offW -ne 0 -or $offH -ne 0) {
                        # The window wandered after placement - snap state or
                        # the app corrected it. Put it back (Place converges).
                        [void][WinLayout.Helper]::Place($h,
                            [int]$s.Left, [int]$s.Top,
                            ([int]$s.Left + [int]$s.Width),
                            ([int]$s.Top  + [int]$s.Height))
                        $movedAny  = $true
                        $reasserted++
                    }
                }

                if (-not $movedAny) { break }   # everything stuck; done early
            }
        }
        # -------------------------------------------------------------- end

        $summary = ("Restored {0} of {1} window(s) (DPI mode: {2})." -f
            $restored, @($saved).Count, $script:DpiMode)
        if ($reasserted -gt 0) {
            $summary += " Re-asserted {0} placement(s) that Windows or the app adjusted afterwards." -f $reasserted
        }
        Write-Host $summary

        if ($missed.Count -gt 0) {
            Write-Host 'Not restored (not open, or no matching title):'
            $missed | Sort-Object -Unique | ForEach-Object { Write-Host "  - $_" }
        }
    }
}
