using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;
using SaveTrigger.Core.Models;
using SaveTrigger.Interop;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;

namespace SaveTrigger.Core.Explorer;

/// <summary>
/// Stage 5: Opens and manages Windows Explorer windows on behalf of the app.
///
/// Ownership model:
///   Only windows registered in _windows are considered "managed". The app will
///   NEVER close, move, or reuse an Explorer window it did not open itself.
///
/// Window lifecycle:
///   - Open a new window via "explorer.exe /select,{file}"
///   - Detect the new HWND by comparing before/after CabinetWClass enumeration
///   - Track it in _windows with its folder path and timestamps
///   - Move to the target monitor and bring to front
///   - On next event for the same folder: reuse (bring to front)
///   - If MaxManagedWindows exceeded: close the oldest managed window
///
/// Thread safety:
///   All mutations to _windows are guarded by _lock. Methods are called from
///   the DebounceService thread pool, so locking is required.
/// </summary>
public sealed class ExplorerWindowManager
{
    private readonly ExplorerTabHelper _tabHelper;
    private readonly AppSettings _settings;
    private readonly ILogger<ExplorerWindowManager> _log;

    // Map of HWND → ManagedWindow for windows opened by this app.
    private readonly Dictionary<nint, ManagedWindow> _windows = [];
    private readonly object _lock = new();

    public ExplorerWindowManager(
        ExplorerTabHelper tabHelper,
        IOptions<AppSettings> settings,
        ILogger<ExplorerWindowManager> log)
    {
        _tabHelper = tabHelper;
        _settings  = settings.Value;
        _log       = log;
    }

    /// <summary>
    /// Opens the folder containing the newly created file in an Explorer window,
    /// with the file selected. Reuses an existing managed window if available.
    /// </summary>
    public Task OpenFolderAsync(string folderPath, string filePath, CancellationToken ct)
    {
        // Run on a background thread — Thread.Sleep calls inside must not block the
        // DebounceService execution context.
        return Task.Run(() => OpenFolderCore(folderPath, filePath, ct), ct);
    }

    // ── Core logic ───────────────────────────────────────────────────────────

    private void OpenFolderCore(string folderPath, string filePath, CancellationToken ct)
    {
        lock (_lock)
        {
            ct.ThrowIfCancellationRequested();

            PruneDeadWindows();

            // ── Check for existing managed window for this folder ────────────
            var existing = _windows.Values.FirstOrDefault(w =>
                string.Equals(w.FolderPath, folderPath, StringComparison.OrdinalIgnoreCase) &&
                NativeMethods.IsWindow(w.Hwnd));

            if (existing != null)
            {
                _log.LogInformation(
                    "Reusing managed window 0x{Hwnd:X} for {Folder}", existing.Hwnd, folderPath);

                existing.LastActivatedAt = DateTime.UtcNow;
                BringToFront(existing.Hwnd);
                MoveToTargetMonitor(existing.Hwnd);
                return;
            }

            // ── Enforce window limit ─────────────────────────────────────────
            if (_windows.Count >= _settings.MaxManagedWindows)
            {
                var oldest = _windows.Values.OrderBy(w => w.LastActivatedAt).First();
                _log.LogInformation(
                    "Window limit ({Max}) reached — closing oldest managed window 0x{Hwnd:X} ({Folder})",
                    _settings.MaxManagedWindows, oldest.Hwnd, oldest.FolderPath);

                NativeMethods.PostMessage(oldest.Hwnd, NativeMethods.WM_CLOSE, 0, 0);
                _windows.Remove(oldest.Hwnd);
            }

            // ── Open a new Explorer window ───────────────────────────────────
            var before = EnumerateExplorerWindows();
            LaunchExplorerWithSelection(filePath);

            // Poll for the new Explorer window to appear (up to ~2 seconds).
            nint newHwnd = FindNewExplorerHwnd(before, ct);

            if (newHwnd == 0)
            {
                _log.LogError(
                    "Could not detect new Explorer window after opening {Folder} — " +
                    "folder was opened but window is not tracked", folderPath);
                return;
            }

            _log.LogInformation(
                "Opened Explorer window 0x{Hwnd:X} for {Folder}", newHwnd, folderPath);

            _windows[newHwnd] = new ManagedWindow
            {
                Hwnd            = newHwnd,
                FolderPath      = folderPath,
                OpenedAt        = DateTime.UtcNow,
                LastActivatedAt = DateTime.UtcNow
            };

            MoveToTargetMonitor(newHwnd);
            BringToFront(newHwnd);
        }
    }

    // ── Window operations ────────────────────────────────────────────────────

    private static void LaunchExplorerWithSelection(string filePath)
    {
        // "explorer.exe /select,<path>" opens the folder with the file selected.
        // If Explorer is already running, Windows routes the request to the existing
        // process — the returned Process object may have a different PID.
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName         = "explorer.exe",
                Arguments        = $"/select,\"{filePath}\"",
                UseShellExecute  = false
            };
            Process.Start(psi);

            // Brief wait for Explorer to begin initializing before we poll for HWNDs.
            Thread.Sleep(150);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to launch Explorer for {filePath}", ex);
        }
    }

    private static nint FindNewExplorerHwnd(HashSet<nint> before, CancellationToken ct)
    {
        // Poll for up to ~2 seconds (8 × 250 ms).
        for (int i = 0; i < 8; i++)
        {
            ct.ThrowIfCancellationRequested();
            Thread.Sleep(250);

            var current = EnumerateExplorerWindows();
            var newWindows = current.Except(before).ToList();
            if (newWindows.Count > 0)
                return newWindows[0];
        }
        return 0;
    }

    private static HashSet<nint> EnumerateExplorerWindows()
    {
        var result = new HashSet<nint>();
        var classBuf = new StringBuilder(64);

        NativeMethods.EnumWindows((hwnd, _) =>
        {
            classBuf.Clear();
            NativeMethods.GetClassName(hwnd, classBuf, classBuf.Capacity);
            var cls = classBuf.ToString();

            // CabinetWClass = folder view, ExploreWClass = legacy Explorer with tree pane
            if (cls is "CabinetWClass" or "ExploreWClass")
                result.Add(hwnd);

            return true; // continue enumeration
        }, 0);

        return result;
    }

    private void MoveToTargetMonitor(nint hwnd)
    {
        try
        {
            var screens = Screen.AllScreens;
            var idx = Math.Clamp(_settings.TargetMonitorIndex, 0, screens.Length - 1);
            var area = screens[idx].WorkingArea;

            _log.LogDebug(
                "Moving window 0x{Hwnd:X} to monitor {Idx} ({Left},{Top} {W}×{H})",
                hwnd, idx, area.Left, area.Top, area.Width, area.Height);

            NativeMethods.SetWindowPos(
                hwnd,
                NativeMethods.HWND_TOP,
                area.Left, area.Top,
                area.Width, area.Height,
                NativeMethods.SWP_SHOWWINDOW);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Could not move window 0x{Hwnd:X} to target monitor", hwnd);
        }
    }

    private static void BringToFront(nint hwnd)
    {
        NativeMethods.ShowWindow(hwnd, NativeMethods.SW_RESTORE);
        NativeMethods.SetForegroundWindow(hwnd);
    }

    private void PruneDeadWindows()
    {
        var dead = _windows.Keys
            .Where(h => !NativeMethods.IsWindow(h))
            .ToList();

        foreach (var h in dead)
        {
            _log.LogDebug("Pruning dead managed window 0x{Hwnd:X}", h);
            _windows.Remove(h);
        }
    }
}
