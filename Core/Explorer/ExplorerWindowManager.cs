using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;
using SaveTrigger.Core.Models;
using SaveTrigger.Interop;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SaveTrigger.Core.Explorer;

/// <summary>
/// Stage 5: Opens and manages Windows Explorer windows on behalf of the app.
///
/// Ownership model:
///   Only windows in _windows are "managed" — the app never closes, moves,
///   or reads windows it did not open.
///
/// Behaviour for a new file event:
///   1. Folder already open in a managed window  → bring to front (same tab, no action).
///   2. Any other managed window exists          → open folder as new tab in that window.
///   3. No managed windows (or tab open failed)  → open a brand-new Explorer window.
///
/// Bug fixes applied in this version:
///   A. Foreground: AttachThreadInput trick so SetForegroundWindow works from
///      a background tray process.
///   B. Monitor: double-move with delay so our SetWindowPos fires AFTER Explorer
///      reads its saved position from the registry and repositions itself.
///   C. Tabs: TryOpenAsNewTab uses Shell.Application Navigate2 with the
///      navOpenInNewTab flag (0x800) — opens folder in a new tab rather than
///      a new top-level window.
/// </summary>
public sealed class ExplorerWindowManager
{
    private readonly AppSettings _settings;
    private readonly ILogger<ExplorerWindowManager> _log;

    private readonly Dictionary<nint, ManagedWindow> _windows = [];
    private readonly object _lock = new();

    public ExplorerWindowManager(
        ExplorerTabHelper tabHelper,
        IOptions<AppSettings> settings,
        ILogger<ExplorerWindowManager> log)
    {
        _settings = settings.Value;
        _log      = log;
    }

    public Task OpenFolderAsync(string folderPath, string filePath, CancellationToken ct)
        => Task.Run(() => OpenFolderCore(folderPath, filePath, ct), ct);

    // ── Core logic ───────────────────────────────────────────────────────────

    private void OpenFolderCore(string folderPath, string filePath, CancellationToken ct)
    {
        lock (_lock)
        {
            ct.ThrowIfCancellationRequested();
            PruneDeadWindows();

            // ── 1. Folder already open in a managed window ───────────────────
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

            // ── 2. Try to open as a new tab in an existing managed window ────
            var anyWindow = _windows.Values.FirstOrDefault(w => NativeMethods.IsWindow(w.Hwnd));
            if (anyWindow != null)
            {
                bool tabbed = TryOpenAsNewTab(anyWindow.Hwnd, folderPath);
                if (tabbed)
                {
                    // The top-level HWND stays the same; update tracked folder.
                    anyWindow.FolderPath      = folderPath;
                    anyWindow.LastActivatedAt = DateTime.UtcNow;

                    // Give the tab time to navigate before we move the window.
                    Thread.Sleep(400);
                    MoveToTargetMonitor(anyWindow.Hwnd);
                    BringToFront(anyWindow.Hwnd);
                    return;
                }
                _log.LogDebug("Tab open failed — will open a new window for {Folder}", folderPath);
            }

            // ── 3. Enforce window limit ──────────────────────────────────────
            if (_windows.Count >= _settings.MaxManagedWindows)
            {
                var oldest = _windows.Values.OrderBy(w => w.LastActivatedAt).First();
                _log.LogInformation(
                    "Window limit ({Max}) reached — closing oldest 0x{Hwnd:X} ({Folder})",
                    _settings.MaxManagedWindows, oldest.Hwnd, oldest.FolderPath);
                NativeMethods.PostMessage(oldest.Hwnd, NativeMethods.WM_CLOSE, 0, 0);
                _windows.Remove(oldest.Hwnd);
            }

            // ── 4. Open a brand-new Explorer window ──────────────────────────
            var before = EnumerateExplorerWindows();
            LaunchExplorerWithSelection(filePath);

            nint newHwnd = FindNewExplorerHwnd(before, ct);
            if (newHwnd == 0)
            {
                _log.LogError(
                    "Could not detect new Explorer window for {Folder} — " +
                    "folder opened but not tracked", folderPath);
                return;
            }

            _log.LogInformation("Opened Explorer window 0x{Hwnd:X} for {Folder}", newHwnd, folderPath);
            _windows[newHwnd] = new ManagedWindow
            {
                Hwnd            = newHwnd,
                FolderPath      = folderPath,
                OpenedAt        = DateTime.UtcNow,
                LastActivatedAt = DateTime.UtcNow
            };

            // FIX B: Explorer reads its saved position from the registry and
            // repositions itself after the window is first shown. We must wait
            // for that to complete before applying our target monitor position,
            // then repeat once more in case of a late second reposition.
            Thread.Sleep(300);
            MoveToTargetMonitor(newHwnd);
            BringToFront(newHwnd);

            Thread.Sleep(200);
            MoveToTargetMonitor(newHwnd);   // second pass — overrides any late reposition
            BringToFront(newHwnd);
        }
    }

    // ── Tab support ──────────────────────────────────────────────────────────

    /// <summary>
    /// Opens <paramref name="folderPath"/> as a new active tab in the Explorer
    /// window identified by <paramref name="existingHwnd"/>.
    ///
    /// Strategy:
    ///   1. Snapshot LocationURLs of all existing tabs in this window (before).
    ///   2. Bring window to front + Ctrl+T → new "Home" tab opens and becomes active.
    ///   3. Re-enumerate tabs; identify the new one by URL exclusion from the snapshot.
    ///   4. Navigate2(folderPath) on the new tab → shows the target folder.
    ///   5. Leave focus on the new tab (no Ctrl+Shift+Tab back).
    /// </summary>
    private bool TryOpenAsNewTab(nint existingHwnd, string folderPath)
    {
        try
        {
            var shellType = Type.GetTypeFromProgID("Shell.Application", throwOnError: false);
            if (shellType == null) return false;

            // ── Snapshot BEFORE Ctrl+T ───────────────────────────────────────
            var urlsBefore = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int countBefore = 0;

            dynamic? shell1 = Activator.CreateInstance(shellType);
            if (shell1 == null) return false;
            try
            {
                dynamic wins1 = shell1.Windows();
                for (int i = 0, c = (int)wins1.Count; i < c; i++)
                {
                    try
                    {
                        dynamic? item = wins1.Item(i);
                        if (item == null) continue;
                        if ((nint)(int)item.HWND != existingHwnd) continue;
                        countBefore++;
                        string? url = item.LocationURL;
                        if (!string.IsNullOrEmpty(url)) urlsBefore.Add(url);
                    }
                    catch { continue; }
                }
            }
            finally { Marshal.ReleaseComObject(shell1); }

            _log.LogDebug("Tab snapshot before Ctrl+T: {Count} tabs in 0x{Hwnd:X}", countBefore, existingHwnd);

            // ── Step 1+2: Focus + Ctrl+T ─────────────────────────────────────
            BringToFront(existingHwnd);
            Thread.Sleep(80);
            NativeMethods.SendKeyCombo(NativeMethods.VK_T, ctrl: true);
            Thread.Sleep(200);

            // ── Step 3+4: Identify new tab and navigate it ───────────────────
            // The new tab is the one whose LocationURL was NOT in the before-snapshot.
            bool navigated = false;
            dynamic? shell2 = Activator.CreateInstance(shellType);
            if (shell2 != null)
            {
                try
                {
                    dynamic wins2 = shell2.Windows();
                    int countAfter = 0;
                    dynamic? newTab = null;
                    string? newTabUrl = null;

                    for (int j = 0, c2 = (int)wins2.Count; j < c2; j++)
                    {
                        try
                        {
                            dynamic? tab = wins2.Item(j);
                            if (tab == null) continue;
                            if ((nint)(int)tab.HWND != existingHwnd) continue;
                            countAfter++;

                            string? url = tab.LocationURL;
                            if (string.IsNullOrEmpty(url) || !urlsBefore.Contains(url))
                            {
                                newTab    = tab;
                                newTabUrl = url;
                            }
                        }
                        catch { continue; }
                    }

                    _log.LogDebug(
                        "Tab snapshot after Ctrl+T: {After} tabs (expected {Expected}), new URL={Url}",
                        countAfter, countBefore + 1, newTabUrl ?? "<null>");

                    if (countAfter <= countBefore)
                    {
                        _log.LogDebug("Ctrl+T was not registered (window lost focus?) — falling back");
                        return false;
                    }

                    if (newTab != null)
                    {
                        try
                        {
                            newTab.Navigate2(folderPath, 0);
                            navigated = true;
                        }
                        catch (Exception ex)
                        {
                            _log.LogDebug(ex, "Navigate2 failed on new tab");
                        }
                    }
                }
                finally { Marshal.ReleaseComObject(shell2); }
            }

            if (!navigated)
            {
                _log.LogDebug("Re-enumerate for new tab failed — closing it and falling back");
                NativeMethods.SendKeyCombo(NativeMethods.VK_W, ctrl: true);
                return false;
            }

            Thread.Sleep(100);
            // Focus stays on the new tab — no Ctrl+Shift+Tab back.

            _log.LogInformation("Opened {Folder} as new tab in window 0x{Hwnd:X}", folderPath, existingHwnd);
            return true;
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "TryOpenAsNewTab failed for 0x{Hwnd:X}", existingHwnd);
        }
        return false;
    }

    // ── Explorer launch & HWND detection ────────────────────────────────────

    private static void LaunchExplorerWithSelection(string filePath)
    {
        var psi = new ProcessStartInfo
        {
            FileName        = "explorer.exe",
            Arguments       = $"/select,\"{filePath}\"",
            UseShellExecute = false
        };
        Process.Start(psi);
        Thread.Sleep(50); // let Explorer begin initializing
    }

    private static nint FindNewExplorerHwnd(HashSet<nint> before, CancellationToken ct)
    {
        for (int i = 0; i < 20; i++)   // up to ~1 s
        {
            ct.ThrowIfCancellationRequested();
            Thread.Sleep(50);

            var current = EnumerateExplorerWindows();
            var newOnes = current.Except(before).ToList();
            if (newOnes.Count > 0) return newOnes[0];
        }
        return 0;
    }

    private static HashSet<nint> EnumerateExplorerWindows()
    {
        var result   = new HashSet<nint>();
        var classBuf = new StringBuilder(64);

        NativeMethods.EnumWindows((hwnd, _) =>
        {
            classBuf.Clear();
            NativeMethods.GetClassName(hwnd, classBuf, classBuf.Capacity);
            if (classBuf.ToString() is "CabinetWClass" or "ExploreWClass")
                result.Add(hwnd);
            return true;
        }, 0);

        return result;
    }

    // ── Window positioning ───────────────────────────────────────────────────

    private void MoveToTargetMonitor(nint hwnd)
    {
        try
        {
            var screens = Screen.AllScreens;
            var idx     = Math.Clamp(_settings.TargetMonitorIndex, 0, screens.Length - 1);
            var area    = screens[idx].WorkingArea;

            _log.LogDebug(
                "SetWindowPos 0x{Hwnd:X} → monitor {Idx} ({L},{T} {W}×{H})",
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
            _log.LogWarning(ex, "MoveToTargetMonitor failed for 0x{Hwnd:X}", hwnd);
        }
    }

    /// <summary>
    /// FIX A: Brings window to foreground reliably from a background process.
    ///
    /// Windows restricts SetForegroundWindow to processes that currently have
    /// foreground rights.  Temporarily attaching our thread's input queue to the
    /// foreground thread's queue grants those rights for the duration of the call.
    /// </summary>
    private void BringToFront(nint hwnd)
    {
        uint foreHwnd     = (uint)NativeMethods.GetWindowThreadProcessId(
                                NativeMethods.GetForegroundWindow(), out _);
        uint ourThreadId  = NativeMethods.GetCurrentThreadId();

        bool attached = false;
        if (foreHwnd != 0 && foreHwnd != ourThreadId)
            attached = NativeMethods.AttachThreadInput(ourThreadId, foreHwnd, true);

        try
        {
            NativeMethods.ShowWindow(hwnd, NativeMethods.SW_RESTORE);
            NativeMethods.BringWindowToTop(hwnd);
            NativeMethods.SetForegroundWindow(hwnd);
        }
        finally
        {
            if (attached)
                NativeMethods.AttachThreadInput(ourThreadId, foreHwnd, false);
        }
    }

    // ── Housekeeping ─────────────────────────────────────────────────────────

    private void PruneDeadWindows()
    {
        var dead = _windows.Keys.Where(h => !NativeMethods.IsWindow(h)).ToList();
        foreach (var h in dead)
        {
            _log.LogDebug("Pruning dead window 0x{Hwnd:X}", h);
            _windows.Remove(h);
        }
    }
}
