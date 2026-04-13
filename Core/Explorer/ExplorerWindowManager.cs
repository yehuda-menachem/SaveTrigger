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
            Thread.Sleep(700);
            MoveToTargetMonitor(newHwnd);
            BringToFront(newHwnd);

            Thread.Sleep(400);
            MoveToTargetMonitor(newHwnd);   // second pass — overrides any late reposition
            BringToFront(newHwnd);
        }
    }

    // ── Tab support ──────────────────────────────────────────────────────────

    /// <summary>
    /// Opens <paramref name="folderPath"/> as a new background tab in the Explorer
    /// window identified by <paramref name="existingHwnd"/>, leaving the currently
    /// active tab unchanged.
    ///
    /// Strategy (Navigate2 flags are unreliable for folder tabs in practice):
    ///   1. Bring window to front (required for SendInput).
    ///   2. Ctrl+T  — opens a new empty tab and makes it active.
    ///   3. Navigate2(folderPath) — navigates the now-active new tab to the folder.
    ///   4. Ctrl+Shift+Tab — returns focus to the tab that was active before step 2.
    ///
    /// The user sees their original tab, while the new folder is available in the
    /// background tab.
    /// </summary>
    private bool TryOpenAsNewTab(nint existingHwnd, string folderPath)
    {
        try
        {
            var shellType = Type.GetTypeFromProgID("Shell.Application", throwOnError: false);
            if (shellType == null) return false;

            dynamic? shell = Activator.CreateInstance(shellType);
            if (shell == null) return false;

            try
            {
                dynamic windows = shell.Windows();
                int count = windows.Count;

                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        dynamic? item = windows.Item(i);
                        if (item == null) continue;

                        nint hwnd = (nint)(int)item.HWND;
                        if (hwnd != existingHwnd) continue;

                        // Step 1: Bring window to foreground — SendInput targets
                        // whichever window has keyboard focus.
                        BringToFront(existingHwnd);
                        Thread.Sleep(150);

                        // Step 2: Ctrl+T — new tab opens and becomes active.
                        NativeMethods.SendKeyCombo(NativeMethods.VK_T, ctrl: true);
                        Thread.Sleep(500); // wait for new tab to fully initialize

                        // Step 3: Navigate the NEW active tab to the folder.
                        //
                        // IMPORTANT: `item` still points to the OLD tab's COM object.
                        // Calling item.Navigate2 would overwrite the original tab's content.
                        // After Ctrl+T, Shell.Application reports the *currently active* tab
                        // for this HWND — re-enumerate with a fresh Shell instance to get
                        // the new tab's COM object and navigate that one instead.
                        bool navigated = false;
                        dynamic? shell2 = Activator.CreateInstance(shellType);
                        if (shell2 != null)
                        {
                            try
                            {
                                dynamic windows2 = shell2.Windows();
                                for (int j = 0, c2 = (int)windows2.Count; j < c2; j++)
                                {
                                    try
                                    {
                                        dynamic? tab = windows2.Item(j);
                                        if (tab == null) continue;
                                        if ((nint)(int)tab.HWND != existingHwnd) continue;
                                        // This is the new active tab — navigate it.
                                        tab.Navigate2(folderPath, 0x02); // navNoHistory
                                        navigated = true;
                                        break;
                                    }
                                    catch { continue; }
                                }
                            }
                            finally { Marshal.ReleaseComObject(shell2); }
                        }

                        if (!navigated)
                        {
                            _log.LogDebug("Re-enumerate for new tab failed — closing it and falling back");
                            NativeMethods.SendKeyCombo(NativeMethods.VK_W, ctrl: true); // close empty tab
                            return false;
                        }

                        Thread.Sleep(200);

                        // Step 4: Ctrl+Shift+Tab — return to the tab that was active
                        // before we opened the new one.
                        NativeMethods.SendKeyCombo(NativeMethods.VK_TAB, ctrl: true, shift: true);

                        _log.LogInformation(
                            "Opened {Folder} as background tab in window 0x{Hwnd:X}",
                            folderPath, existingHwnd);
                        return true;
                    }
                    catch { continue; }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(shell);
            }
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
        Thread.Sleep(150); // let Explorer begin initializing
    }

    private static nint FindNewExplorerHwnd(HashSet<nint> before, CancellationToken ct)
    {
        for (int i = 0; i < 10; i++)   // up to ~2.5 s
        {
            ct.ThrowIfCancellationRequested();
            Thread.Sleep(250);

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
