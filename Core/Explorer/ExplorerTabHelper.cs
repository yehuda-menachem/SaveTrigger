using Microsoft.Extensions.Logging;
using System.Runtime.InteropServices;
using SaveTrigger.Interop;

namespace SaveTrigger.Core.Explorer;

/// <summary>
/// Uses Shell COM (Shell.Application / IShellWindows) to query which Explorer
/// windows are currently open and what folder each one is showing.
///
/// This is used by ExplorerWindowManager to:
///   a) Detect if a target folder is already open in a managed window.
///   b) Optionally adopt an existing Explorer window instead of opening a new one.
///
/// The implementation uses dynamic COM dispatch (via 'dynamic') on the
/// IShellWindows collection to avoid implementing the full IWebBrowser2 vtable.
/// All COM calls are wrapped in try/catch so a failure doesn't break the pipeline.
/// </summary>
public sealed class ExplorerTabHelper
{
    private readonly ILogger<ExplorerTabHelper> _log;

    public ExplorerTabHelper(ILogger<ExplorerTabHelper> log)
    {
        _log = log;
    }

    /// <summary>
    /// Searches currently open Explorer windows for one showing the given folder.
    /// Returns the HWND if found, or <c>0</c> if not found or on any COM error.
    /// </summary>
    public nint GetExplorerWindowForFolder(string folderPath)
    {
        try
        {
            return SearchShellWindows(folderPath);
        }
        catch (COMException ex)
        {
            _log.LogDebug(ex, "COM error while searching shell windows for {Folder}", folderPath);
            return 0;
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "Unexpected error searching shell windows for {Folder}", folderPath);
            return 0;
        }
    }

    private nint SearchShellWindows(string folderPath)
    {
        // Create Shell.Application COM object via late binding.
        // ProgID "Shell.Application" → CLSID 13709620-C279-11CE-A49E-444553540000
        var shellAppType = Type.GetTypeFromProgID("Shell.Application", throwOnError: false);
        if (shellAppType == null)
        {
            _log.LogDebug("Shell.Application ProgID not available");
            return 0;
        }

        object? shell = null;
        try
        {
            shell = Activator.CreateInstance(shellAppType);
            if (shell == null) return 0;

            // shell.Windows() returns an IShellWindows collection
            dynamic shellDynamic = shell;
            dynamic windows = shellDynamic.Windows();
            int count = windows.Count;

            var normalizedTarget = NormalizePath(folderPath);

            for (int i = 0; i < count; i++)
            {
                try
                {
                    dynamic? item = windows.Item(i);
                    if (item == null) continue;

                    // LocationURL looks like: "file:///C:/Users/username/Downloads/"
                    string? locationUrl = item.LocationURL;
                    if (string.IsNullOrEmpty(locationUrl)) continue;

                    var localPath = DecodeFileUrl(locationUrl);
                    if (localPath == null) continue;

                    if (string.Equals(NormalizePath(localPath), normalizedTarget,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        // HWND comes back as int from IDispatch; cast carefully.
                        nint hwnd = (nint)(int)item.HWND;
                        _log.LogDebug(
                            "Found existing Explorer window HWND=0x{Hwnd:X} for {Folder}",
                            hwnd, folderPath);
                        return hwnd;
                    }
                }
                catch
                {
                    // Individual item access may fail (e.g., window closed mid-enumeration).
                    continue;
                }
            }
        }
        finally
        {
            if (shell != null)
                Marshal.ReleaseComObject(shell);
        }

        return 0;
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    /// <summary>Decodes a file:/// URL to a local path, or returns null on failure.</summary>
    private static string? DecodeFileUrl(string url)
    {
        try
        {
            if (!url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                return null;

            // Uri.LocalPath handles URL decoding and path separator normalization.
            var uri = new Uri(url);
            return uri.LocalPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        catch
        {
            return null;
        }
    }

    private static string NormalizePath(string path) =>
        path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
}
