using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;
using SaveTrigger.Core.Models;
using SaveTrigger.Interop;

namespace SaveTrigger.Core.Correlation;

/// <summary>
/// Stage 4 of the pipeline: decides whether a file event should trigger an
/// Explorer window opening. Accepts only files that were plausibly created by
/// the current user on this machine.
///
/// Decision logic:
///  1. Reject if the path is on a remote/network drive.
///  2. Reject if there is no recent foreground window activity.
///  3. Accept if the file's creation timestamp falls within the recent activity window.
///  4. Accept if a save-related window was recently in the foreground.
///  5. Default: reject.
///
/// Every decision (accept or reject) is logged with a reason so the debug log
/// can be used to tune thresholds without code changes.
/// </summary>
public sealed class LocalOriginCorrelator
{
    private readonly ActivityTracker _activityTracker;
    private readonly AppSettings _settings;
    private readonly ILogger<LocalOriginCorrelator> _log;

    // Window class names associated with Windows Save/Open dialog boxes.
    private static readonly HashSet<string> SaveDialogClasses = new(StringComparer.OrdinalIgnoreCase)
    {
        "#32770",           // Standard Win32 common dialog (Save As, Open)
        "ToolbarWindow32",  // Toolbar within common dialogs
        "DirectUIHWND",     // Modern shell picker content area
        "SHELLDLL_DefView", // Shell browser inside dialog
        "SysListView32",    // File list inside classic dialog
    };

    // Title substrings associated with save operations (English).
    private static readonly string[] SaveTitleKeywords =
    [
        "Save", "Export", "Download", "Speichern", "Enregistrer",
        "Guardar", "Salva", "Save As", "Save File", "Choose a file name",
    ];

    public LocalOriginCorrelator(
        ActivityTracker activityTracker,
        IOptions<AppSettings> settings,
        ILogger<LocalOriginCorrelator> log)
    {
        _activityTracker = activityTracker;
        _settings        = settings.Value;
        _log             = log;
    }

    /// <summary>
    /// Evaluates a file event and returns whether it should trigger an Explorer open.
    /// </summary>
    public CorrelationResult Evaluate(FileEvent evt)
    {
        // ── 1. Network drive check ───────────────────────────────────────────
        var root = GetDriveRoot(evt.FullPath);
        if (root != null)
        {
            var driveType = NativeMethods.GetDriveType(root);
            if (driveType == NativeMethods.DRIVE_REMOTE)
            {
                return Reject(evt, [], $"network/remote drive (root={root}, driveType={driveType})");
            }
        }

        // ── 2. Activity window check ─────────────────────────────────────────
        var recent = _activityTracker.GetRecentActivity();
        if (recent.Count == 0)
        {
            return Reject(evt, recent,
                $"no foreground activity in last {_settings.ActivityWindowSeconds}s");
        }

        // ── 3. Timestamp correlation ─────────────────────────────────────────
        DateTime fileCreatedUtc;
        try
        {
            fileCreatedUtc = File.GetCreationTimeUtc(evt.FullPath);
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "Could not read creation time for {Path}", evt.FullPath);
            fileCreatedUtc = evt.DetectedAt; // fall back to detection time
        }

        var lastActivity = recent.Max(s => s.CapturedAt);
        var firstActivity = recent.Min(s => s.CapturedAt);

        // Accept if file was created during or shortly after the activity window.
        // "Shortly after" covers the case where the user finished saving before we
        // captured the foreground change.
        var deltaSeconds = (fileCreatedUtc - lastActivity).TotalSeconds;

        if (fileCreatedUtc >= firstActivity.AddSeconds(-2) &&
            fileCreatedUtc <= lastActivity.AddSeconds(_settings.ActivityWindowSeconds))
        {
            return Accept(evt, recent, fileCreatedUtc,
                $"file creation at {fileCreatedUtc:HH:mm:ss.fff} within activity window " +
                $"[{firstActivity:HH:mm:ss}–{lastActivity:HH:mm:ss}]");
        }

        // ── 4. Save dialog heuristic ─────────────────────────────────────────
        // Even if the timestamp is slightly outside the window, if a save dialog
        // was recently visible we give benefit of the doubt.
        var hasSaveDialog = recent.Any(s =>
            SaveDialogClasses.Contains(s.ClassName) ||
            SaveTitleKeywords.Any(kw =>
                s.WindowTitle.Contains(kw, StringComparison.OrdinalIgnoreCase)));

        if (hasSaveDialog && Math.Abs(deltaSeconds) <= _settings.ActivityWindowSeconds * 2)
        {
            return Accept(evt, recent, fileCreatedUtc,
                $"save dialog detected in recent activity (delta={deltaSeconds:F1}s)");
        }

        // ── 5. Default: reject ───────────────────────────────────────────────
        return Reject(evt, recent,
            $"file created at {fileCreatedUtc:HH:mm:ss.fff}, last activity {lastActivity:HH:mm:ss.fff}, " +
            $"delta={deltaSeconds:F1}s — outside correlation window");
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private CorrelationResult Accept(
        FileEvent evt,
        IReadOnlyList<ForegroundSnapshot> recent,
        DateTime createdAt,
        string reason)
    {
        _log.LogDebug("ACCEPTED {File} — {Reason}", evt.FullPath, reason);
        return new CorrelationResult(true, reason, createdAt, recent);
    }

    private CorrelationResult Reject(
        FileEvent evt,
        IReadOnlyList<ForegroundSnapshot> recent,
        string reason)
    {
        _log.LogDebug("REJECTED {File} — {Reason}", evt.FullPath, reason);
        return new CorrelationResult(
            false, reason,
            DateTime.MinValue, // creation time not meaningful for rejections
            recent);
    }

    /// <summary>
    /// Returns the drive root suitable for GetDriveType:
    ///   - Local drive:  "C:\"
    ///   - UNC path:     "\\server\share\"  (GetDriveType handles UNC directly)
    ///   - Unknown:      null
    /// </summary>
    private static string? GetDriveRoot(string path)
    {
        try { return Path.GetPathRoot(path); }
        catch { return null; }
    }
}
