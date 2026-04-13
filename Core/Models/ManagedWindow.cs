namespace SaveTrigger.Core.Models;

/// <summary>
/// Tracks an Explorer window that was opened by this application.
/// Only windows in this registry may be closed or moved by the app.
/// </summary>
public sealed class ManagedWindow
{
    /// <summary>Win32 window handle (HWND).</summary>
    public nint Hwnd { get; set; }

    /// <summary>
    /// The folder path this window is currently showing.
    /// Mutable so it can be updated when a new tab is opened in the same window.
    /// </summary>
    public required string FolderPath { get; set; }

    /// <summary>UTC time when the app opened this window.</summary>
    public DateTime OpenedAt { get; init; }

    /// <summary>
    /// UTC time of last activation (bring-to-front). Used to determine
    /// which window to close when the MaxManagedWindows limit is reached.
    /// </summary>
    public DateTime LastActivatedAt { get; set; }
}
