using System.IO;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;
using SaveTrigger.Core.Models;
using SaveTrigger.Tray;
using System.Threading.Channels;

namespace SaveTrigger.Core.Pipeline;

/// <summary>
/// Stage 1: Creates one FileSystemWatcher per configured monitored folder and
/// feeds raw file events into the shared Channel for downstream processing.
///
/// Filtering at this stage:
///   - Temp file extensions (.tmp, .crdownload, etc.)
///   - Filename prefixes (~$, .~ for Office/lock files)
///   - Pause state (via AppState.IsPaused)
///
/// Both Created and Renamed events are forwarded — many editors write to a temp
/// file and then rename it to the final name (e.g., Notepad++, Visual Studio).
/// </summary>
public sealed class FileWatcherService : IHostedService, IDisposable
{
    private readonly Channel<FileEvent> _channel;
    private readonly AppSettings _settings;
    private readonly AppState _appState;
    private readonly ILogger<FileWatcherService> _log;
    private readonly List<FileSystemWatcher> _watchers = [];

    public FileWatcherService(
        Channel<FileEvent> channel,
        IOptions<AppSettings> settings,
        AppState appState,
        ILogger<FileWatcherService> log)
    {
        _channel  = channel;
        _settings = settings.Value;
        _appState = appState;
        _log      = log;
    }

    // ── IHostedService ───────────────────────────────────────────────────────

    public Task StartAsync(CancellationToken cancellationToken)
    {
        if (_settings.MonitoredFolders.Count == 0)
        {
            _log.LogWarning("No monitored folders configured — FileWatcherService is idle");
            return Task.CompletedTask;
        }

        foreach (var folder in _settings.MonitoredFolders)
        {
            if (!Directory.Exists(folder))
            {
                _log.LogWarning("Monitored folder does not exist and will be skipped: {Folder}", folder);
                continue;
            }

            try
            {
                var watcher = new FileSystemWatcher(folder)
                {
                    IncludeSubdirectories = true,
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                    // Max buffer: 64 KB — reduces event loss on high-traffic directories.
                    InternalBufferSize = 65536,
                    EnableRaisingEvents = true
                };

                watcher.Created += OnCreated;
                watcher.Renamed += OnRenamed;
                watcher.Error   += OnError;

                _watchers.Add(watcher);
                _log.LogInformation("Watching {Folder} (recursive)", folder);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Failed to create watcher for {Folder}", folder);
            }
        }

        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        foreach (var w in _watchers)
            w.EnableRaisingEvents = false;

        // Signal the channel writer that no more items will be written.
        _channel.Writer.TryComplete();

        _log.LogInformation("FileWatcherService stopped ({N} watchers)", _watchers.Count);
        return Task.CompletedTask;
    }

    // ── Event handlers ───────────────────────────────────────────────────────

    private void OnCreated(object sender, FileSystemEventArgs e)
    {
        if (_appState.IsPaused) return;

        var watcher = (FileSystemWatcher)sender;
        EnqueueIfRelevant(e.FullPath, WatcherChangeTypes.Created, watcher.Path);
    }

    private void OnRenamed(object sender, RenamedEventArgs e)
    {
        if (_appState.IsPaused) return;

        // The *new* name is the final file — treat rename-to as a creation.
        // This handles the common editor pattern: write to .tmp then rename.
        var watcher = (FileSystemWatcher)sender;
        EnqueueIfRelevant(e.FullPath, WatcherChangeTypes.Renamed, watcher.Path);
    }

    private void OnError(object sender, ErrorEventArgs e)
    {
        var ex = e.GetException();
        _log.LogWarning(ex, "FileSystemWatcher error (possible buffer overflow or path issue)");

        // Re-enable raising events if the watcher is still alive.
        if (sender is FileSystemWatcher { EnableRaisingEvents: false } w)
        {
            try { w.EnableRaisingEvents = true; }
            catch (Exception rex)
            {
                _log.LogError(rex, "Could not re-enable watcher after error");
            }
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void EnqueueIfRelevant(string fullPath, WatcherChangeTypes changeType, string watcherRoot)
    {
        var name = Path.GetFileName(fullPath);

        if (string.IsNullOrEmpty(name))
            return;

        // Skip temp file extensions
        var ext = Path.GetExtension(name);
        if (!string.IsNullOrEmpty(ext) &&
            _settings.TempFileExtensions.Any(t =>
                ext.Equals(t, StringComparison.OrdinalIgnoreCase)))
        {
            _log.LogDebug("Ignoring temp-extension file: {File}", fullPath);
            return;
        }

        // Skip known lock/temp prefixes
        if (_settings.IgnorePrefixes.Any(p =>
                name.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
        {
            _log.LogDebug("Ignoring prefix-filtered file: {File}", fullPath);
            return;
        }

        var dir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(dir)) return;

        var evt = new FileEvent(fullPath, dir, changeType, DateTime.UtcNow, watcherRoot);

        // TryWrite is non-blocking. If the channel is full, the oldest item is dropped
        // (BoundedChannelFullMode.DropOldest) — logged at debug level.
        if (!_channel.Writer.TryWrite(evt))
        {
            _log.LogDebug("Channel full — dropped event for {File}", fullPath);
        }
        else
        {
            _log.LogDebug("Queued {ChangeType} event: {File}", changeType, fullPath);
        }
    }

    public void Dispose()
    {
        foreach (var w in _watchers) w.Dispose();
        _watchers.Clear();
    }
}
