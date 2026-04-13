using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;
using SaveTrigger.Core.Correlation;
using SaveTrigger.Core.Explorer;
using SaveTrigger.Core.Models;
using System.Threading.Channels;

namespace SaveTrigger.Core.Pipeline;

/// <summary>
/// Stages 2–5 of the pipeline, running as a single hosted background service.
///
/// Stage 2 — Debounce:
///   Groups events by directory. Only processes a group after DebounceMs have
///   elapsed since the last event in that directory. This collapses "save storms"
///   (many events for one logical save) into a single action.
///
/// Stage 3 — File Stabilization:
///   Delegates to FileStabilizer to wait until the file is fully written.
///
/// Stage 4 — Correlation:
///   Delegates to LocalOriginCorrelator to decide if the event was local.
///
/// Stage 5 — Action:
///   Delegates to ExplorerWindowManager to open/reuse the correct Explorer window.
/// </summary>
public sealed class DebounceService : BackgroundService
{
    private readonly Channel<FileEvent> _channel;
    private readonly FileStabilizer _stabilizer;
    private readonly LocalOriginCorrelator _correlator;
    private readonly ExplorerWindowManager _explorer;
    private readonly AppSettings _settings;
    private readonly ILogger<DebounceService> _log;

    // Keyed by DirectoryPath (case-insensitive on Windows).
    private readonly Dictionary<string, DebounceGroup> _groups =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly object _groupLock = new();

    // In-flight ProcessGroupAsync tasks — tracked so we can await them on shutdown.
    private readonly List<Task> _inflight = [];
    private readonly object _inflightLock = new();

    public DebounceService(
        Channel<FileEvent> channel,
        FileStabilizer stabilizer,
        LocalOriginCorrelator correlator,
        ExplorerWindowManager explorer,
        IOptions<AppSettings> settings,
        ILogger<DebounceService> log)
    {
        _channel    = channel;
        _stabilizer = stabilizer;
        _correlator = correlator;
        _explorer   = explorer;
        _settings   = settings.Value;
        _log        = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _log.LogInformation("DebounceService started (debounce={DebounceMs}ms)", _settings.DebounceMs);

        // Timer fires every 200 ms to flush expired debounce groups.
        using var flushTimer = new System.Threading.Timer(
            _ => FlushExpiredGroups(stoppingToken),
            null,
            dueTime: TimeSpan.FromMilliseconds(200),
            period: TimeSpan.FromMilliseconds(200));

        try
        {
            // Reads events from the channel until it is completed (FileWatcherService stopped)
            // or the cancellation token fires.
            await foreach (var evt in _channel.Reader.ReadAllAsync(stoppingToken))
            {
                AddToGroup(evt);
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown path.
        }

        // Flush any remaining groups on shutdown (best-effort).
        _log.LogDebug("DebounceService: channel drained, flushing remaining groups");
        List<DebounceGroup> remaining;
        lock (_groupLock)
        {
            remaining = [.. _groups.Values];
            _groups.Clear();
        }

        var finalTasks = remaining
            .Select(g => ProcessGroupAsync(g, CancellationToken.None))
            .ToList();

        // Also wait for any in-flight tasks from the timer.
        List<Task> inFlight;
        lock (_inflightLock)
            inFlight = [.. _inflight];

        await Task.WhenAll([.. finalTasks, .. inFlight]);

        _log.LogInformation("DebounceService stopped");
    }

    // ── Stage 2: Debounce ────────────────────────────────────────────────────

    private void AddToGroup(FileEvent evt)
    {
        lock (_groupLock)
        {
            if (_groups.TryGetValue(evt.DirectoryPath, out var existing))
            {
                existing.Events.Add(evt);
                existing.LastEventAt = DateTime.UtcNow;

                // Prefer the most recent non-temp-looking file as primary.
                if (existing.PrimaryEvent == null ||
                    evt.DetectedAt > existing.PrimaryEvent.DetectedAt)
                {
                    existing.PrimaryEvent = evt;
                }
            }
            else
            {
                _groups[evt.DirectoryPath] = new DebounceGroup
                {
                    DirectoryPath = evt.DirectoryPath,
                    WindowStart   = DateTime.UtcNow,
                    LastEventAt   = DateTime.UtcNow,
                    PrimaryEvent  = evt
                };
                _groups[evt.DirectoryPath].Events.Add(evt);
            }
        }
    }

    private void FlushExpiredGroups(CancellationToken ct)
    {
        List<DebounceGroup> toProcess;

        lock (_groupLock)
        {
            var now = DateTime.UtcNow;
            var expired = _groups
                .Where(kv => (now - kv.Value.LastEventAt).TotalMilliseconds >= _settings.DebounceMs)
                .Select(kv => kv.Key)
                .ToList();

            if (expired.Count == 0) return;

            toProcess = expired.Select(k => _groups[k]).ToList();
            foreach (var k in expired)
                _groups.Remove(k);
        }

        _log.LogDebug("Flushing {N} debounce group(s)", toProcess.Count);

        foreach (var group in toProcess)
        {
            // Fire-and-forget into the thread pool; track for graceful shutdown.
            var task = Task.Run(
                () => ProcessGroupAsync(group, ct),
                CancellationToken.None); // don't cancel the task itself with ct here

            lock (_inflightLock)
                _inflight.Add(task);

            // Clean up completed tasks to avoid the list growing indefinitely.
            task.ContinueWith(_ =>
            {
                lock (_inflightLock)
                    _inflight.Remove(task);
            }, TaskContinuationOptions.ExecuteSynchronously);
        }
    }

    // ── Stages 3–5: Stabilize → Correlate → Act ─────────────────────────────

    private async Task ProcessGroupAsync(DebounceGroup group, CancellationToken ct)
    {
        var primary = group.PrimaryEvent;
        if (primary == null) return;

        _log.LogDebug(
            "Processing group: dir={Dir}, events={Count}, primary={File}",
            group.DirectoryPath, group.Events.Count, primary.FullPath);

        // Stage 3: Stabilization
        bool stable;
        try
        {
            stable = await _stabilizer.WaitForStableAsync(primary.FullPath, ct);
        }
        catch (OperationCanceledException)
        {
            _log.LogDebug("Stabilization cancelled for {File}", primary.FullPath);
            return;
        }

        if (!stable)
        {
            _log.LogDebug("Skipping unstable/deleted file: {File}", primary.FullPath);
            return;
        }

        // Stage 4: Local origin correlation
        var result = _correlator.Evaluate(primary);

        if (!result.Accepted)
        {
            _log.LogDebug(
                "Correlation rejected {File}: {Reason}",
                primary.FullPath, result.Reason);
            return;
        }

        _log.LogInformation(
            "File detected locally: {File} (reason: {Reason})",
            primary.FullPath, result.Reason);

        // Stage 5: Open Explorer
        try
        {
            await _explorer.OpenFolderAsync(group.DirectoryPath, primary.FullPath, ct);
        }
        catch (OperationCanceledException)
        {
            _log.LogDebug("Explorer open cancelled for {Dir}", group.DirectoryPath);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to open Explorer for {Dir}", group.DirectoryPath);
        }
    }
}
