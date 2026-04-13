using System.Collections.Concurrent;

namespace SaveTrigger.Core.Correlation;

/// <summary>
/// A foreground window snapshot captured by the WinEvent hook.
/// </summary>
public sealed record ForegroundSnapshot(
    nint Hwnd,
    string WindowTitle,
    string ClassName,
    DateTime CapturedAt
);

/// <summary>
/// Thread-safe sliding window of recent foreground window changes.
/// Snapshots older than ActivityWindowSeconds are pruned automatically.
/// </summary>
public sealed class ActivityWindow
{
    private readonly ConcurrentQueue<ForegroundSnapshot> _snapshots = new();
    private readonly int _windowSeconds;

    public ActivityWindow(int windowSeconds)
    {
        _windowSeconds = windowSeconds;
    }

    public void Add(ForegroundSnapshot snapshot)
    {
        _snapshots.Enqueue(snapshot);
        Prune();
    }

    public IReadOnlyList<ForegroundSnapshot> GetRecent()
    {
        Prune();
        return _snapshots.ToArray();
    }

    public bool HasAny()
    {
        Prune();
        return !_snapshots.IsEmpty;
    }

    private void Prune()
    {
        var cutoff = DateTime.UtcNow.AddSeconds(-_windowSeconds);
        // ConcurrentQueue has no efficient remove; dequeue from front while stale
        while (_snapshots.TryPeek(out var oldest) && oldest.CapturedAt < cutoff)
            _snapshots.TryDequeue(out _);
    }
}
