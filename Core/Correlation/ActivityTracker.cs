using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;
using SaveTrigger.Interop;

namespace SaveTrigger.Core.Correlation;

/// <summary>
/// Hosted service that tracks foreground window changes using a WinEvent hook.
/// Must run on a dedicated STA thread with a Win32 message loop so the
/// WINEVENT_OUTOFCONTEXT hook fires correctly.
///
/// Public API: GetRecentActivity() — consumed by LocalOriginCorrelator.
/// </summary>
public sealed class ActivityTracker : IHostedService, IDisposable
{
    private readonly ActivityWindow _window;
    private readonly ILogger<ActivityTracker> _log;

    private Thread? _staThread;

    // Thread ID of the STA message-loop thread — used to post WM_QUIT on stop.
    private volatile uint _staThreadId;

    // Set to true once the message loop thread has started and captured its ID.
    private readonly TaskCompletionSource<bool> _threadReady = new(
        TaskCreationOptions.RunContinuationsAsynchronously);

    public ActivityTracker(
        IOptions<AppSettings> settings,
        ILogger<ActivityTracker> log)
    {
        _log = log;
        _window = new ActivityWindow(settings.Value.ActivityWindowSeconds);
    }

    // ── IHostedService ───────────────────────────────────────────────────────

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        _staThread = new Thread(RunMessageLoop)
        {
            IsBackground = true,
            Name = "ActivityTracker-STA"
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        // Wait until the thread has captured its ID and installed the hook.
        await _threadReady.Task.WaitAsync(TimeSpan.FromSeconds(5), cancellationToken);
        _log.LogInformation("ActivityTracker started (thread {ThreadId})", _staThreadId);
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        if (_staThreadId != 0)
        {
            // Post WM_QUIT to the STA thread's message queue, breaking the GetMessage loop.
            NativeMethods.PostThreadMessage(_staThreadId, NativeMethods.WM_QUIT, 0, 0);
        }

        // Give the thread a moment to exit cleanly — don't block indefinitely on shutdown.
        _staThread?.Join(TimeSpan.FromSeconds(2));
        _log.LogInformation("ActivityTracker stopped");
        return Task.CompletedTask;
    }

    // ── Public API ───────────────────────────────────────────────────────────

    /// <summary>
    /// Returns foreground window snapshots from the past ActivityWindowSeconds.
    /// Thread-safe — backed by ConcurrentQueue.
    /// </summary>
    public IReadOnlyList<ForegroundSnapshot> GetRecentActivity() => _window.GetRecent();

    // ── STA message-loop thread ──────────────────────────────────────────────

    private void RunMessageLoop()
    {
        _staThreadId = NativeMethods.GetCurrentThreadId();

        WinEventHook? hook = null;

        try
        {
            hook = new WinEventHook(
                NativeMethods.EVENT_SYSTEM_FOREGROUND,
                NativeMethods.EVENT_SYSTEM_FOREGROUND);

            hook.EventReceived += OnForegroundChanged;

            _log.LogDebug("WinEvent hook installed on thread {ThreadId}", _staThreadId);

            // Signal startup complete AFTER the hook is installed.
            _threadReady.TrySetResult(true);

            // Pump Win32 messages — this is what makes WINEVENT_OUTOFCONTEXT fire.
            int result;
            NativeMethods.MSG msg;
            while ((result = NativeMethods.GetMessage(out msg, nint.Zero, 0, 0)) != 0)
            {
                if (result == -1)
                {
                    _log.LogWarning("GetMessage returned error on ActivityTracker thread");
                    break;
                }
                NativeMethods.TranslateMessage(ref msg);
                NativeMethods.DispatchMessage(ref msg);
            }
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ActivityTracker STA thread encountered a fatal error");
            _threadReady.TrySetException(ex);
        }
        finally
        {
            hook?.Dispose();
            _log.LogDebug("ActivityTracker message loop exited");
        }
    }

    private void OnForegroundChanged(uint eventType, nint hwnd, string title, string className)
    {
        var snapshot = new ForegroundSnapshot(hwnd, title, className, DateTime.UtcNow);
        _window.Add(snapshot);

        _log.LogDebug(
            "Foreground changed → [{Class}] \"{Title}\" HWND=0x{Hwnd:X}",
            className, title, hwnd);
    }

    public void Dispose()
    {
        // Thread is background — no additional cleanup needed beyond StopAsync.
    }
}
