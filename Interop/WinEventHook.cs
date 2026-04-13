using System.Runtime.InteropServices;
using System.Text;

namespace SaveTrigger.Interop;

/// <summary>
/// Managed wrapper around the Win32 SetWinEventHook / UnhookWinEvent API.
///
/// Threading requirement (CRITICAL):
///   SetWinEventHook with WINEVENT_OUTOFCONTEXT requires the calling thread to
///   pump Win32 messages (GetMessage / DispatchMessage loop). The hook callback
///   fires on the same thread that called SetWinEventHook. Callers MUST create
///   this object on a dedicated STA thread that runs a message loop.
///
/// GC safety:
///   The native delegate (_procDelegate) is stored as a field so the GC cannot
///   collect it while the hook is active. Collecting it would cause a crash when
///   the native side fires the callback.
/// </summary>
internal sealed class WinEventHook : IDisposable
{
    private nint _hookHandle;

    // Keep the delegate alive — if GC collects it, the callback pointer becomes invalid.
    private readonly NativeMethods.WinEventProc _procDelegate;

    private bool _disposed;

    /// <summary>
    /// Fired on the hook thread for each captured event.
    /// Parameters: (eventType, hwnd, windowTitle, className)
    /// </summary>
    public event Action<uint, nint, string, string>? EventReceived;

    /// <param name="eventMin">First event code in the range to hook.</param>
    /// <param name="eventMax">Last event code in the range to hook.</param>
    public WinEventHook(uint eventMin, uint eventMax)
    {
        _procDelegate = OnWinEvent;

        _hookHandle = NativeMethods.SetWinEventHook(
            eventMin,
            eventMax,
            nint.Zero,          // in-process: no module handle needed for out-of-context
            _procDelegate,
            0,                  // monitor all processes
            0,                  // monitor all threads
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        if (_hookHandle == nint.Zero)
            throw new InvalidOperationException(
                $"SetWinEventHook failed for events [{eventMin:X4}–{eventMax:X4}]. " +
                "Ensure this is called from an STA thread with a running message loop.");
    }

    private void OnWinEvent(
        nint hWinEventHook,
        uint eventType,
        nint hwnd,
        int  idObject,
        int  idChild,
        uint idEventThread,
        uint dwmsEventTime)
    {
        if (hwnd == nint.Zero) return;

        var titleBuf = new StringBuilder(256);
        var classBuf = new StringBuilder(256);

        NativeMethods.GetWindowText(hwnd, titleBuf, titleBuf.Capacity);
        NativeMethods.GetClassName(hwnd, classBuf, classBuf.Capacity);

        try
        {
            EventReceived?.Invoke(eventType, hwnd, titleBuf.ToString(), classBuf.ToString());
        }
        catch
        {
            // Never throw from a native callback — it will crash the process.
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_hookHandle != nint.Zero)
        {
            NativeMethods.UnhookWinEvent(_hookHandle);
            _hookHandle = nint.Zero;
        }
    }
}
