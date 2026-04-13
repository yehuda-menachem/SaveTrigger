using System.Runtime.InteropServices;

namespace SaveTrigger.Interop;

/// <summary>
/// Centralized P/Invoke declarations. All Win32 calls live here so they can be
/// audited and updated in one place without touching business logic.
/// </summary>
internal static class NativeMethods
{
    // ── Foreground window ────────────────────────────────────────────────────

    [DllImport("user32.dll")]
    internal static extern nint GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetForegroundWindow(nint hWnd);

    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(nint hWnd, out uint processId);

    // ── Window text / class ──────────────────────────────────────────────────

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetWindowText(nint hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetClassName(nint hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

    // ── Window positioning ───────────────────────────────────────────────────

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

    // hWndInsertAfter constants
    internal static readonly nint HWND_TOP     = 0;
    internal static readonly nint HWND_TOPMOST = new(-1);

    // uFlags for SetWindowPos
    internal const uint SWP_NOSIZE         = 0x0001;
    internal const uint SWP_NOMOVE         = 0x0002;
    internal const uint SWP_NOZORDER       = 0x0004;
    internal const uint SWP_NOACTIVATE     = 0x0010;
    internal const uint SWP_SHOWWINDOW     = 0x0040;
    internal const uint SWP_ASYNCWINDOWPOS = 0x4000;

    // ── Show / focus window ──────────────────────────────────────────────────

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ShowWindow(nint hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool BringWindowToTop(nint hWnd);

    /// <summary>
    /// Attaches the message queue of thread idAttach to idAttachTo.
    /// Required for SetForegroundWindow to work from a background process:
    /// attach to the current foreground thread, call SetForegroundWindow, then detach.
    /// </summary>
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool AttachThreadInput(uint idAttach, uint idAttachTo,
        [MarshalAs(UnmanagedType.Bool)] bool fAttach);

    internal const int SW_RESTORE  = 9;
    internal const int SW_SHOW     = 5;
    internal const int SW_MAXIMIZE = 3;

    // ── Window state ─────────────────────────────────────────────────────────

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsWindow(nint hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsWindowVisible(nint hWnd);

    // ── Window enumeration ───────────────────────────────────────────────────

    internal delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    // ── Messages ─────────────────────────────────────────────────────────────

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool PostMessage(nint hWnd, uint Msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool PostThreadMessage(uint idThread, uint Msg, nint wParam, nint lParam);

    internal const uint WM_CLOSE = 0x0010;
    internal const uint WM_QUIT  = 0x0012;

    // ── WinEvent hook ────────────────────────────────────────────────────────

    /// <summary>
    /// Callback signature for SetWinEventHook. The delegate field holding this
    /// MUST be kept alive (pinned in a field) for the lifetime of the hook, or
    /// the GC will collect it and the hook callback will crash.
    /// </summary>
    internal delegate void WinEventProc(
        nint hWinEventHook,
        uint eventType,
        nint hwnd,
        int  idObject,
        int  idChild,
        uint idEventThread,
        uint dwmsEventTime);

    [DllImport("user32.dll")]
    internal static extern nint SetWinEventHook(
        uint eventMin,
        uint eventMax,
        nint hmodWinEventProc,
        WinEventProc lpfnWinEventProc,
        uint idProcess,
        uint idThread,
        uint dwFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool UnhookWinEvent(nint hWinEventHook);

    // WinEvent constants
    internal const uint EVENT_SYSTEM_FOREGROUND = 0x0003; // foreground window changed
    internal const uint WINEVENT_OUTOFCONTEXT   = 0x0000; // callback on caller's thread (needs message pump)
    internal const uint WINEVENT_SKIPOWNPROCESS = 0x0002; // don't call back for our own process

    // ── Win32 message loop ───────────────────────────────────────────────────
    // Required for WINEVENT_OUTOFCONTEXT hooks — the calling thread must pump messages.

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
        public nint  hwnd;
        public uint  message;
        public nint  wParam;
        public nint  lParam;
        public uint  time;
        public int   ptX;
        public int   ptY;
    }

    /// <summary>Returns >0 for normal messages, 0 for WM_QUIT, -1 on error.</summary>
    [DllImport("user32.dll")]
    internal static extern int GetMessage(out MSG lpMsg, nint hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool TranslateMessage(ref MSG lpMsg);

    [DllImport("user32.dll")]
    internal static extern nint DispatchMessage(ref MSG lpmsg);

    // ── Thread ID ────────────────────────────────────────────────────────────

    [DllImport("kernel32.dll")]
    internal static extern uint GetCurrentThreadId();

    // ── SendInput (keyboard simulation) ──────────────────────────────────────
    // Used to send Ctrl+T / Ctrl+Shift+Tab to Explorer for tab management.

    [StructLayout(LayoutKind.Sequential)]
    internal struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint   dwFlags;
        public uint   time;
        public nint   dwExtraInfo;
    }

    // INPUT must be exactly 40 bytes on 64-bit Windows.
    // Layout: type(4) + padding(4) + union(32). KEYBDINPUT starts at offset 8.
    [StructLayout(LayoutKind.Explicit, Size = 40)]
    internal struct INPUT
    {
        [FieldOffset(0)] public uint       type;
        [FieldOffset(8)] public KEYBDINPUT ki;
    }

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern uint SendInput(uint nInputs,
        [In] INPUT[] pInputs, int cbSize);

    internal const uint   INPUT_KEYBOARD   = 1;
    internal const uint   KEYEVENTF_KEYUP  = 0x0002;
    internal const ushort VK_CONTROL       = 0x11;
    internal const ushort VK_SHIFT         = 0x10;
    internal const ushort VK_TAB           = 0x09;
    internal const ushort VK_T             = 0x54;
    internal const ushort VK_W             = 0x57;

    /// <summary>
    /// Sends a keyboard shortcut (optional Ctrl + optional Shift + key) via SendInput.
    /// The window must be in the foreground for SendInput to target it correctly.
    /// </summary>
    internal static void SendKeyCombo(ushort key, bool ctrl = false, bool shift = false)
    {
        var list = new System.Collections.Generic.List<INPUT>(6);

        if (ctrl)  list.Add(MakeKey(VK_CONTROL, false));
        if (shift) list.Add(MakeKey(VK_SHIFT,   false));
        list.Add(MakeKey(key, false));
        list.Add(MakeKey(key, true));
        if (shift) list.Add(MakeKey(VK_SHIFT,   true));
        if (ctrl)  list.Add(MakeKey(VK_CONTROL, true));

        var arr = list.ToArray();
        SendInput((uint)arr.Length, arr, System.Runtime.InteropServices.Marshal.SizeOf<INPUT>());
    }

    private static INPUT MakeKey(ushort vk, bool keyUp) => new INPUT
    {
        type = INPUT_KEYBOARD,
        ki   = new KEYBDINPUT { wVk = vk, dwFlags = keyUp ? KEYEVENTF_KEYUP : 0u }
    };

    // ── Drive type ───────────────────────────────────────────────────────────

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    internal static extern uint GetDriveType(string lpRootPathName);

    internal const uint DRIVE_UNKNOWN     = 0;
    internal const uint DRIVE_NO_ROOT_DIR = 1;
    internal const uint DRIVE_REMOVABLE   = 2;
    internal const uint DRIVE_FIXED       = 3;
    internal const uint DRIVE_REMOTE      = 4; // network / mapped drive
    internal const uint DRIVE_CDROM       = 5;
    internal const uint DRIVE_RAMDISK     = 6;
}
