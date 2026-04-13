using System.Runtime.InteropServices;

namespace SaveTrigger.Interop;

// Shell COM interfaces used to enumerate open Explorer windows.
// We use minimal interface definitions — only the members we actually call.
// Dynamic dispatch (via 'dynamic') is used for IShellWindows items to avoid
// implementing the full IWebBrowser2 vtable.

/// <summary>
/// IShellWindows — COM interface for the running Shell.Application object.
/// Provides access to all currently open Explorer / IE windows.
/// ProgID: "Shell.Application" → .Windows() property returns IShellWindows.
/// </summary>
[ComImport]
[Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
internal interface IShellWindows
{
    [DispId(0x60020000)] int Count { get; }
    [DispId(0x60020001)] object? Item(object index);
    [DispId(0x60020002)] object _NewEnum();
}

/// <summary>
/// CoClass for the ShellWindows collection.
/// Usage: new ShellWindowsCoClass() as IShellWindows
/// </summary>
[ComImport]
[Guid("9BA05971-F6A8-11CF-A442-00A0C90A8F39")]
[ClassInterface(ClassInterfaceType.None)]
internal class ShellWindowsCoClass { }

/// <summary>
/// Well-known GUIDs used during Shell COM service queries.
/// </summary>
internal static class ShellGuids
{
    public static readonly Guid SID_STopLevelBrowser =
        new("4C96BE40-915C-11CF-99D3-00AA004AE837");

    public static readonly Guid IID_IShellBrowser =
        new("000214E2-0000-0000-C000-000000000046");
}
