using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SaveTrigger.Interop;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace SaveTrigger.Tray;

/// <summary>
/// Hosted service that shows the system tray icon and context menu.
///
/// Threading model:
///   NotifyIcon must live on an STA thread with a running Windows message loop.
///   This service spawns a dedicated STA thread and calls Application.Run() on it.
///   Shutdown is triggered by posting WM_QUIT to that thread via PostThreadMessage.
///
/// Menu items:
///   - Pause / Resume  — toggles AppState.IsPaused
///   - Open Logs       — opens %APPDATA%\SaveTrigger\Logs\ in Explorer
///   - Exit            — stops the host
/// </summary>
public sealed class TrayIconService : IHostedService, IDisposable
{
    private readonly IHostApplicationLifetime _lifetime;
    private readonly AppState _appState;
    private readonly ILogger<TrayIconService> _log;

    private Thread? _staThread;
    private volatile uint _staThreadId;

    // Owned by the STA thread — must not be accessed from other threads.
    private NotifyIcon? _notifyIcon;
    private ToolStripMenuItem? _pauseMenuItem;

    // Log folder path — computed once, read from both threads (readonly after ctor).
    private readonly string _logPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SaveTrigger", "Logs");

    public TrayIconService(
        IHostApplicationLifetime lifetime,
        AppState appState,
        ILogger<TrayIconService> log)
    {
        _lifetime = lifetime;
        _appState = appState;
        _log      = log;
    }

    // ── IHostedService ───────────────────────────────────────────────────────

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _staThread = new Thread(RunTrayThread)
        {
            IsBackground = true,
            Name         = "TrayIcon-STA"
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        _log.LogInformation("TrayIconService started");
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        if (_staThreadId != 0)
        {
            // Post WM_QUIT to the STA thread — breaks Application.Run().
            NativeMethods.PostThreadMessage(_staThreadId, NativeMethods.WM_QUIT, 0, 0);
        }

        _staThread?.Join(TimeSpan.FromSeconds(3));
        _log.LogInformation("TrayIconService stopped");
        return Task.CompletedTask;
    }

    // ── STA thread ───────────────────────────────────────────────────────────

    private void RunTrayThread()
    {
        _staThreadId = NativeMethods.GetCurrentThreadId();

        try
        {
            Application.EnableVisualStyles();
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            _pauseMenuItem = new ToolStripMenuItem("Pause", null, OnPauseResumeClicked);

            var menu = new ContextMenuStrip();
            menu.Items.Add(_pauseMenuItem);
            menu.Items.Add("Open Logs Folder", null, OnOpenLogsClicked);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Exit", null, OnExitClicked);

            _notifyIcon = new NotifyIcon
            {
                Text            = "SaveTrigger — Running",
                Icon            = CreateDefaultIcon(),
                ContextMenuStrip = menu,
                Visible         = true
            };

            // Register shutdown hook: when the host stops externally (e.g., Ctrl+C),
            // post WM_QUIT to this thread so Application.Run() exits cleanly.
            _lifetime.ApplicationStopping.Register(() =>
            {
                if (_staThreadId != 0)
                    NativeMethods.PostThreadMessage(_staThreadId, NativeMethods.WM_QUIT, 0, 0);
            });

            // Blocks until WM_QUIT is posted.
            Application.Run();
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "TrayIconService STA thread encountered an error");
        }
        finally
        {
            // Hide and dispose the icon before the thread exits so it doesn't
            // linger as a ghost icon in the tray.
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }
        }
    }

    // ── Menu event handlers (STA thread) ─────────────────────────────────────

    private void OnPauseResumeClicked(object? sender, EventArgs e)
    {
        _appState.IsPaused = !_appState.IsPaused;
        bool paused = _appState.IsPaused;

        if (_pauseMenuItem != null)
            _pauseMenuItem.Text = paused ? "Resume" : "Pause";

        if (_notifyIcon != null)
            _notifyIcon.Text = paused ? "SaveTrigger — Paused" : "SaveTrigger — Running";

        _log.LogInformation("SaveTrigger {State}", paused ? "paused" : "resumed");
    }

    private void OnOpenLogsClicked(object? sender, EventArgs e)
    {
        try
        {
            Directory.CreateDirectory(_logPath);
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{_logPath}\"")
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Could not open logs folder: {Path}", _logPath);
        }
    }

    private void OnExitClicked(object? sender, EventArgs e)
    {
        _log.LogInformation("Exit requested from tray menu");

        if (_notifyIcon != null)
            _notifyIcon.Visible = false;

        // StopApplication triggers IHostApplicationLifetime.ApplicationStopping,
        // which begins graceful shutdown of all hosted services.
        _lifetime.StopApplication();
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates a simple programmatic icon (white 'S' on a blue background) so
    /// the app runs without an embedded .ico resource file.
    /// Replace with a proper .ico file for production.
    /// </summary>
    private static Icon CreateDefaultIcon()
    {
        try
        {
            using var bmp = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g   = Graphics.FromImage(bmp);

            g.Clear(Color.FromArgb(0, 120, 215));  // Windows blue

            using var font  = new Font("Arial", 18, FontStyle.Bold, GraphicsUnit.Pixel);
            using var brush = new SolidBrush(Color.White);

            var sf = new StringFormat
            {
                Alignment     = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString("S", font, brush, new RectangleF(0, 0, 32, 32), sf);

            // Convert Bitmap to Icon
            var hIcon = bmp.GetHicon();
            return Icon.FromHandle(hIcon);
        }
        catch
        {
            // Last resort: use SystemIcons which is always available.
            return SystemIcons.Application;
        }
    }

    public void Dispose()
    {
        _notifyIcon?.Dispose();
    }
}
