using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SaveTrigger.Interop;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
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
    /// Creates a radar-style icon: concentric green rings on a dark circular background,
    /// with a sweep arm and a ping dot — matching the app's radar logo aesthetic.
    /// Transparent corners are preserved by embedding PNG inside the ICO stream.
    /// </summary>
    private static Icon CreateDefaultIcon()
    {
        try
        {
            using var bmp = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g   = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Fully transparent canvas — corners stay transparent
            g.Clear(Color.Transparent);

            const float cx = 16f, cy = 16f;
            const float radius = 15f;

            // Circular dark background (clipped to circle)
            using var bgPath = new GraphicsPath();
            bgPath.AddEllipse(cx - radius, cy - radius, radius * 2, radius * 2);
            g.SetClip(bgPath);
            g.FillEllipse(new SolidBrush(Color.FromArgb(255, 10, 26, 15)), cx - radius, cy - radius, radius * 2, radius * 2);

            // Sweep wedge (upper-right, ~65°)
            g.FillPie(new SolidBrush(Color.FromArgb(50, 0, 220, 90)), 1, 1, 30, 30, -90f, -65f);

            // Concentric rings
            using var ringPen = new Pen(Color.FromArgb(150, 0, 200, 80), 0.8f);
            foreach (float r in new[] { 13.5f, 9.5f, 5.5f })
                g.DrawEllipse(ringPen, cx - r, cy - r, r * 2, r * 2);

            // Crosshair lines
            using var crossPen = new Pen(Color.FromArgb(55, 0, 200, 80), 0.5f);
            g.DrawLine(crossPen, 2, cy, 30, cy);
            g.DrawLine(crossPen, cx, 2, cx, 30);

            // Sweep arm (toward upper-right ~45°)
            using var sweepPen = new Pen(Color.FromArgb(210, 0, 255, 100), 1f);
            g.DrawLine(sweepPen, cx, cy, 27f, 5f);

            // Ping dot
            using var pingBrush = new SolidBrush(Color.FromArgb(255, 0, 255, 110));
            g.FillEllipse(pingBrush, 23f, 5.5f, 4f, 4f);

            // Center dot
            using var centerBrush = new SolidBrush(Color.FromArgb(230, 0, 255, 110));
            g.FillEllipse(centerBrush, cx - 1.5f, cy - 1.5f, 3f, 3f);

            g.ResetClip();

            // Border ring
            using var borderPen = new Pen(Color.FromArgb(80, 0, 200, 80), 0.6f);
            g.DrawEllipse(borderPen, cx - radius, cy - radius, radius * 2, radius * 2);

            return IconFromBitmap(bmp);
        }
        catch
        {
            return SystemIcons.Application;
        }
    }

    /// <summary>
    /// Converts a 32bpp ARGB bitmap to an Icon while preserving full alpha transparency
    /// by embedding the PNG inside an ICO stream (supported on Vista+).
    /// </summary>
    private static Icon IconFromBitmap(Bitmap bmp)
    {
        using var pngStream = new System.IO.MemoryStream();
        bmp.Save(pngStream, System.Drawing.Imaging.ImageFormat.Png);
        byte[] png = pngStream.ToArray();

        using var icoStream = new System.IO.MemoryStream();
        using var bw = new System.IO.BinaryWriter(icoStream);

        // ICONDIR
        bw.Write((short)0);            // reserved
        bw.Write((short)1);            // type: ICO
        bw.Write((short)1);            // image count

        // ICONDIRENTRY (16 bytes)
        bw.Write((byte)bmp.Width);
        bw.Write((byte)bmp.Height);
        bw.Write((byte)0);             // color count (0 = >256)
        bw.Write((byte)0);             // reserved
        bw.Write((short)1);            // planes
        bw.Write((short)32);           // bits per pixel
        bw.Write((int)png.Length);     // image data size
        bw.Write((int)22);             // offset: 6 (ICONDIR) + 16 (ICONDIRENTRY)

        bw.Write(png);
        bw.Flush();

        icoStream.Position = 0;
        return new Icon(icoStream);
    }

    public void Dispose()
    {
        _notifyIcon?.Dispose();
    }
}
