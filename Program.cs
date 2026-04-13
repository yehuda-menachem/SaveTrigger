using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using SaveTrigger.Configuration;
using SaveTrigger.Core.Correlation;
using SaveTrigger.Core.Explorer;
using SaveTrigger.Core.Models;
using SaveTrigger.Core.Pipeline;
using SaveTrigger.Logging;
using SaveTrigger.Tray;
using System.Threading.Channels;

namespace SaveTrigger;

/// <summary>
/// Application entry point.
///
/// Threading notes:
///   [STAThread] is required for COM initialization (Shell COM, Windows.Forms COM).
///   We use a synchronous Main that blocks on the async host via GetAwaiter().GetResult()
///   so the STA apartment state is preserved for COM calls made before any await.
///   The tray and WinEvent hook each create their own dedicated STA threads.
/// </summary>
static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        RunAsync(args).GetAwaiter().GetResult();
    }

    private static async Task RunAsync(string[] args)
    {
        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SaveTrigger", "Logs");

        // Ensure log directory exists before the bootstrap logger tries to write.
        Directory.CreateDirectory(logPath);

        // Install a minimal logger immediately so startup errors are captured.
        LoggingConfiguration.ConfigureBootstrapLogger(logPath);

        try
        {
            var host = Host.CreateDefaultBuilder(args)
                .UseContentRoot(AppContext.BaseDirectory)
                .UseSerilog((ctx, services, config) =>
                    LoggingConfiguration.ConfigureFullLogger(config, logPath))
                .ConfigureServices((ctx, services) =>
                {
                    // ── Settings ─────────────────────────────────────────────
                    services.Configure<AppSettings>(
                        ctx.Configuration.GetSection("AppSettings"));

                    // ── Shared pipeline channel ──────────────────────────────
                    // Bounded with DropOldest so high-frequency events never block
                    // the FileSystemWatcher callbacks or exhaust memory.
                    services.AddSingleton(_ =>
                        Channel.CreateBounded<FileEvent>(new BoundedChannelOptions(1000)
                        {
                            FullMode    = BoundedChannelFullMode.DropOldest,
                            SingleReader = true,
                            SingleWriter = false   // multiple watcher threads write
                        }));

                    // ── Singletons ───────────────────────────────────────────
                    services.AddSingleton<AppState>();
                    services.AddSingleton<ExplorerTabHelper>();
                    services.AddSingleton<ExplorerWindowManager>();
                    services.AddSingleton<FileStabilizer>();
                    services.AddSingleton<LocalOriginCorrelator>();

                    // ActivityTracker is both a singleton (so LocalOriginCorrelator
                    // can inject it) and a hosted service (so the host manages its
                    // lifetime). Register once, resolve as hosted service via factory.
                    services.AddSingleton<ActivityTracker>();

                    // ── Hosted services (startup order) ──────────────────────
                    // 1. TrayIcon   — user sees feedback immediately
                    // 2. Activity   — hook installed before watchers fire
                    // 3. Watcher    — starts emitting events
                    // 4. Debounce   — starts consuming events
                    services.AddHostedService<TrayIconService>();
                    services.AddHostedService(p => p.GetRequiredService<ActivityTracker>());
                    services.AddHostedService<FileWatcherService>();
                    services.AddHostedService<DebounceService>();
                })
                .Build();

            Log.Information("SaveTrigger starting");
            await host.RunAsync();
            Log.Information("SaveTrigger stopped cleanly");
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "SaveTrigger terminated unexpectedly");
        }
        finally
        {
            // Flush and close all Serilog sinks before the process exits.
            await Log.CloseAndFlushAsync();
        }
    }
}
