using Serilog;
using Serilog.Events;

namespace SaveTrigger.Logging;

/// <summary>
/// Configures Serilog with two separate file sinks:
///
///   user-{date}.log   — Information and above. Human-readable one-line entries for
///                       events the user cares about (file detected, folder opened, etc.)
///
///   debug-{date}.log  — Debug and above. Full structured detail including raw events,
///                       debounce decisions, correlation reasoning, Explorer actions,
///                       and any errors. Intended for diagnosing missed or false detections.
///
/// Both logs rotate daily. The bootstrap logger is installed before the host starts so
/// that early startup errors are captured rather than lost to the void.
/// </summary>
public static class LoggingConfiguration
{
    private const string UserTemplate =
        "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}";

    private const string DebugTemplate =
        "[{Timestamp:HH:mm:ss.fff} {Level:u3}] [{SourceContext}] {Message:lj}{NewLine}{Exception}";

    /// <summary>
    /// Minimal logger active from app launch until the full logger replaces it.
    /// Only writes to the debug log so startup errors are captured.
    /// </summary>
    public static void ConfigureBootstrapLogger(string logPath)
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(
                path: Path.Combine(logPath, "debug-.log"),
                rollingInterval: RollingInterval.Day,
                outputTemplate: DebugTemplate,
                retainedFileCountLimit: 30)
            .CreateBootstrapLogger();
    }

    /// <summary>
    /// Full two-sink logger. Called by UseSerilog() after the host is configured.
    /// </summary>
    public static void ConfigureFullLogger(
        LoggerConfiguration config,
        string logPath)
    {
        config
            .MinimumLevel.Debug()
            .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
            .MinimumLevel.Override("Microsoft.Hosting.Lifetime", LogEventLevel.Information)

            // ── User log: Information and above ─────────────────────────────
            .WriteTo.Logger(lc => lc
                .Filter.ByIncludingOnly(e => e.Level >= LogEventLevel.Information)
                .WriteTo.File(
                    path: Path.Combine(logPath, "user-.log"),
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: UserTemplate,
                    retainedFileCountLimit: 30))

            // ── Debug log: everything ────────────────────────────────────────
            .WriteTo.Logger(lc => lc
                .WriteTo.File(
                    path: Path.Combine(logPath, "debug-.log"),
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: DebugTemplate,
                    retainedFileCountLimit: 30))

            .Enrich.FromLogContext();
    }
}
