using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SaveTrigger.Configuration;

namespace SaveTrigger.Core.Pipeline;

/// <summary>
/// Stage 3: Waits until a file is fully written before allowing the pipeline to proceed.
///
/// Strategy:
///   1. Poll FileInfo.Length twice, separated by FileStabilizationPollMs.
///      If the size changed, the file is still being written.
///   2. Try to open the file with exclusive access (FileShare.None).
///      If this succeeds, no other process holds the file open.
///   3. Retry the above up to FileStabilizationRetries times.
///
/// Returns true when the file is stable, false if it cannot be stabilized
/// (deleted during polling, still locked after all retries, etc.).
/// </summary>
public sealed class FileStabilizer
{
    private readonly AppSettings _settings;
    private readonly ILogger<FileStabilizer> _log;

    public FileStabilizer(IOptions<AppSettings> settings, ILogger<FileStabilizer> log)
    {
        _settings = settings.Value;
        _log      = log;
    }

    public async Task<bool> WaitForStableAsync(string filePath, CancellationToken ct)
    {
        for (int attempt = 1; attempt <= _settings.FileStabilizationRetries; attempt++)
        {
            ct.ThrowIfCancellationRequested();

            // Small initial wait to let the OS flush the write.
            await Task.Delay(_settings.FileStabilizationPollMs, ct);

            if (!File.Exists(filePath))
            {
                _log.LogDebug("Stabilization: {File} no longer exists (deleted before processing)", filePath);
                return false;
            }

            // ── Size stability check ─────────────────────────────────────────
            long size1;
            long size2;
            try
            {
                size1 = new FileInfo(filePath).Length;
            }
            catch (Exception ex)
            {
                _log.LogDebug(ex, "Stabilization attempt {N}: could not stat {File}", attempt, filePath);
                continue;
            }

            await Task.Delay(_settings.FileStabilizationPollMs, ct);

            if (!File.Exists(filePath)) return false;

            try
            {
                size2 = new FileInfo(filePath).Length;
            }
            catch (Exception ex)
            {
                _log.LogDebug(ex, "Stabilization attempt {N}: second stat failed for {File}", attempt, filePath);
                continue;
            }

            if (size1 != size2)
            {
                _log.LogDebug(
                    "Stabilization attempt {N}: {File} still growing ({S1} → {S2} bytes)",
                    attempt, filePath, size1, size2);
                continue;
            }

            // ── Exclusive-open check ─────────────────────────────────────────
            try
            {
                using var fs = new FileStream(
                    filePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.None,   // exclusive — fails if anyone else has it open
                    bufferSize: 1,
                    FileOptions.None);

                _log.LogDebug(
                    "Stabilization: {File} is stable after {N} attempt(s) ({Size} bytes)",
                    filePath, attempt, size2);
                return true;
            }
            catch (IOException)
            {
                _log.LogDebug(
                    "Stabilization attempt {N}: {File} still locked by another process",
                    attempt, filePath);
                // continue to next attempt
            }
            catch (UnauthorizedAccessException ex)
            {
                // We can't open the file at all — still report as stable so the
                // folder opens even if we can't verify exclusive access.
                _log.LogDebug(ex, "Stabilization: {File} is read-protected; treating as stable", filePath);
                return true;
            }
        }

        _log.LogWarning(
            "Stabilization failed for {File} after {N} retries — skipping",
            filePath, _settings.FileStabilizationRetries);
        return false;
    }
}
