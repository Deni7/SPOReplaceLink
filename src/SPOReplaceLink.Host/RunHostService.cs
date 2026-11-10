using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Engine;

namespace SPOReplaceLink.Host;

/// <summary>
/// STA-3: a run interrupted by a host crash or a reboot is continued when the host starts;
/// on shutdown the active run is stopped gracefully (STA-4).
/// </summary>
internal sealed class RunHostService(RunService runs, IConfiguration configuration, ILogger<RunHostService> logger) : IHostedService
{
    public Task StartAsync(CancellationToken cancellationToken)
    {
        if (configuration.GetValue("continueInterrupted", true))
        {
            try
            {
                runs.ContinueInterrupted();
            }
            catch (Exception ex) when (ex is InvalidOperationException or KeyNotFoundException)
            {
                logger.LogError(ex, L.T("Не вдалося продовжити перерваний прогін"));
            }
        }
        return Task.CompletedTask;
    }

    public async Task StopAsync(CancellationToken cancellationToken) => await runs.DisposeAsync();
}
