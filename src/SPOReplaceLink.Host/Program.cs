var host = SPOReplaceLink.Host.HostApp.Build(args);
host.Services.GetRequiredService<SPOReplaceLink.Core.Localization.LanguageStore>().Apply();
host.Run();

/// <summary>Entry-point marker of the host assembly (for tests: WebApplicationFactory).</summary>
namespace SPOReplaceLink.Host
{
    public sealed class HostEntry;
}
