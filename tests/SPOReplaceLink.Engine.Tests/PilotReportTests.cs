using System.Text.Json;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.State;
using SPOReplaceLink.Tests.Shared;
using static SPOReplaceLink.Engine.Tests.EngineEnv;

namespace SPOReplaceLink.Engine.Tests;

/// <summary>Stage 5: what a pilot run measures and the sample for the manual check.</summary>
public class PilotReportTests
{
    [Fact]
    public async Task ApplyReportShowsLoadVersionsAndFilesChangedBySharePoint()
    {
        await using var env = new EngineEnv();
        for (int i = 0; i < 3; i++)
            env.AddXlsx($"Fin/c{i}.xlsx", OldUrl);
        env.AddXlsx("Fin/same.xlsx", NewUrl);
        bool promoted = false;
        env.Graph.OnUploaded = (_, item) =>
        {
            if (promoted)
                return;
            promoted = true;
            item.Content = [.. item.Content, .. new byte[100]]; // library columns written into the file
        };
        long run = env.Create(RunModes.Apply);
        Assert.Equal(RunStatuses.Completed, await env.RunAsync(run));

        string text = PilotReport.Text(env.Store, run);
        Assert.Contains("записано в SharePoint: 3", text);
        Assert.Contains("RU на файл:", text);
        Assert.Contains("номер версії до запису відомий: 3 з 3", text);
        Assert.Contains("SharePoint змінив розмір файлу під час збереження (за відповіддю на запис): 1 з 3", text);
        Assert.Contains("дописує у файли властивості", text);
        Assert.Contains("межа бюджету RU", text);
        Assert.Contains("на профіль «test»", text);
        Assert.Contains("с на файл", text);
        var load = PilotReport.Load(env.Store.GetMetrics(run).Select(m => m.Json));
        Assert.True(load.Ru >= 4 * 2, $"RU {load.Ru}");

        var sample = PilotReport.Sample(env.Store, run).ToList();
        Assert.Equal(4, sample.Count); // header + 3 changed files
        Assert.All(sample.Skip(1), row =>
        {
            Assert.StartsWith("https://", row[6]);
            Assert.Contains("/Fin/c", row[6]);
            Assert.Contains(NewUrl, row[7]);
            Assert.Equal("1.0", row[5]);
        });
        Assert.Equal(sample.Select(r => r[0]), PilotReport.Sample(env.Store, run).Select(r => r[0])); // the same every time
        Assert.Equal(2, PilotReport.Sample(env.Store, run, 2).Count() - 1);
    }

    [Fact]
    public async Task ScanReportSaysNothingWasWritten()
    {
        await using var env = new EngineEnv();
        env.AddXlsx("a.xlsx", OldUrl);
        long run = env.Create(RunModes.Scan);
        await env.RunAsync(run);
        string text = PilotReport.Text(env.Store, run);
        Assert.Contains("буде змінено: 1", text);
        Assert.Contains("прогін нічого не записував", text);

        using var stream = new MemoryStream();
        ReportExporter.Write(env.Store, run, "pilot", stream);
        Assert.Contains("Звіт пілоту", System.Text.Encoding.UTF8.GetString(stream.ToArray()));
    }

    [Fact]
    public void LoadAddsUpSessionsOfARun()
    {
        var t0 = new DateTimeOffset(2026, 10, 26, 20, 0, 0, TimeSpan.Zero);
        string Snap(DateTimeOffset started, int minute, long finished, long ru, long resized) => JsonSerializer.Serialize(new MetricsSnapshot(
            t0.AddMinutes(minute), finished, finished, 0, 0, 0, new Dictionary<string, long>(), new Dictionary<string, long>(),
            [new AppSnapshot("app-a", ru, ru, 0, 0, resized, 0, ru, null, 1, 0, false)], 0, 0, null, resized, started), Profile.JsonOptions);

        // Two starts of the same run: counters begin at zero again after a restart. The second session ends exactly
        // where the first one did, so only the start time tells them apart. Minutes 3–4 of the first session and the
        // weekend between the sessions processed nothing and are not active time.
        var s1 = t0;
        var s2 = t0.AddDays(2);
        var load = PilotReport.Load([
            Snap(s1, 1, 5, 10, 1), Snap(s1, 2, 9, 30, 2), Snap(s1, 3, 9, 31, 2), Snap(s1, 4, 9, 32, 2),
            Snap(s2, 2 * 24 * 60 + 1, 3, 32, 2), Snap(s2, 2 * 24 * 60 + 2, 6, 40, 2), "not json"]);
        Assert.Equal(72, load.Ru);
        Assert.Equal(4, load.ResizedBySpo);
        Assert.Equal(4, load.Throttled429);
        Assert.Equal(72, load.RuByApp["app-a"]);
        Assert.Equal(TimeSpan.FromMinutes(4), load.ActiveTime);
    }
}
