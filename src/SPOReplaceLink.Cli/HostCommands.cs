using System.Globalization;
using System.Text.Json.Nodes;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Cli;

/// <summary>Commands that work through the host: settings, runs, retry, reports.</summary>
internal sealed class HostCommands(HostClient host, TextWriter output)
{
    public const int RunNotFinished = 3;

    // A method, not a cached dictionary: the word follows the current language.
    private static string Kind(string kind) => kind switch
    {
        "profiles" => L.T("профіль"),
        "connections" => L.T("підключення"),
        _ => L.T("шаблон"),
    };

    // ------------------------------------------------------------------ settings

    public async Task<int> ListSettings(string kind)
    {
        var items = (await host.GetAsync($"api/{kind}"))!.AsArray();
        if (items.Count == 0)
            output.WriteLine(L.T("(порожньо)"));
        foreach (var i in items)
            output.WriteLine($"{i!["name"],-30} {i["updatedAt"]}");
        return ExitCodes.Ok;
    }

    public async Task<int> GetSetting(string kind, string name, string? file)
    {
        string json = await host.GetTextAsync($"api/{kind}/{Uri.EscapeDataString(name)}");
        if (file is null)
            output.WriteLine(json);
        else
            await File.WriteAllTextAsync(file, json);
        return ExitCodes.Ok;
    }

    /// <summary>Import from JSON (CFG-1). The name comes from --name, the "name" field, or the file name.</summary>
    public async Task<int> PutSetting(string kind, string file, string? name)
    {
        string json = await File.ReadAllTextAsync(file);
        name ??= (JsonNode.Parse(json, documentOptions: new() { CommentHandling = System.Text.Json.JsonCommentHandling.Skip, AllowTrailingCommas = true })?["name"]?.GetValue<string>())
            ?? Path.GetFileNameWithoutExtension(file).Split('.')[0];
        var result = await host.PutJsonTextAsync($"api/{kind}/{Uri.EscapeDataString(name)}", json);
        output.WriteLine(L.F("Збережено {0} «{1}».", Kind(kind), name));
        return PrintIssues(result?["issues"]?.AsArray());
    }

    public async Task<int> DeleteSetting(string kind, string name)
    {
        await host.DeleteAsync($"api/{kind}/{Uri.EscapeDataString(name)}");
        output.WriteLine(L.F("Видалено {0} «{1}».", Kind(kind), name));
        return ExitCodes.Ok;
    }

    public async Task<int> ValidateProfile(string name)
    {
        var result = (await host.PostAsync($"api/profiles/{Uri.EscapeDataString(name)}/validate"))!;
        int code = PrintIssues(result["issues"]?.AsArray());
        output.WriteLine(result["valid"]!.GetValue<bool>() ? L.T("Профіль готовий до запуску.") : L.T("Профіль містить помилки."));
        return result["valid"]!.GetValue<bool>() ? ExitCodes.Ok : Math.Max(code, ExitCodes.InvalidInput);
    }

    public async Task<int> TestConnection(string name, string? profile)
    {
        string query = profile is null ? string.Empty : "?profile=" + Uri.EscapeDataString(profile);
        var checks = (await host.PostAsync($"api/connections/{Uri.EscapeDataString(name)}/test{query}"))!.AsArray();
        bool ok = true;
        foreach (var c in checks)
        {
            bool passed = c!["ok"]!.GetValue<bool>();
            ok &= passed;
            output.WriteLine($"{(passed ? "OK  " : "FAIL")} {c["what"]}: {c["message"]}");
        }
        return ok ? ExitCodes.Ok : ExitCodes.Attention;
    }

    private int PrintIssues(JsonArray? issues)
    {
        bool errors = false;
        foreach (var i in issues ?? [])
        {
            string severity = i!["severity"]!.GetValue<string>();
            errors |= severity == "error";
            string rule = i["rule"] is { } r ? L.F(" [правило {0}]", r) : string.Empty;
            output.WriteLine($"{(severity == "error" ? L.T("ПОМИЛКА") : L.T("Попередження"))}{rule}: {i["message"]}");
        }
        return errors ? ExitCodes.InvalidInput : ExitCodes.Ok;
    }

    // ------------------------------------------------------------------ runs

    /// <summary>inventory / scan / apply / verify of a whole profile.</summary>
    public async Task<int> StartRun(string mode, string profile, bool wait, bool yes, CancellationToken ct)
    {
        if (mode == "apply" && !Confirm(L.F("Запустити apply за профілем «{0}»? Змінені файли буде записано в SharePoint.", profile), yes))
            return ExitCodes.InvalidInput;
        var created = (await host.PostAsync("api/runs", new { mode, profile, start = true }, ct))!;
        long id = created["id"]!.GetValue<long>();
        output.WriteLine(L.F("Прогін {0} ({1}) запущено.", id, mode));
        return wait ? await Wait(id, ct) : ExitCodes.Ok;
    }

    /// <summary>RTY-4: a new run over files of a finished run picked by filter.</summary>
    public async Task<int> Retry(long runId, RetryArgs a, bool wait, bool yes, CancellationToken ct)
    {
        var paths = new List<string>();
        if (a.ListFile is not null)
            paths.AddRange(ReadPathList(a.ListFile));
        var parent = (await host.GetAsync($"api/runs/{runId}", ct))!;
        string mode = a.Mode ?? parent["run"]!["mode"]!.GetValue<string>();
        if (mode == "apply" && !Confirm(L.F("Повторити прогін {0} у режимі apply?", runId), yes))
            return ExitCodes.InvalidInput;
        var body = new
        {
            statuses = a.Statuses.Length == 0 && a.ErrorClasses.Length == 0 && a.FileIds.Length == 0 && paths.Count == 0 ? ["failed"] : a.Statuses,
            errorClasses = a.ErrorClasses,
            extensions = a.Extensions,
            pathMask = a.PathMask,
            paths,
            fileIds = a.FileIds,
            mode,
            currentProfile = a.CurrentProfile,
            start = true,
        };
        var created = (await host.PostAsync($"api/runs/{runId}/retry", body, ct))!;
        long id = created["id"]!.GetValue<long>();
        output.WriteLine(L.F("Повторний прогін {0} ({1}): файлів {2}.", id, mode, created["files"]));
        return wait ? await Wait(id, ct) : ExitCodes.Ok;
    }

    public async Task<int> VerifyRun(long runId, bool wait, CancellationToken ct)
    {
        var created = (await host.PostAsync($"api/runs/{runId}/verify", null, ct))!;
        long id = created["id"]!.GetValue<long>();
        output.WriteLine(L.F("Перевірка {0} файлів, записаних прогоном {1}: {2}.", id, runId, created["files"]));
        return wait ? await Wait(id, ct) : ExitCodes.Ok;
    }

    public async Task<int> Control(long runId, string action, bool wait, CancellationToken ct)
    {
        await host.PostAsync($"api/runs/{runId}/{action}", null, ct);
        output.WriteLine(action switch
        {
            "start" => L.F("Прогін {0} продовжено.", runId),
            "pause" => L.F("Прогін {0} на паузі: нові файли не беруться.", runId),
            "resume" => L.F("Прогін {0} продовжено після паузи.", runId),
            _ => L.F("Прогін {0}: зупинка після завершення поточних файлів.", runId),
        });
        return wait ? await Wait(runId, ct) : ExitCodes.Ok;
    }

    /// <summary>Progress until the run reaches a final status. Ctrl+C asks the host for a graceful stop (STA-4).</summary>
    public async Task<int> Wait(long runId, CancellationToken ct)
    {
        bool stopSent = false;
        DateTimeOffset lastLine = DateTimeOffset.MinValue;
        while (true)
        {
            if (ct.IsCancellationRequested && !stopSent)
            {
                stopSent = true;
                output.WriteLine(L.T("Зупинка: поточні файли завершуються…"));
                try
                {
                    await host.PostAsync($"api/runs/{runId}/stop");
                }
                catch (HostException)
                {
                }
            }
            var status = (await host.GetAsync($"api/runs/{runId}"))!;
            string state = status["run"]!["status"]!.GetValue<string>();
            bool active = status["active"]!.GetValue<bool>();
            if (DateTimeOffset.UtcNow - lastLine > TimeSpan.FromSeconds(5) || !active)
            {
                output.WriteLine(ProgressLine(status));
                lastLine = DateTimeOffset.UtcNow;
            }
            if (!active && state is not ("created" or "running" or "paused"))
            {
                output.WriteLine();
                output.Write(await host.GetTextAsync($"api/runs/{runId}/reports/summary"));
                return state switch
                {
                    "completed" => status["attention"]!.GetValue<long>() > 0 ? ExitCodes.Attention : ExitCodes.Ok,
                    "completed_with_errors" => ExitCodes.Attention,
                    _ => RunNotFinished,
                };
            }
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(1), stopSent ? CancellationToken.None : ct);
            }
            catch (OperationCanceledException)
            {
            }
        }
    }

    internal static string ProgressLine(JsonNode status)
    {
        var run = status["run"]!;
        long files = status["files"]!.GetValue<long>();
        long open = status["nonFinal"]!.GetValue<long>();
        long done = files - open;
        var counts = status["counts"]!.AsObject();
        long failed = counts.Where(c => c.Key.StartsWith("failed:", StringComparison.Ordinal)).Sum(c => c.Value!.GetValue<long>());
        long deferred = counts.TryGetPropertyValue("deferred", out var d) ? d!.GetValue<long>() : 0;
        string percent = files > 0 ? (100.0 * done / files).ToString("0", CultureInfo.InvariantCulture) + "%" : "—";
        string line = L.F("[{0}{1}] файлів {2}, готово {3} ({4}), замін {5}, помилок {6}, відкладено {7}, потребують уваги {8}",
            run["status"], status["paused"]!.GetValue<bool>() ? L.T(", пауза") : string.Empty, files, done, percent,
            status["replacements"], failed, deferred, status["attention"]);
        if (status["live"] is JsonObject live)
        {
            line += L.F(", {0:0} файл/хв", live["filesPerMinute"]!.GetValue<double>());
            foreach (var app in live["apps"]!.AsArray())
            {
                string paused = app!["pausedUntil"] is null ? string.Empty : L.T(" пауза");
                line += $" | {app["name"]}: {app["ruSpent"]} RU, 429/503 {app["throttled429"]}/{app["throttled503"]}{paused}";
            }
        }
        return line;
    }

    public async Task<int> Status(long runId)
    {
        var status = (await host.GetAsync($"api/runs/{runId}"))!;
        output.WriteLine(ProgressLine(status));
        if (status["run"]!["stopReason"] is { } reason)
            output.WriteLine(L.F("Причина зупинки: {0}", reason));
        foreach (var (code, n) in status["counts"]!.AsObject().OrderBy(c => c.Key, StringComparer.Ordinal))
            output.WriteLine($"  {code,-28} {n}");
        return ExitCodes.Ok;
    }

    public async Task<int> Runs(int take)
    {
        foreach (var r in (await host.GetAsync($"api/runs?take={take}"))!.AsArray())
        {
            output.WriteLine($"{r!["id"],6} {r["mode"],-10} {r["status"],-22} {r["profileName"],-20} {r["createdAt"]}" +
                             (r["parentId"] is { } p ? L.F("  (з прогону {0})", p) : string.Empty));
        }
        return ExitCodes.Ok;
    }

    public async Task<int> Files(long runId, string? status, string? errorClass, string? search, int take)
    {
        string query = $"take={take}";
        if (status is not null)
            query += "&status=" + Uri.EscapeDataString(status);
        if (errorClass is not null)
            query += "&errorClass=" + Uri.EscapeDataString(errorClass);
        if (search is not null)
            query += "&search=" + Uri.EscapeDataString(search);
        var page = (await host.GetAsync($"api/runs/{runId}/files?{query}"))!;
        foreach (var f in page["items"]!.AsArray())
            output.WriteLine($"{f!["id"],8} {f["statusCode"],-24} {f["path"]}{(f["error"] is { } e ? "  — " + e : string.Empty)}");
        output.WriteLine(L.F("Усього за фільтром: {0}", page["total"]));
        return ExitCodes.Ok;
    }

    /// <summary>Deletes a run that is not running, with everything it recorded; asks first unless --yes.</summary>
    public async Task<int> DeleteRun(long runId, bool yes, CancellationToken ct)
    {
        var status = (await host.GetAsync($"api/runs/{runId}", ct))!;
        var run = status["run"]!;
        long written = status["counts"]!.AsObject().TryGetPropertyValue("done", out var d) ? d!.GetValue<long>() : 0;
        string question = L.F("Видалити прогін {0} ({1}, {2}, файлів {3})?", runId, run["mode"], run["status"], status["files"]);
        if (written > 0)
            question += L.F(" Він записав у SharePoint {0} файлів: журнал цих змін і версії до запису (для відкату) буде втрачено.", written);
        if (!Confirm(question, yes))
            return ExitCodes.InvalidInput;
        await host.DeleteAsync($"api/runs/{runId}", ct);
        output.WriteLine(L.F("Прогін {0} видалено.", runId));
        return ExitCodes.Ok;
    }

    public async Task<int> Ignore(long fileId, string comment)
    {
        await host.PostAsync($"api/files/{fileId}/ignore", new { comment });
        output.WriteLine(L.F("Файл {0} позначено як ignored.", fileId));
        return ExitCodes.Ok;
    }

    /// <summary>REP-1…REP-3, RTY-8: all reports of a run into a folder.</summary>
    public async Task<int> Export(long runId, string directory, int? prefixDepth = null)
    {
        Directory.CreateDirectory(directory);
        // Same list as ReportExporter.Names on the host; sample and pilot are for the pilot (stage 5).
        foreach (string name in new[] { "files", "links", "residuals", "attention", "prefixes", "sample", "summary", "pilot" })
        {
            string file = Path.Combine(directory, name + (name is "summary" or "pilot" ? ".txt" : ".csv"));
            string query = name == "prefixes" && prefixDepth is { } depth ? $"?depth={depth}" : string.Empty;
            await using var source = await host.GetStreamAsync($"api/runs/{runId}/reports/{name}{query}");
            await using var target = File.Create(file);
            await source.CopyToAsync(target);
        }
        output.WriteLine(L.F("Звіти прогону {0}: {1}", runId, Path.GetFullPath(directory)));
        return ExitCodes.Ok;
    }

    private bool Confirm(string question, bool yes)
    {
        if (yes)
            return true;
        if (Console.IsInputRedirected)
        {
            output.WriteLine(question + L.T(" Для запуску без запитання додайте --yes."));
            return false;
        }
        output.Write(question + " [y/N] ");
        string? answer = Console.ReadLine();
        return answer is not null && (answer.Trim().Equals("y", StringComparison.OrdinalIgnoreCase) || answer.Trim().Equals("так", StringComparison.OrdinalIgnoreCase)); // i18n-ignore: a Ukrainian "yes" is accepted in any language
    }

    /// <summary>RTY-4: paths from a CSV — the "path" column, or the first column.</summary>
    internal static IEnumerable<string> ReadPathList(string file)
    {
        var lines = File.ReadAllLines(file).Where(l => l.Trim().Length > 0).ToList();
        if (lines.Count == 0)
            yield break;
        var header = SplitCsv(lines[0]);
        int column = header.FindIndex(h => h.Trim().Equals("path", StringComparison.OrdinalIgnoreCase));
        foreach (string line in column >= 0 ? lines.Skip(1) : lines)
        {
            var cells = SplitCsv(line);
            if (cells.Count > Math.Max(column, 0))
                yield return cells[Math.Max(column, 0)].Replace('\\', '/');
        }
    }

    private static List<string> SplitCsv(string line)
    {
        var cells = new List<string>();
        var sb = new System.Text.StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (quoted)
            {
                if (c == '"' && i + 1 < line.Length && line[i + 1] == '"')
                {
                    sb.Append('"');
                    i++;
                }
                else if (c == '"')
                {
                    quoted = false;
                }
                else
                {
                    sb.Append(c);
                }
            }
            else if (c == '"')
            {
                quoted = true;
            }
            else if (c is ',' or ';')
            {
                cells.Add(sb.ToString());
                sb.Clear();
            }
            else
            {
                sb.Append(c);
            }
        }
        cells.Add(sb.ToString());
        return cells;
    }
}

internal sealed record RetryArgs(
    string[] Statuses,
    string[] ErrorClasses,
    string[] Extensions,
    string? PathMask,
    long[] FileIds,
    string? ListFile,
    string? Mode,
    bool CurrentProfile);
