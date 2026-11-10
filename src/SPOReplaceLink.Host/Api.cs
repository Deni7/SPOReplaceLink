using Microsoft.AspNetCore.Diagnostics;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.Engine;
using SPOReplaceLink.Ui;
using SPOReplaceLink.Ui.Services;
using SPOReplaceLink.State;

namespace SPOReplaceLink.Host;

public sealed record CreateRunRequest(string Mode, string? Profile, bool Start = true);

public sealed record ChildRunRequest(
    List<string>? Statuses = null,
    List<string>? ErrorClasses = null,
    List<string>? Extensions = null,
    string? PathMask = null,
    List<string>? Paths = null,
    List<long>? FileIds = null,
    string? Mode = null,
    bool CurrentProfile = false,
    bool Start = true)
{
    public FileFilter ToFilter() => new()
    {
        Statuses = Statuses ?? [],
        ErrorClasses = ErrorClasses ?? [],
        Extensions = Extensions ?? [],
        PathMask = PathMask,
        Paths = Paths ?? [],
        FileIds = FileIds ?? [],
    };
}

public sealed record IgnoreRequest(string Comment);

public sealed record IssueDto(string Severity, string Message, int? Rule)
{
    public static List<IssueDto> From(IEnumerable<ValidationIssue> issues) =>
        issues.Select(i => new IssueDto(i.Severity == IssueSeverity.Error ? "error" : "warning", i.Message, i.RuleIndex)).ToList();
}

public sealed record FileDetails(FileRow File, IReadOnlyList<Core.Processing.LinkRecord> Links, IReadOnlyList<Core.Processing.Residual> Residuals,
    IReadOnlyList<AttemptRow> Attempts);

internal static class Api
{
    private static readonly Dictionary<string, string> Kinds = new(StringComparer.OrdinalIgnoreCase)
    {
        ["profiles"] = SettingKinds.Profile,
        ["connections"] = SettingKinds.Connection,
        ["templates"] = SettingKinds.Template,
    };

    public static void MapApi(this WebApplication app)
    {
        var api = app.MapGroup("/api");
        api.MapGet("/health", () => Results.Ok(new { status = "ok", version = typeof(Api).Assembly.GetName().Version?.ToString() }));

        // ---- settings (CFG-1, GUI-16, GUI-22, GUI-26)
        api.MapGet("/{kind}", (string kind, RunService runs) =>
            Results.Ok(runs.Store.ListSettings(Kind(kind)).Select(s => new { s.Name, s.UpdatedAt })));
        api.MapGet("/{kind}/{name}", (string kind, string name, bool? download, RunService runs) =>
            runs.Store.GetSetting(Kind(kind), name) is { } json
                ? download == true
                    ? Results.File(System.Text.Encoding.UTF8.GetBytes(json), "application/json", $"{name}.{Kind(kind)}.json")
                    : Results.Text(json, "application/json")
                : Results.NotFound());
        api.MapPut("/{kind}/{name}", async (string kind, string name, HttpRequest request, SettingsService settings) =>
        {
            using var reader = new StreamReader(request.Body);
            var issues = settings.SaveJson(Kind(kind), name, await reader.ReadToEndAsync());
            return Results.Ok(new { name, issues = IssueDto.From(issues) });
        });
        api.MapGet("/templates/{name}/csv", (string name, SettingsService settings) =>
            settings.GetTemplate(name) is { } t ? Results.File(RuleImport.ToCsv(t.Rules), "text/csv; charset=utf-8", $"{name}.rules.csv") : Results.NotFound());        api.MapDelete("/{kind}/{name}", (string kind, string name, SettingsService settings) =>
            settings.Delete(Kind(kind), name) ? Results.NoContent() : Results.NotFound());
        api.MapPost("/profiles/{name}/validate", (string name, RunService runs) =>
        {
            var v = runs.ValidateProfile(runs.GetProfile(name));
            return Results.Ok(new { valid = v.IsValid, issues = IssueDto.From(v.Issues) });
        });
        api.MapPost("/connections/{name}/test", async (string name, string? profile, RunService runs, CancellationToken ct) =>
            Results.Ok(await runs.TestConnectionAsync(name, profile, ct)));

        // ---- runs (FEATURES 7, STA-4, STA-5)
        api.MapGet("/runs", (RunService runs, int? take) => Results.Ok(runs.Store.ListRuns(take ?? 50).Select(RunInfo.From)));
        api.MapPost("/runs", (CreateRunRequest request, RunService runs) =>
        {
            long id = runs.CreateRun(request.Mode, request.Profile);
            if (request.Start)
                runs.Start(id);
            return Results.Created($"/api/runs/{id}", new { id });
        });
        api.MapGet("/runs/{id:long}", (long id, StatusService status) => status.Get(id, fresh: true) is { } s ? Results.Ok(s) : Results.NotFound());
        api.MapPost("/runs/{id:long}/start", (long id, RunService runs) => Ok(() => runs.Start(id)));
        api.MapPost("/runs/{id:long}/pause", (long id, RunService runs) => Ok(() => runs.Pause(id)));
        api.MapPost("/runs/{id:long}/resume", (long id, RunService runs) => Ok(() => runs.Resume(id)));
        api.MapPost("/runs/{id:long}/stop", (long id, RunService runs) => Ok(() => runs.Stop(id)));
        api.MapDelete("/runs/{id:long}", (long id, RunService runs) => Ok(() => runs.DeleteRun(id)));
        api.MapPost("/runs/{id:long}/retry", (long id, ChildRunRequest request, RunService runs) => Child(runs, id, request, request.Mode));
        api.MapPost("/runs/{id:long}/verify", (long id, RunService runs) =>
            Child(runs, id, new ChildRunRequest(Statuses: [FileStatuses.Done]), RunModes.Verify));
        api.MapGet("/runs/{id:long}/files", (long id, RunService runs, string? status, string? errorClass, string? ext, string? search, int? target, int? skip, int? take) =>
        {
            var filter = Filter(status, errorClass, ext, search, target);
            return Results.Ok(new
            {
                total = runs.Store.CountFiles(id, filter),
                items = runs.Store.QueryFiles(id, filter, skip ?? 0, Math.Clamp(take ?? 100, 1, 1000)),
            });
        });
        api.MapGet("/runs/{id:long}/metrics", (long id, RunService runs) =>
            Results.Ok(runs.Store.GetMetrics(id).Select(m => new { at = m.At, data = System.Text.Json.JsonDocument.Parse(m.Json).RootElement })));
        api.MapGet("/runs/{id:long}/reports/{name}", (long id, string name, string? format, int? depth, RunService runs) =>
        {
            if (!ReportExporter.Names.Contains(name) || runs.Store.GetRun(id) is null)
                return Results.NotFound();
            if (depth is < 1 or > ReportExporter.MaxPrefixDepth)
                return Results.BadRequest(L.F("depth: від 1 до {0}.", ReportExporter.MaxPrefixDepth));
            var ms = new MemoryStream();
            Exports.WriteReport(runs.Store, id, name, format ?? "csv", ms, depth ?? ReportExporter.DefaultPrefixDepth);
            ms.Position = 0;
            string type = format == "xlsx" ? Exports.XlsxType : ReportExporter.IsText(name) ? "text/plain; charset=utf-8" : "text/csv; charset=utf-8";
            return Results.File(ms, type, Exports.ReportFileName(id, name, format ?? "csv"));
        });
        // GUI-10: the files table with the same filter as on screen.
        api.MapGet("/runs/{id:long}/files/export", (long id, RunService runs, string? format, string? status, string? errorClass, string? ext, string? search, int? target) =>
        {
            var ms = new MemoryStream();
            Exports.WriteFiles(runs.Store, id, Filter(status, errorClass, ext, search, target), format ?? "csv", ms);
            ms.Position = 0;
            return format == "xlsx"
                ? Results.File(ms, Exports.XlsxType, $"run-{id}-files.xlsx")
                : Results.File(ms, "text/csv; charset=utf-8", $"run-{id}-files.csv");
        });
        // ---- files (GUI-6, RTY-6)
        api.MapGet("/files/{id:long}", (long id, RunService runs) =>
            runs.Store.GetFile(id) is { } f
                ? Results.Ok(new FileDetails(f, runs.Store.GetLinks(id), runs.Store.GetResiduals(id), runs.Store.GetAttempts(id)))
                : Results.NotFound());
        api.MapPost("/files/{id:long}/ignore", (long id, IgnoreRequest request, RunService runs) =>
            string.IsNullOrWhiteSpace(request.Comment)
                ? Results.BadRequest(new { error = L.T("Потрібен коментар: чому файл ігнорується.") })
                : runs.Store.Ignore(id, request.Comment, DateTimeOffset.UtcNow) ? Results.NoContent() : Results.NotFound());
    }

    private static IResult Child(RunService runs, long parent, ChildRunRequest request, string? mode)
    {
        long id = runs.CreateChildRun(parent, request.ToFilter(), mode, request.CurrentProfile);
        var run = runs.Store.GetRun(id)!;
        if (request.Start && RunStatuses.CanContinue(run.Status))
            runs.Start(id);
        return Results.Created($"/api/runs/{id}", new { id, files = runs.Store.CountFiles(id, new FileFilter()) });
    }

    public static FileFilter Filter(string? status, string? errorClass, string? ext, string? search, int? target) => new()
    {
        Statuses = Split(status),
        ErrorClasses = Split(errorClass),
        Extensions = Split(ext),
        Search = search,
        TargetIndex = target,
    };

    private static IResult Ok(Action action)
    {
        action();
        return Results.NoContent();
    }

    private static string Kind(string kind) =>
        Kinds.TryGetValue(kind, out var k) ? k : throw new KeyNotFoundException(L.F("Невідомий розділ «{0}».", kind));

    private static List<string> Split(string? value) =>
        string.IsNullOrWhiteSpace(value) ? [] : value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
}

internal static class ApiErrors
{
    public static async Task Handle(HttpContext context)
    {
        var error = context.Features.Get<IExceptionHandlerFeature>()?.Error;
        (int status, object body) = error switch
        {
            RunValidationException v => (400, (object)new { error = v.Message, issues = IssueDto.From(v.Issues) }),
            ProfileException or ArgumentException or BadHttpRequestException => (400, new { error = error.Message }),
            KeyNotFoundException => (404, new { error = error.Message }),
            InvalidOperationException => (409, new { error = error.Message }),
            _ => (500, new { error = error?.Message ?? L.T("Внутрішня помилка") }),
        };
        context.Response.StatusCode = status;
        await context.Response.WriteAsJsonAsync(body);
    }
}
