using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Rules;
using SPOReplaceLink.OpenXml;
using SPOReplaceLink.OpenXml.Tests;

namespace SPOReplaceLink.ExcelWorker.Tests;

/// <summary>
/// Stage 4 criterion on a real Excel: xls and xlsb fixed, a hung Excel killed, no EXCEL.EXE left.
/// Pass without checking anything where Excel is not installed.
/// </summary>
public sealed class RealExcelTests : IDisposable
{
    private const string Host = TestWorkbook.Host;
    private const string OldFolder = Host + "/Document/Fin";
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "spo-rl-realexcel", Guid.NewGuid().ToString("N")[..8]);

    public RealExcelTests() => Directory.CreateDirectory(_dir);

    private static bool ExcelInstalled => OperatingSystem.IsWindows() && Type.GetTypeFromProgID("Excel.Application") is not null;

    private ExcelPool Pool(double timeoutSeconds = 120) =>
        new(new ExcelPoolOptions(1, TimeSpan.FromSeconds(timeoutSeconds), Path.Combine(_dir, "state")));

    private static WorkbookProcessor Processor()
    {
        var build = RuleSet.Build(
        [
            new("/Document/", Host + "/sites/finance/"),
            new("/Document/Analytics/", Host + "/sites/finance/analytics_migrated/"),
        ], Host);
        return new WorkbookProcessor(build.RuleSet!, new PlacesOptions());
    }

    /// <summary>A real workbook of the given format made by Excel from the test package with links in every place.</summary>
    private async Task<string> MakeAsync(ExcelPool pool, string ext, int format)
    {
        string xlsx = TestWorkbook.Default().Save(Path.Combine(_dir, "src", "Report.xlsx"));
        string path = Path.Combine(_dir, "src", "Report" + ext);
        await pool.ConvertAsync(xlsx, path, format, CancellationToken.None);
        return path;
    }

    [Theory]
    [InlineData(".xls", XlFormat.Xls)]
    [InlineData(".xlsb", XlFormat.Xlsb)]
    public async Task ExcelFormatsAreFixedInTheirOwnFormat(string ext, int format)
    {
        if (!ExcelInstalled)
            return;
        await using var pool = Pool();
        string input = await MakeAsync(pool, ext, format);
        var stage = new ExcelStage(Processor(), pool);
        string output = Path.Combine(_dir, "out" + ext);

        var result = await stage.ProcessAsync(input, output, OldFolder, "Report" + ext, Path.Combine(_dir, "w1"), CancellationToken.None);

        Assert.Equal(FileStatus.Planned, result.Status);
        Assert.Empty(result.Residuals);
        Assert.True(result.ReplacementsOf(LinkKind.Hyperlink) >= 2);
        Assert.True(result.ReplacementsOf(LinkKind.Formula) >= 2);
        Assert.True(result.ReplacementsOf(LinkKind.LinkSource) >= 1);
        Assert.False(Directory.Exists(Path.Combine(_dir, "w1")));

        // What Excel sees in the result: the same format, nothing points to the old place.
        var seen = await pool.InspectAsync(output, CancellationToken.None);
        Assert.All(seen.Hyperlinks!.Concat(seen.LinkSources!).Concat(seen.Formulas!),
            link => Assert.DoesNotContain("/Document/", link, StringComparison.OrdinalIgnoreCase));
        Assert.Contains(seen.Hyperlinks!, h => h.StartsWith(Host + "/sites/finance/", StringComparison.Ordinal));
        Assert.Contains(seen.LinkSources!, s => s.Contains("analytics_migrated", StringComparison.Ordinal));
        Assert.Equal(format, (await pool.ConvertAsync(output, Path.Combine(_dir, "probe.xlsx"), XlFormat.Xlsx, CancellationToken.None)).SourceFormat);

        // A second pass finds nothing to do.
        var again = await stage.ProcessAsync(output, Path.Combine(_dir, "again" + ext), OldFolder, "Report" + ext, Path.Combine(_dir, "w2"), CancellationToken.None);
        Assert.Equal(FileStatus.Unchanged, again.Status);
        Assert.Empty(again.Residuals);
    }

    [Fact]
    public async Task LinkToTheDriveOfTheTempFolderIsNotWritten()
    {
        if (!ExcelInstalled)
            return;
        await using var pool = Pool();
        // An external book on the same drive as the temp folder that no rule moves: Excel would save it without the drive letter.
        string drive = Path.GetPathRoot(_dir)!.TrimEnd('\\');
        var wb = TestWorkbook.Default();
        wb.Set("xl/externalLinks/_rels/externalLink1.xml.rels", wb.Text("xl/externalLinks/_rels/externalLink1.xml.rels")
            .Replace("/Document/Analytics/Source.xlsx", $"file:///{drive}/Data/Source.xlsx"));
        string xlsx = wb.Save(Path.Combine(_dir, "src", "Local.xlsx"));
        string xls = Path.Combine(_dir, "src", "Local.xls");
        await pool.ConvertAsync(xlsx, xls, XlFormat.Xls, CancellationToken.None);
        string output = Path.Combine(_dir, "local-out.xls");

        var result = await new ExcelStage(Processor(), pool).ProcessAsync(xls, output, OldFolder, "Local.xls", Path.Combine(_dir, "w"), CancellationToken.None);

        Assert.Equal(FileStatus.Failed, result.Status);
        Assert.Equal(ErrorClass.Validation, result.ErrorClass);
        Assert.Contains(drive, result.Error, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task CachedValuesOfExternalLinksSurvive()
    {
        if (!ExcelInstalled)
            return;
        await using var pool = Pool();
        string input = await MakeAsync(pool, ".xls", XlFormat.Xls);
        string output = Path.Combine(_dir, "out.xls");
        await new ExcelStage(Processor(), pool).ProcessAsync(input, output, OldFolder, "Report.xls", Path.Combine(_dir, "w"), CancellationToken.None);

        // Back to xlsx to read the cell: A3 = [1]Sheet1!A1 keeps its cached 42 (Excel's ChangeLink would make it #REF!).
        string probe = Path.Combine(_dir, "probe.xlsx");
        await pool.ConvertAsync(output, probe, XlFormat.Xlsx, CancellationToken.None);
        string sheet = TestWorkbook.ReadPart(probe, "xl/worksheets/sheet1.xml");
        Assert.Matches("<c r=\"A3\"[^>]*>.*?<v>42</v>", sheet);
    }

    [Fact]
    public async Task FileThatIsNotAWorkbookIsAFileError()
    {
        if (!ExcelInstalled)
            return;
        await using var pool = Pool();
        // Excel would open text as a worksheet; the worker refuses anything that is not a workbook format.
        string path = Path.Combine(_dir, "garbage.xls");
        await File.WriteAllTextAsync(path, "not a workbook");
        var result = await new ExcelStage(Processor(), pool).ProcessAsync(path, Path.Combine(_dir, "o.xls"), OldFolder, "garbage.xls", Path.Combine(_dir, "w"), CancellationToken.None);
        Assert.Equal(FileStatus.Failed, result.Status);
        Assert.Equal(ErrorClass.File, result.ErrorClass);
    }

    [Fact]
    public async Task PasswordPromptIsClosedAndExcelRestarted()
    {
        if (!ExcelInstalled)
            return;
        await using var pool = Pool();
        string path = Path.Combine(_dir, "book.xls");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "fixtures", "password-open.xls"), path);

        // The prompt some Excel versions show although a password is passed; here it is provoked by passing none.
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.SendAsync(new WorkerRequest("prompt", path), CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Password, ex.Kind);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(30), sw.Elapsed.ToString());
        Assert.Empty(pool.Processes);
        Assert.Equal(1, pool.Status.Dialogs);

        // The usual open of the same file is a password error without a dialog, and the worker is kept.
        var plain = await Assert.ThrowsAsync<ExcelException>(() => pool.InspectAsync(path, CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Password, plain.Kind);
        Assert.Single(pool.Processes);
    }

    [Fact]
    public async Task HungExcelIsKilledWithItsWorker()
    {
        if (!ExcelInstalled)
            return;
        await using var pool = Pool(timeoutSeconds: 3);
        await MakeAsync(pool, ".xls", XlFormat.Xls);
        var (worker, excel) = pool.Processes.Single();
        Assert.NotNull(excel);

        var ex = await Assert.ThrowsAsync<ExcelException>(() => pool.SendAsync(new WorkerRequest("stall"), CancellationToken.None));
        Assert.Equal(ExcelErrorKind.Timeout, ex.Kind);
        Assert.True(ExcelPoolTests.Exited(worker));
        Assert.True(ExcelPoolTests.Exited(excel!.Value));
    }

    [Fact]
    public async Task DisposeLeavesNoExcel()
    {
        if (!ExcelInstalled)
            return;
        var pool = Pool();
        await MakeAsync(pool, ".xls", XlFormat.Xls);
        var (worker, excel) = pool.Processes.Single();
        await pool.DisposeAsync();
        Assert.True(ExcelPoolTests.Exited(worker));
        Assert.True(ExcelPoolTests.Exited(excel!.Value));
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
