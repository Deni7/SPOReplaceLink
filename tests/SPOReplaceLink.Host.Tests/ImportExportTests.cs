using System.Text;
using SPOReplaceLink.Engine;
using SPOReplaceLink.OpenXml;

namespace SPOReplaceLink.Host.Tests;

public class ImportExportTests
{
    private static RuleImport.Result Parse(string name, string text) =>
        RuleImport.Parse(name, new MemoryStream(Encoding.UTF8.GetBytes(text)));

    [Theory]
    [InlineData("old,new\r\nhttps://h/A,https://h/B\r\n")]
    [InlineData("https://h/A;https://h/B")]
    [InlineData("https://h/A\thttps://h/B")]
    [InlineData("https://h/A => https://h/B")]
    [InlineData("https://h/A -> https://h/B")]
    [InlineData("https://h/A → https://h/B")]
    [InlineData("Старий шлях|Новий шлях\nhttps://h/A|https://h/B")]
    public void ReadsMappingTables(string text)
    {
        var result = Parse("rules.csv", text);
        var rule = Assert.Single(result.Rules);
        Assert.Equal("https://h/A", rule.Old);
        Assert.Equal("https://h/B", rule.New);
    }

    [Fact]
    public void ReportsBadLines()
    {
        var result = Parse("rules.txt", "https://h/A,https://h/B\nonly-one-value\n\n/x/,/y/,з назвою\n");
        Assert.Equal(2, result.Rules.Count);
        Assert.Equal("з назвою", result.Rules[1].Name);
        Assert.Single(result.Warnings);
    }

    [Fact]
    public void QuotedCsvWithCommas()
    {
        var result = Parse("rules.csv", "\"https://h/Fin, 2024\",\"https://h/new\"");
        Assert.Equal("https://h/Fin, 2024", Assert.Single(result.Rules).Old);
    }

    [Fact]
    public void XlsxRoundTrip()
    {
        using var ms = new MemoryStream();
        SimpleXlsxWriter.Write(ms, "rules", [["Старий", "Новий"], ["https://h/Звіт", "https://h/sites/x"], ["/a/", "/b/"], ["42", "x&<y>"]]);
        ms.Position = 0;
        var rows = SimpleXlsxReader.ReadFirstSheet(ms);
        Assert.Equal(4, rows.Count);
        Assert.Equal("https://h/Звіт", rows[1][0]);
        Assert.Equal("x&<y>", rows[3][1]);
        Assert.Equal("42", rows[3][0]);

        ms.Position = 0;
        var imported = RuleImport.Parse("map.xlsx", ms);
        Assert.Equal(3, imported.Rules.Count);
    }

    [Fact]
    public void TemplateCsvExportImportsBack()
    {
        var csv = RuleImport.ToCsv([new() { Old = "https://h/A", New = "https://h/B", Name = "n, 1" }]);
        var result = RuleImport.Parse("t.csv", new MemoryStream(csv));
        var rule = Assert.Single(result.Rules);
        Assert.Equal("n, 1", rule.Name);
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    public void ColumnNames(int index, string name) => Assert.Equal(name, SimpleXlsxWriter.Column(index));
}
