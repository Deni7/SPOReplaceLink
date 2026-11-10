using SPOReplaceLink.OpenXml.Parts;
using SPOReplaceLink.OpenXml.Text;

namespace SPOReplaceLink.OpenXml.Tests;

public class FormulaLiteralTests
{
    private static string[] Literals(string formula) =>
        PartRewriters.StringLiterals(formula).Select(r => formula.Substring(r.Start, r.Length)).ToArray();

    [Fact]
    public void FindsLiterals() =>
        Assert.Equal(["a", "b\"\"c", ""], Literals("CONCAT(\"a\",\"b\"\"c\",\"\")"));

    [Fact]
    public void SkipsQuotedSheetNames() =>
        Assert.Equal(["x"], Literals("'My \"odd\" ''sheet'''!A1&\"x\""));

    [Fact]
    public void UnterminatedLiteralRunsToEnd() => Assert.Equal(["abc"], Literals("\"abc"));

    [Fact]
    public void UnescapesWithOffsets()
    {
        string source = "a&amp;b&#x417;&lt;";
        string value = XmlText.Unescape(source, out int[] offsets);
        Assert.Equal("a&bЗ<", value);
        Assert.Equal([0, 1, 6, 7, 14, 18], offsets);
    }

    [Fact]
    public void UnknownEntityStaysAsText() => Assert.Equal("&nbsp;x", XmlText.Unescape("&nbsp;x"));
}
