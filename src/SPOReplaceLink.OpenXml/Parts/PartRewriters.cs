using System.Text;
using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Urls;
using SPOReplaceLink.OpenXml.Text;

namespace SPOReplaceLink.OpenXml.Parts;

/// <summary>
/// Rewriters for the parts that hold links. Each one edits only attribute values or element text
/// that actually change; markup, namespaces and whitespace stay byte-identical (OOX-11).
/// </summary>
internal static partial class PartRewriters
{
    [GeneratedRegex(@"<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*>", RegexOptions.CultureInvariant)]
    private static partial Regex RelationshipTag();

    [GeneratedRegex(@"(?<name>[A-Za-z_][\w.:-]*)\s*=\s*(?:""(?<dq>[^""]*)""|'(?<sq>[^']*)')", RegexOptions.CultureInvariant)]
    private static partial Regex Attribute();

    // <f ...>text</f> and <v>text</v> with any namespace prefix; self-closing shared formulas are skipped.
    [GeneratedRegex(@"<(?:[A-Za-z_][\w.-]*:)?(?<tag>[fv])(?:\s[^>]*)?(?<!/)>(?<text>[^<]*)</(?:[A-Za-z_][\w.-]*:)?\k<tag>>", RegexOptions.CultureInvariant)]
    private static partial Regex FormulaOrValue();

    [GeneratedRegex(@"https?://|file:|\\\\", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex EmbeddedLinkStart();

    [GeneratedRegex(@"<(?:[A-Za-z_][\w.-]*:)?absPath\b[^>]*>", RegexOptions.CultureInvariant)]
    private static partial Regex AbsPathTag();

    /// <summary>x15ac:absPath/@url in workbook.xml — where the file was last saved.</summary>
    public static string RewriteWorkbookPath(string xml, LinkContext ctx)
    {
        var edits = new TextEdits();
        foreach (var (index, length, quote) in AbsPathUrls(xml))
            ReplaceAttributeValue(edits, ctx, LinkKind.WorkbookPath, index, length, quote, xml);
        return edits.Apply(xml);
    }

    /// <summary>The same text with absPath values blanked, for the residual scan.</summary>
    public static string MaskWorkbookPath(string xml)
    {
        var chars = xml.ToCharArray();
        foreach (var (index, length, _) in AbsPathUrls(xml))
            Array.Fill(chars, ' ', index, length);
        return new string(chars);
    }

    private static IEnumerable<(int Index, int Length, char Quote)> AbsPathUrls(string xml)
    {
        foreach (Match tag in AbsPathTag().Matches(xml))
        {
            foreach (Match attr in Attribute().Matches(tag.Value))
            {
                if (attr.Groups["name"].Value != "url")
                    continue;
                var value = ValueGroup(attr);
                yield return (tag.Index + value.Index, value.Length, tag.Value[value.Index - 1]);
            }
        }
    }

    /// <summary>OOX-1, OOX-2, XLB-1: external targets in any .rels part.</summary>
    public static string RewriteRels(string xml, LinkContext ctx)
    {
        bool externalLinkRels = ctx.Part.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase);
        var edits = new TextEdits();
        foreach (Match tag in RelationshipTag().Matches(xml))
        {
            Group? target = null;
            string? type = null;
            bool external = false;
            foreach (Match attr in Attribute().Matches(tag.Value))
            {
                string name = LocalName(attr.Groups["name"].Value);
                var value = ValueGroup(attr);
                switch (name)
                {
                    case "TargetMode":
                        external = value.Value == "External";
                        break;
                    case "Target":
                        target = value;
                        break;
                    case "Type":
                        type = value.Value;
                        break;
                }
            }
            if (!external || target is null)
                continue;

            var kind = externalLinkRels ? LinkKind.LinkSource
                : type?.EndsWith("/hyperlink", StringComparison.Ordinal) == true ? LinkKind.Hyperlink
                : LinkKind.OtherExternal;
            ReplaceAttributeValue(edits, ctx, kind, tag.Index + target.Index, target.Length, tag.Value[target.Index - 1], xml);
        }
        return edits.Apply(xml);
    }

    /// <summary>OOX-3: absolute locations written as attributes in externalLinkN.xml (alternate URLs).</summary>
    public static string RewriteExternalLink(string xml, LinkContext ctx)
    {
        var edits = new TextEdits();
        foreach (Match attr in Attribute().Matches(xml))
        {
            string name = attr.Groups["name"].Value;
            if (name == "xmlns" || name.StartsWith("xmlns:", StringComparison.Ordinal) || LocalName(name) == "Ignorable")
                continue;
            var value = ValueGroup(attr);
            if (!LinkClassifier.IsAbsolute(LinkClassifier.Classify(XmlText.Unescape(value.Value))))
                continue;
            ReplaceAttributeValue(edits, ctx, LinkKind.LinkSource, value.Index, value.Length, xml[value.Index - 1], xml);
        }
        return edits.Apply(xml);
    }

    /// <summary>OOX-4: text URLs in formula string literals, and cached string results that are links.</summary>
    public static string RewriteSheet(string xml, LinkContext ctx)
    {
        var edits = new TextEdits();
        foreach (Match m in FormulaOrValue().Matches(xml))
        {
            var text = m.Groups["text"];
            if (text.Length == 0)
                continue;
            string? replaced = m.Groups["tag"].Value == "f"
                ? RewriteFormula(xml.Substring(text.Index, text.Length), ctx)
                : RewriteCachedValue(xml.Substring(text.Index, text.Length), ctx);
            if (replaced is not null)
                edits.Replace(text.Index, text.Length, replaced);
        }
        return edits.Apply(xml);
    }

    private static string? RewriteCachedValue(string escaped, LinkContext ctx)
    {
        string value = XmlText.Unescape(escaped);
        var form = LinkClassifier.Classify(value);
        if (!LinkClassifier.IsAbsolute(form) && form != LinkForm.ServerRelative)
            return null;
        string? replaced = ctx.Rewrite(value, LinkKind.CachedValue);
        return replaced is null ? null : XmlText.EscapeText(replaced);
    }

    /// <summary>Rewrites links inside string literals of one formula; returns the escaped text or null.</summary>
    internal static string? RewriteFormula(string escaped, LinkContext ctx)
    {
        string formula = XmlText.Unescape(escaped, out int[] offsets);
        var edits = new TextEdits();
        foreach (var (start, length) in StringLiterals(formula))
        {
            string literal = formula.Substring(start, length).Replace("\"\"", "\"");
            string? replaced = RewriteLiteral(literal, ctx);
            if (replaced is null)
                continue;
            string newEscaped = XmlText.EscapeText(replaced.Replace("\"", "\"\""));
            int sourceStart = offsets[start];
            edits.Replace(sourceStart, offsets[start + length] - sourceStart, newEscaped);
        }
        return edits.Count == 0 ? null : edits.Apply(escaped);
    }

    private static string? RewriteLiteral(string literal, LinkContext ctx)
    {
        var form = LinkClassifier.Classify(literal);
        if (LinkClassifier.IsAbsolute(form) || form == LinkForm.ServerRelative)
            return ctx.Rewrite(literal, LinkKind.Formula);

        // A link inside a longer text: "see https://…/a.xlsx". It ends at whitespace.
        StringBuilder? sb = null;
        int copied = 0;
        foreach (Match start in EmbeddedLinkStart().Matches(literal))
        {
            if (start.Index < copied)
                continue;
            int end = start.Index;
            while (end < literal.Length && !char.IsWhiteSpace(literal[end]))
                end++;
            string? replaced = ctx.Rewrite(literal[start.Index..end], LinkKind.Formula);
            if (replaced is null)
                continue;
            sb ??= new StringBuilder(literal.Length + 32);
            sb.Append(literal, copied, start.Index - copied).Append(replaced);
            copied = end;
        }
        if (sb is null)
            return null;
        return sb.Append(literal, copied, literal.Length - copied).ToString();
    }

    /// <summary>Ranges of string literal contents (without the quotes) in a formula; "" is an escaped quote.</summary>
    internal static IEnumerable<(int Start, int Length)> StringLiterals(string formula)
    {
        int i = 0;
        while (i < formula.Length)
        {
            if (formula[i] == '\'')
            {
                // Quoted sheet name: 'My Sheet'!A1 — skip it, '' is an escaped apostrophe.
                i++;
                while (i < formula.Length && !(formula[i] == '\'' && (i + 1 >= formula.Length || formula[i + 1] != '\'')))
                    i += formula[i] == '\'' ? 2 : 1;
                i++;
                continue;
            }
            if (formula[i] != '"')
            {
                i++;
                continue;
            }
            int start = ++i;
            while (i < formula.Length)
            {
                if (formula[i] == '"')
                {
                    if (i + 1 < formula.Length && formula[i + 1] == '"')
                    {
                        i += 2;
                        continue;
                    }
                    break;
                }
                i++;
            }
            yield return (start, i - start);
            i++;
        }
    }

    private static void ReplaceAttributeValue(TextEdits edits, LinkContext ctx, LinkKind kind, int index, int length, char quote, string xml)
    {
        string value = XmlText.Unescape(xml.Substring(index, length));
        string? replaced = ctx.Rewrite(value, kind);
        if (replaced is not null)
            edits.Replace(index, length, XmlText.EscapeAttribute(replaced, quote));
    }

    private static Group ValueGroup(Match attr) => attr.Groups["dq"].Success ? attr.Groups["dq"] : attr.Groups["sq"];

    private static string LocalName(string name)
    {
        int colon = name.IndexOf(':');
        return colon < 0 ? name : name[(colon + 1)..];
    }
}
