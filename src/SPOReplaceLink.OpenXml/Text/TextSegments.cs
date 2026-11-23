using System.Text;

namespace SPOReplaceLink.OpenXml.Text;

/// <summary>
/// Reads a large XML part in pieces cut at safe places, so no single string has to hold the whole part
/// (.NET strings stop at ~1G chars, and a sheet of a few million rows is bigger than that).
/// </summary>
internal static class TextSegments
{
    /// <summary>
    /// Pieces of about <paramref name="size"/> chars; each ends where <paramref name="cut"/> says (an index after the
    /// last safe place in the text, or ≤ 0 when there is none yet and more text is needed). The last piece is the rest.
    /// </summary>
    public static IEnumerable<string> Read(TextReader reader, int size, Func<string, int> cut)
    {
        var buffer = new char[size];
        var carry = new StringBuilder();
        while (true)
        {
            int n = reader.ReadBlock(buffer, 0, buffer.Length);
            if (n == 0)
            {
                if (carry.Length > 0)
                    yield return carry.ToString();
                yield break;
            }
            carry.Append(buffer, 0, n);
            if (carry.Length < size)
                continue;
            string text = carry.ToString();
            int at = cut(text);
            if (at <= 0)
                continue;
            yield return at == text.Length ? text : text[..at];
            carry.Clear().Append(text, at, text.Length - at);
        }
    }

    /// <summary>
    /// After the last '>': prefixes never contain '>' (it is escaped in XML text), and the char before a piece is '>'
    /// as it would be in the whole text, so the residual scan finds the same matches piece by piece.
    /// </summary>
    public static int AfterLastTag(string text) => text.LastIndexOf('>') + 1;

    /// <summary>After the last cell end tag (&lt;/c&gt; with any prefix): &lt;f&gt; and &lt;v&gt; never cross it.</summary>
    public static int AfterLastCellEnd(string text)
    {
        for (int i = text.LastIndexOf("c>", StringComparison.Ordinal); i > 0; i = text.LastIndexOf("c>", i - 1, StringComparison.Ordinal))
        {
            if (i >= 2 && text[i - 1] == '/' && text[i - 2] == '<')
                return i + 2;
            if (text[i - 1] == ':')
            {
                int k = i - 2;
                while (k >= 0 && (char.IsLetterOrDigit(text[k]) || text[k] is '_' or '.' or '-'))
                    k--;
                if (k >= 1 && k < i - 2 && text[k] == '/' && text[k - 1] == '<')
                    return i + 2;
            }
        }
        return -1;
    }
}
