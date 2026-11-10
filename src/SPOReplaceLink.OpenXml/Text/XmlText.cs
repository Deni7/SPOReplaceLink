using System.Globalization;
using System.Text;

namespace SPOReplaceLink.OpenXml.Text;

/// <summary>
/// XML escaping that keeps a map back to the source, so a single value can be replaced in place
/// without re-serializing the rest of the part (OOX-11).
/// </summary>
internal static class XmlText
{
    /// <param name="offsets">offsets[i] — index in <paramref name="source"/> where unescaped char i starts;
    /// offsets[result.Length] == source.Length.</param>
    public static string Unescape(string source, out int[] offsets)
    {
        var sb = new StringBuilder(source.Length);
        var map = new List<int>(source.Length + 1);
        int i = 0;
        while (i < source.Length)
        {
            if (source[i] == '&')
            {
                int semi = source.IndexOf(';', i + 1);
                if (semi > i && TryEntity(source.AsSpan(i + 1, semi - i - 1), out string? value))
                {
                    foreach (char c in value!)
                    {
                        map.Add(i);
                        sb.Append(c);
                    }
                    i = semi + 1;
                    continue;
                }
            }
            map.Add(i);
            sb.Append(source[i]);
            i++;
        }
        map.Add(source.Length);
        offsets = map.ToArray();
        return sb.ToString();
    }

    public static string Unescape(string source) => Unescape(source, out _);

    public static string EscapeText(string value) =>
        value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    public static string EscapeAttribute(string value, char quote)
    {
        string s = EscapeText(value);
        return quote == '"' ? s.Replace("\"", "&quot;") : s.Replace("'", "&apos;");
    }

    private static bool TryEntity(ReadOnlySpan<char> name, out string? value)
    {
        value = name switch
        {
            "amp" => "&",
            "lt" => "<",
            "gt" => ">",
            "quot" => "\"",
            "apos" => "'",
            _ => null,
        };
        if (value is not null)
            return true;

        if (name.Length < 2 || name[0] != '#')
            return false;
        bool hex = name[1] is 'x' or 'X';
        var digits = hex ? name[2..] : name[1..];
        if (!int.TryParse(digits, hex ? NumberStyles.AllowHexSpecifier : NumberStyles.None, CultureInfo.InvariantCulture, out int code)
            || !Rune.IsValid(code))
            return false;
        value = char.ConvertFromUtf32(code);
        return true;
    }
}
