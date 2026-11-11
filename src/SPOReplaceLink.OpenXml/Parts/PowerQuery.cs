using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.OpenXml.Text;

namespace SPOReplaceLink.OpenXml.Parts;

/// <summary>
/// OOX-8: Power Query formulas. Excel keeps them in customXml/itemN.xml as &lt;DataMashup&gt; with base64 of the
/// [MS-QDEFF] structure: version, then four length-prefixed blocks — package parts (a zip with Formulas/Section1.m),
/// permissions, metadata and permission bindings. Only the zip changes, and in it only the .m entries with
/// replacements; the other blocks are copied as they are.
/// </summary>
/// <remarks>
/// Permission bindings are a checksum of the package and the permissions, protected by DPAPI for the user who saved
/// the file. After the formulas change it no longer matches, and Excel falls back to the default query permissions
/// (as it does for any other user who opens the file): privacy levels may be asked again on the next refresh.
/// </remarks>
internal static partial class PowerQuery
{
    [GeneratedRegex(@"<(?:[A-Za-z_][\w.-]*:)?DataMashup\b[^>]*>(?<b64>[^<]*)</(?:[A-Za-z_][\w.-]*:)?DataMashup>", RegexOptions.CultureInvariant)]
    private static partial Regex Element();

    public static bool IsMashup(string xml) => Element().IsMatch(xml);

    /// <summary>Rewrites links in text literals of the M formulas; returns the part text, unchanged if there is nothing to replace.</summary>
    public static string Rewrite(string xml, LinkContext ctx)
    {
        var edits = new TextEdits();
        foreach (Match m in Element().Matches(xml))
        {
            var b64 = m.Groups["b64"];
            var mashup = Mashup.Parse(Decode(b64.Value));
            if (RewritePackage(mashup.PackageParts, ctx) is not { } package)
                continue;
            edits.Replace(b64.Index, b64.Length, Convert.ToBase64String((mashup with { PackageParts = package }).ToBytes()));
        }
        return edits.Apply(xml);
    }

    /// <summary>Decoded texts of the mashup for the residual scan: the formulas, permissions and metadata.</summary>
    public static IEnumerable<(string Name, string Text)> Contents(string xml)
    {
        foreach (Match m in Element().Matches(xml))
        {
            var mashup = Mashup.Parse(Decode(m.Groups["b64"].Value));
            foreach (var (name, bytes) in PackageEntries(mashup.PackageParts))
                yield return (name, PartText.Decode(bytes).Text);
            yield return ("Permissions", Lenient(mashup.Permissions));
            yield return ("Metadata", Lenient(mashup.Metadata));
        }
    }

    /// <summary>Ranges of text literal contents (without the quotes) in M code; comments and #"quoted identifiers" are skipped.</summary>
    internal static IEnumerable<(int Start, int Length)> TextLiterals(string m)
    {
        int i = 0;
        while (i < m.Length)
        {
            char c = m[i];
            char next = i + 1 < m.Length ? m[i + 1] : '\0';
            if (c == '/' && next == '/')
            {
                int eol = m.IndexOfAny(['\r', '\n'], i);
                i = eol < 0 ? m.Length : eol;
                continue;
            }
            if (c == '/' && next == '*')
            {
                int close = m.IndexOf("*/", i + 2, StringComparison.Ordinal);
                i = close < 0 ? m.Length : close + 2;
                continue;
            }
            bool identifier = c == '#' && next == '"';
            if (!identifier && c != '"')
            {
                i++;
                continue;
            }
            int start = identifier ? i + 2 : i + 1;
            int end = start;
            while (end < m.Length && !(m[end] == '"' && (end + 1 >= m.Length || m[end + 1] != '"')))
                end += m[end] == '"' ? 2 : 1;
            if (!identifier)
                yield return (start, Math.Min(end, m.Length) - start);
            i = end + 1;
        }
    }

    /// <summary>Links in the text literals of one M document; returns the new text or null.</summary>
    internal static string? RewriteFormulas(string m, LinkContext ctx)
    {
        var edits = new TextEdits();
        foreach (var (start, length) in TextLiterals(m))
        {
            string literal = m.Substring(start, length).Replace("\"\"", "\"");
            // M has no server-relative paths: "/" and "\" are separators in Text.Split and the like.
            if (PartRewriters.RewriteLiteral(literal, ctx, LinkKind.PowerQuery, serverRelative: false) is { } replaced)
                edits.Replace(start, length, replaced.Replace("\"", "\"\""));
        }
        return edits.Count == 0 ? null : edits.Apply(m);
    }

    private static byte[]? RewritePackage(byte[] package, LinkContext ctx)
    {
        var changed = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var (name, bytes) in PackageEntries(package))
        {
            var text = PartText.Decode(bytes);
            if (RewriteFormulas(text.Text, ctx) is { } updated)
                changed[name] = text.Encode(updated);
        }
        if (changed.Count == 0)
            return null;

        using var ms = new MemoryStream();
        ms.Write(package);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var (name, bytes) in changed)
            {
                using var stream = zip.GetEntry(name)!.Open();
                stream.SetLength(0);
                stream.Write(bytes);
            }
        }
        return ms.ToArray();
    }

    /// <summary>The .m entries of the package parts zip (Formulas/Section1.m).</summary>
    private static List<(string Name, byte[] Bytes)> PackageEntries(byte[] package)
    {
        var entries = new List<(string, byte[])>();
        using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".m", StringComparison.OrdinalIgnoreCase)))
        {
            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            entries.Add((entry.FullName, ms.ToArray()));
        }
        return entries;
    }

    private static byte[] Decode(string base64)
    {
        try
        {
            return Convert.FromBase64String(XmlText.Unescape(base64));
        }
        catch (FormatException ex)
        {
            throw new InvalidDataException(L.T("Power Query (DataMashup): не base64."), ex);
        }
    }

    private static string Lenient(byte[] bytes) => Encoding.UTF8.GetString(bytes);

    /// <summary>[MS-QDEFF] 2.2 top-level structure; bytes after the four blocks, if any, are kept.</summary>
    private sealed record Mashup(uint Version, byte[] PackageParts, byte[] Permissions, byte[] Metadata, byte[] Bindings, byte[] Tail)
    {
        public static Mashup Parse(byte[] data)
        {
            int pos = 0;
            uint version = ReadUInt(data, ref pos);
            if (version != 0)
                throw new InvalidDataException(L.F("Power Query (DataMashup): невідома версія {0}.", version));
            byte[] package = ReadBlock(data, ref pos);
            byte[] permissions = ReadBlock(data, ref pos);
            byte[] metadata = ReadBlock(data, ref pos);
            byte[] bindings = ReadBlock(data, ref pos);
            return new Mashup(version, package, permissions, metadata, bindings, data[pos..]);
        }

        public byte[] ToBytes()
        {
            using var ms = new MemoryStream();
            Span<byte> u = stackalloc byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(u, Version);
            ms.Write(u);
            foreach (byte[] block in new[] { PackageParts, Permissions, Metadata, Bindings })
            {
                BinaryPrimitives.WriteInt32LittleEndian(u, block.Length);
                ms.Write(u);
                ms.Write(block);
            }
            ms.Write(Tail);
            return ms.ToArray();
        }

        private static uint ReadUInt(byte[] data, ref int pos)
        {
            if (pos + 4 > data.Length)
                throw new InvalidDataException(L.T("Power Query (DataMashup): структура обрізана."));
            uint value = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(pos, 4));
            pos += 4;
            return value;
        }

        private static byte[] ReadBlock(byte[] data, ref int pos)
        {
            uint length = ReadUInt(data, ref pos);
            if (length > (uint)(data.Length - pos))
                throw new InvalidDataException(L.T("Power Query (DataMashup): структура обрізана."));
            byte[] block = data.AsSpan(pos, (int)length).ToArray();
            pos += (int)length;
            return block;
        }
    }
}
