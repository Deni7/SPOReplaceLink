using System.Text;

namespace SPOReplaceLink.OpenXml.Parts;

/// <summary>Collects replacements of source ranges and applies them in one pass; everything else is copied as is.</summary>
internal sealed class TextEdits
{
    private readonly List<(int Start, int Length, string Text)> _edits = [];

    public int Count => _edits.Count;

    public void Replace(int start, int length, string text) => _edits.Add((start, length, text));

    public string Apply(string source)
    {
        if (_edits.Count == 0)
            return source;
        var sb = new StringBuilder(source.Length + 256);
        int pos = 0;
        foreach (var (start, length, text) in _edits.OrderBy(e => e.Start))
        {
            if (start < pos)
                throw new InvalidOperationException("Overlapping edits.");
            sb.Append(source, pos, start - pos).Append(text);
            pos = start + length;
        }
        sb.Append(source, pos, source.Length - pos);
        return sb.ToString();
    }
}
