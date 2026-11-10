using System.Text;

namespace SPOReplaceLink.OpenXml.Text;

/// <summary>Text of an XML part together with the encoding it was stored in, so it is written back the same way.</summary>
internal sealed record PartText(string Text, Encoding Encoding, bool HasBom)
{
    private static readonly UTF8Encoding Utf8 = new(false, true);

    public static PartText Decode(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return new PartText(Utf8.GetString(bytes, 3, bytes.Length - 3), Utf8, true);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return new PartText(Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2), Encoding.Unicode, true);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return new PartText(Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2), Encoding.BigEndianUnicode, true);
        return new PartText(Utf8.GetString(bytes), Utf8, false);
    }

    public byte[] Encode(string text)
    {
        byte[] body = Encoding.GetBytes(text);
        if (!HasBom)
            return body;
        byte[] bom = Encoding.GetPreamble();
        if (bom.Length == 0 && Encoding is UTF8Encoding)
            bom = [0xEF, 0xBB, 0xBF];
        var result = new byte[bom.Length + body.Length];
        bom.CopyTo(result, 0);
        body.CopyTo(result, bom.Length);
        return result;
    }
}
