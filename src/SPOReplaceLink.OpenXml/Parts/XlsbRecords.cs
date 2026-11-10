namespace SPOReplaceLink.OpenXml.Parts;

/// <summary>Minimal reader of BIFF12 record framing in xlsb parts ([MS-XLSB] 2.1.4).</summary>
internal static class XlsbRecords
{
    /// <summary>BrtAbsPath15: folder the workbook was last saved in.</summary>
    public const int AbsPath15 = 2071;

    /// <summary>Copy of the part with payloads of the given record type zeroed. Stops quietly at malformed data.</summary>
    public static byte[] MaskRecords(byte[] part, int recordType)
    {
        byte[]? copy = null;
        int pos = 0;
        while (pos < part.Length)
        {
            if (!TryReadVarint(part, ref pos, 2, out int type) || !TryReadVarint(part, ref pos, 4, out int size)
                || size < 0 || pos + size > part.Length)
                break;
            if (type == recordType)
            {
                copy ??= (byte[])part.Clone();
                Array.Clear(copy, pos, size);
            }
            pos += size;
        }
        return copy ?? part;
    }

    private static bool TryReadVarint(byte[] data, ref int pos, int maxBytes, out int value)
    {
        value = 0;
        for (int i = 0; i < maxBytes; i++)
        {
            if (pos >= data.Length)
                return false;
            byte b = data[pos++];
            value |= (b & 0x7F) << (7 * i);
            if ((b & 0x80) == 0)
                return true;
        }
        return true;
    }
}
