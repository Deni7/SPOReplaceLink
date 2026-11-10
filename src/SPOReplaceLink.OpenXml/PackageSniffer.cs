using System.Buffers.Binary;
using System.Text;

namespace SPOReplaceLink.OpenXml;

internal enum PackageKind
{
    Zip,
    /// <summary>OLE compound file: xls, or an OOXML file encrypted with a password to open.</summary>
    Compound,
    Unknown,
}

internal static class PackageSniffer
{
    private static readonly byte[] ZipSignature = [0x50, 0x4B, 0x03, 0x04];
    internal static readonly byte[] CompoundSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private static readonly byte[] EncryptionInfo = Encoding.Unicode.GetBytes("EncryptionInfo");
    private static readonly byte[] EncryptedPackage = Encoding.Unicode.GetBytes("EncryptedPackage");

    public static PackageKind Detect(string path)
    {
        Span<byte> header = stackalloc byte[8];
        using var fs = File.OpenRead(path);
        int read = fs.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
        if (read >= 4 && header[..4].SequenceEqual(ZipSignature))
            return PackageKind.Zip;
        if (read == 8 && header.SequenceEqual(CompoundSignature))
            return PackageKind.Compound;
        return PackageKind.Unknown;
    }

    /// <summary>OOX-15: an encrypted OOXML file is a compound file with EncryptionInfo and EncryptedPackage streams.</summary>
    public static bool IsEncryptedOoxml(string path)
    {
        // Directory entries are spread over the file, so look through all of it.
        byte[] bytes = File.ReadAllBytes(path);
        return bytes.AsSpan().IndexOf(EncryptionInfo) >= 0 && bytes.AsSpan().IndexOf(EncryptedPackage) >= 0;
    }

    private const int BiffBof = 0x0809;
    private const int BiffEof = 0x000A;
    private const int BiffFilePass = 0x002F;
    private const int BiffFileSharing = 0x005B;
    private const int BiffBoundSheet = 0x0085;

    /// <summary>
    /// A password an xls asks for before Excel opens it, read from the workbook globals: FILEPASS (encrypted, a password
    /// to open) or FILESHARING with a password hash (a password to modify). Excel must never be asked to open such a file:
    /// some versions show the password prompt although a password is passed, and wait for an answer.
    /// </summary>
    public static XlsPassword DetectXlsPassword(string path)
    {
        using var file = CompoundFile.Open(path);
        if (file is null)
            return XlsPassword.None;
        // An encrypted xlsx/xlsb saved with the extension xls.
        if (file.Has("EncryptedPackage"))
            return XlsPassword.Open;
        // BIFF8 keeps the workbook in "Workbook", BIFF5 (Excel 5/95) in "Book"; both protections come before the sheets.
        byte[]? globals = file.Read("Workbook", 256 * 1024) ?? file.Read("Book", 256 * 1024);
        if (globals is null || globals.Length < 4 || BinaryPrimitives.ReadUInt16LittleEndian(globals) != BiffBof)
            return XlsPassword.None;
        for (int pos = 0; pos + 4 <= globals.Length;)
        {
            int type = BinaryPrimitives.ReadUInt16LittleEndian(globals.AsSpan(pos));
            int length = BinaryPrimitives.ReadUInt16LittleEndian(globals.AsSpan(pos + 2));
            var data = globals.AsSpan(pos + 4, Math.Min(length, globals.Length - pos - 4));
            switch (type)
            {
                case BiffFilePass:
                    return XlsPassword.Open;
                case BiffFileSharing when data.Length >= 4 && BinaryPrimitives.ReadUInt16LittleEndian(data[2..]) != 0:
                    return XlsPassword.Modify;
                case BiffEof or BiffBoundSheet:
                    return XlsPassword.None;
            }
            pos += 4 + length;
        }
        return XlsPassword.None;
    }
}

public enum XlsPassword
{
    None,
    /// <summary>Encrypted: a password to open.</summary>
    Open,
    /// <summary>A password to modify (write reservation).</summary>
    Modify,
}
