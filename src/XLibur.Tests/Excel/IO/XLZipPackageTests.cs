using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO.Packaging;

namespace XLibur.Tests.Excel.IO;

public class XLZipPackageTests
{
    private const string Styles = "xl/styles.xml";

    [Test]
    public async Task A_part_only_read_is_stored_at_the_level_when_the_package_closes()
    {
        // Reading an entry makes the archive write it again, so a part the save only reads is
        // replaced at the requested level when the package closes, with the content it had.
        var source = NewPackage();
        var (content, isCompressed) = ReadEntry(source, Styles);
        await Assert.That(isCompressed).IsTrue();

        using var stream = new MemoryStream();
        stream.Write(source);
        var package = new XLZipPackage(stream, CompressionLevel.NoCompression);
        var part = package.GetPart(new Uri("/" + Styles, UriKind.Relative));
        using (var read = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var copy = new MemoryStream())
        {
            read.CopyTo(copy);
            await Assert.That(copy.ToArray()).IsEquivalentTo(content);
        }

        package.Close();

        var (savedContent, savedIsCompressed) = ReadEntry(stream.ToArray(), Styles);
        await Assert.That(savedIsCompressed).IsFalse();
        await Assert.That(savedContent).IsEquivalentTo(content);
    }

    private static byte[] NewPackage()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1").Cell("A1").Value = 1;
        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.ToArray();
    }

    private static (byte[] Content, bool IsCompressed) ReadEntry(byte[] package, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var entry = zip.GetEntry(name)!;
        using var content = new MemoryStream();
        using (var stream = entry.Open())
            stream.CopyTo(content);

        return (content.ToArray(), entry.CompressedLength < entry.Length);
    }
}
