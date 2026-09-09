using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Hand-editing of a saved package's parts, for load-path tests that need a file XLibur itself
/// would never write — an attribute XLibur always pairs with another one, or a value outside the
/// enum it maps to. Producing those from the public API is impossible by construction, which is
/// exactly why the reader has never been exercised against them.
/// </summary>
internal static class SheetPackage
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string WorkbookPart = "xl/workbook.xml";

    /// <summary>
    /// Applies <paramref name="rewrite"/> to <c>xl/worksheets/sheet1.xml</c> inside
    /// <paramref name="package"/>, in place. Returns the same stream for chaining.
    /// </summary>
    internal static MemoryStream RewriteSheet1(this MemoryStream package, Func<string, string> rewrite)
        => package.RewritePart(Sheet1, rewrite);

    /// <summary>
    /// Applies <paramref name="rewrite"/> to <c>xl/workbook.xml</c> inside <paramref name="package"/>,
    /// in place. Returns the same stream for chaining.
    /// </summary>
    internal static MemoryStream RewriteWorkbook(this MemoryStream package, Func<string, string> rewrite)
        => package.RewritePart(WorkbookPart, rewrite);

    /// <summary>The text of <c>xl/worksheets/sheet1.xml</c>.</summary>
    internal static string Sheet1Xml(this MemoryStream package) => package.PartXml(Sheet1);

    /// <summary>The text of <c>xl/workbook.xml</c>.</summary>
    internal static string WorkbookXml(this MemoryStream package) => package.PartXml(WorkbookPart);

    private static MemoryStream RewritePart(this MemoryStream package, string partName,
        Func<string, string> rewrite)
    {
        package.Position = 0;

        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.Entries.First(e =>
                e.FullName.Equals(partName, StringComparison.OrdinalIgnoreCase));

            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();

            var rewritten = Encoding.UTF8.GetBytes(rewrite(xml));

            using var write = entry.Open();
            write.SetLength(0);
            write.Write(rewritten, 0, rewritten.Length);
        }

        package.Position = 0;
        return package;
    }

    private static string PartXml(this MemoryStream package, string partName)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partName, StringComparison.OrdinalIgnoreCase));

        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
