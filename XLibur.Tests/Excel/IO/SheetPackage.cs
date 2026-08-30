using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Hand-editing of a saved package's first worksheet part, for load-path tests that need a file
/// XLibur itself would never write — an attribute XLibur always pairs with another one, or a value
/// outside the enum it maps to. Producing those from the public API is impossible by construction,
/// which is exactly why the reader has never been exercised against them.
/// </summary>
internal static class SheetPackage
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    /// <summary>
    /// Applies <paramref name="rewrite"/> to <c>xl/worksheets/sheet1.xml</c> inside
    /// <paramref name="package"/>, in place. Returns the same stream for chaining.
    /// </summary>
    internal static MemoryStream RewriteSheet1(this MemoryStream package, Func<string, string> rewrite)
    {
        package.Position = 0;

        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.Entries.First(e =>
                e.FullName.Equals(Sheet1, StringComparison.OrdinalIgnoreCase));

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

    /// <summary>The text of <c>xl/worksheets/sheet1.xml</c>.</summary>
    internal static string Sheet1Xml(this MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(Sheet1, StringComparison.OrdinalIgnoreCase));

        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
