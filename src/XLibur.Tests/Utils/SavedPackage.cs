using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace XLibur.Tests.Utils;

/// <summary>
/// Reads, and hand-edits, the raw parts of a saved package. Tests use it to assert on the XML a
/// save wrote, and to produce load-path inputs XLibur itself would never write — an attribute it
/// always pairs with another one, or a value outside the enum it maps to.
/// </summary>
/// <remarks>
/// <para>
/// Every reader rewinds the stream first and leaves it open, so one saved package can be read
/// any number of times and then loaded.
/// </para>
/// <para>
/// A part is named by its path inside the zip, e.g. <c>xl/workbook.xml</c>, matched ignoring case.
/// The missing-part contract is explicit: <see cref="ReadPart"/>, <see cref="PartBytes"/> and
/// <see cref="ReadPartUnder"/> throw, naming the parts the package does have, so a wrong path
/// fails loudly instead of letting a <c>DoesNotContain</c> assertion pass on an empty string. A
/// test for which a part may legitimately be absent says so with <see cref="TryReadPart"/> or
/// <see cref="PartExists"/>.
/// </para>
/// </remarks>
internal static class SavedPackage
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string WorkbookPart = "xl/workbook.xml";

    /// <summary>The text of <c>xl/worksheets/sheet1.xml</c>.</summary>
    internal static string Sheet1Xml(this Stream package) => package.ReadPart(Sheet1);

    /// <summary>The text of <c>xl/workbook.xml</c>.</summary>
    internal static string WorkbookXml(this Stream package) => package.ReadPart(WorkbookPart);

    /// <summary>The names of every part in <paramref name="package"/>, e.g. <c>xl/workbook.xml</c>.</summary>
    internal static string[] PartNames(this Stream package)
    {
        using var archive = OpenRead(package);
        return archive.Entries.Select(e => e.FullName).ToArray();
    }

    /// <summary>The names of the parts whose path starts with <paramref name="prefix"/>, e.g. <c>xl/slicers/</c>.</summary>
    internal static string[] PartsUnder(this Stream package, string prefix)
    {
        using var archive = OpenRead(package);
        return archive.Entries
            .Where(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .ToArray();
    }

    /// <summary>Whether the package has a part named <paramref name="partName"/>.</summary>
    internal static bool PartExists(this Stream package, string partName)
    {
        using var archive = OpenRead(package);
        return Find(archive, partName) is not null;
    }

    /// <summary>Whether the package has a part whose path starts with <paramref name="prefix"/>.</summary>
    internal static bool PartExistsUnder(this Stream package, string prefix)
    {
        using var archive = OpenRead(package);
        return FindUnder(archive, prefix) is not null;
    }

    /// <summary>The raw bytes of the part named <paramref name="partName"/>. Throws when it is missing.</summary>
    internal static byte[] PartBytes(this Stream package, string partName)
    {
        using var archive = OpenRead(package);
        var entry = Find(archive, partName) ?? throw Missing(archive, $"no part '{partName}'");

        using var entryStream = entry.Open();
        using var buffer = new MemoryStream();
        entryStream.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>The text of the part named <paramref name="partName"/>. Throws when it is missing.</summary>
    internal static string ReadPart(this Stream package, string partName)
    {
        using var archive = OpenRead(package);
        var entry = Find(archive, partName) ?? throw Missing(archive, $"no part '{partName}'");
        return ReadText(entry);
    }

    /// <summary>
    /// The text of the part named <paramref name="partName"/>, or <c>false</c> when the package has
    /// no such part. For a test where the part's absence is itself an acceptable outcome.
    /// </summary>
    internal static bool TryReadPart(this Stream package, string partName, [NotNullWhen(true)] out string? xml)
    {
        using var archive = OpenRead(package);
        var entry = Find(archive, partName);
        xml = entry is null ? null : ReadText(entry);
        return xml is not null;
    }

    /// <summary>
    /// The text of the first part whose path starts with <paramref name="prefix"/>. Throws when there
    /// is none.
    /// </summary>
    /// <remarks>
    /// For a part whose exact name is not part of the contract. The OpenXML SDK names a part it
    /// creates itself differently from one Excel wrote — <c>xl/threadedcomments/threadedcomment.xml</c>
    /// rather than <c>xl/threadedComments/threadedComment1.xml</c> — and Excel resolves parts
    /// through relationships, not names.
    /// </remarks>
    internal static string ReadPartUnder(this Stream package, string prefix)
    {
        using var archive = OpenRead(package);
        var entry = FindUnder(archive, prefix) ?? throw Missing(archive, $"no part under '{prefix}'");
        return ReadText(entry);
    }

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

    /// <summary>
    /// Applies <paramref name="rewrite"/> to the part named <paramref name="partName"/> inside
    /// <paramref name="package"/>, in place. Returns the same stream, rewound, for chaining.
    /// </summary>
    internal static MemoryStream RewritePart(this MemoryStream package, string partName,
        Func<string, string> rewrite)
    {
        package.Position = 0;

        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = Find(archive, partName) ?? throw Missing(archive, $"no part '{partName}'");

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

    private static ZipArchive OpenRead(Stream package)
    {
        package.Position = 0;
        return new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
    }

    private static ZipArchiveEntry? Find(ZipArchive archive, string partName) =>
        archive.Entries.FirstOrDefault(e => e.FullName.Equals(partName, StringComparison.OrdinalIgnoreCase));

    private static ZipArchiveEntry? FindUnder(ZipArchive archive, string prefix) =>
        archive.Entries.FirstOrDefault(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

    // StreamReader drops a byte order mark, so a part is compared for what it says.
    private static string ReadText(ZipArchiveEntry entry)
    {
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static InvalidOperationException Missing(ZipArchive archive, string what) =>
        new($"The package has {what}. It has: {string.Join(", ", archive.Entries.Select(e => e.FullName))}");
}
