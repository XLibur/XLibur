using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Streaming;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// XLibur has two write paths — the ordinary DOM save and <see cref="XLStreamingWorkbook"/> — and
/// both emit <c>&lt;pane&gt;</c> and <c>&lt;col&gt;</c>. A load-and-compare test cannot see a
/// disagreement between them, because the reader normalises both spellings back to one model: that
/// is exactly how <c>state="frozenSplit"</c> vs <c>state="frozen"</c> shipped unnoticed. These tests
/// read the bytes.
/// </summary>
public class WritePathAgreementTests
{
    #region Pane

    [Test]
    [Arguments(1, 0)]
    [Arguments(0, 2)]
    [Arguments(2, 1)]
    public async Task Both_write_paths_agree_on_the_pane(int freezeRows, int freezeColumns)
    {
        using var dom = SaveViaDom(freezeRows, freezeColumns);
        using var streamed = SaveViaStreaming(freezeRows, freezeColumns);

        var domPane = PaneTag(dom);
        var streamedPane = PaneTag(streamed);

        await Assert.That(domPane).IsNotEmpty();
        await Assert.That(streamedPane).IsNotEmpty();

        foreach (var name in new[] { "state", "activePane", "topLeftCell", "xSplit", "ySplit" })
            await Assert.That(Attribute(streamedPane, name)).IsEqualTo(Attribute(domPane, name));
    }

    #endregion Pane

    #region Column

    [Test]
    public async Task Both_write_paths_agree_on_a_column()
    {
        using var dom = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            var column = ws.Column(2);
            column.Width = 33;
            column.OutlineLevel = 1;
            ws.Cell("A1").Value = "x";
            wb.SaveAs(dom);
        }

        using var streamed = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(streamed))
        {
            var sheet = wb.AddWorksheet("S");
            var column = sheet.Column(2);
            column.Width = 33;
            column.OutlineLevel = 1;
            sheet.AppendRow("x");
            wb.Finish();
        }

        var domCol = ColTag(dom, min: 2);
        var streamedCol = ColTag(streamed, min: 2);

        await Assert.That(domCol).IsNotEmpty();
        await Assert.That(streamedCol).IsNotEmpty();

        foreach (var name in new[] { "width", "customWidth", "hidden", "outlineLevel", "collapsed" })
            await Assert.That(Attribute(streamedCol, name)).IsEqualTo(Attribute(domCol, name));
    }

    #endregion Column

    #region Helpers

    private static MemoryStream SaveViaDom(int freezeRows, int freezeColumns)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.SheetView.Freeze(freezeRows, freezeColumns);
            ws.Cell("A1").Value = "x";
            wb.SaveAs(ms);
        }

        return ms;
    }

    private static MemoryStream SaveViaStreaming(int freezeRows, int freezeColumns)
    {
        var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("S");
            sheet.FreezePanes(freezeRows, freezeColumns);
            sheet.AppendRow("x");
            wb.Finish();
        }

        return ms;
    }

    /// <summary>
    /// Matches an element with or without a namespace prefix. The DOM path serialises through the
    /// OpenXML SDK, which prefixes every element (<c>&lt;x:pane&gt;</c>); the streaming path writes
    /// a default namespace and no prefix (<c>&lt;pane&gt;</c>). That is a serialisation difference,
    /// not a disagreement about content, so the harness sees past it and the attribute comparison
    /// below does the actual work.
    /// </summary>
    private const string AnyPrefix = "<(?:[A-Za-z_][\\w.-]*:)?";

    private static string PaneTag(MemoryStream package)
        => Match(ReadSheet1(package), $"{AnyPrefix}pane\\b[^>]*>");

    private static string ColTag(MemoryStream package, uint min)
        => Match(ReadSheet1(package), $"{AnyPrefix}col\\b[^>]*\\bmin=\"{min}\"[^>]*>");

    private static string Match(string xml, string pattern)
    {
        var match = Regex.Match(xml, pattern);
        return match.Success ? match.Value : string.Empty;
    }

    /// <summary>The attribute's value, or <c>null</c> when the attribute is absent.</summary>
    private static string? Attribute(string tag, string name)
    {
        var match = Regex.Match(tag, $"\\b{name}=\"([^\"]*)\"");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string ReadSheet1(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals("xl/worksheets/sheet1.xml", StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream);
        return reader.ReadToEnd();
    }

    #endregion Helpers
}
