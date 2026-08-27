using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// D18. A <c>&lt;pane state="split"&gt;</c> — Excel's draggable split bar, what View → Split gives
/// you without Freeze Panes — was dropped on load, because the reader took the two splits only for
/// a frozen state; and a split the model did hold was written back as a freeze, because the pane
/// resolver hardcoded <c>state="frozen"</c>. A frozen pane cannot be dragged and does not scroll
/// independently, so the round trip silently changed how the sheet behaves.
/// </summary>
public class SheetViewSplitPaneTests
{
    [Test]
    public async Task An_unfrozen_split_survives_a_round_trip()
    {
        using var package = SplitPackage();
        await Assert.That(Attribute(PaneTag(package), "state")).IsEqualTo("split");

        using (var wb = new XLWorkbook(package))
        {
            var view = wb.Worksheet("S").SheetView;
            await Assert.That(view.SplitRow).IsEqualTo(3);
            await Assert.That(view.SplitColumn).IsEqualTo(2);
            await Assert.That(view.FreezePanes).IsFalse();

            wb.SaveAs(package);
        }

        var pane = PaneTag(package);
        await Assert.That(Attribute(pane, "state")).IsEqualTo("split");
        await Assert.That(Attribute(pane, "xSplit")).IsEqualTo("2");
        await Assert.That(Attribute(pane, "ySplit")).IsEqualTo("3");
    }

    [Test]
    public async Task A_split_set_without_freezing_is_written_as_a_split()
    {
        // SplitRow and SplitColumn are on the public interface, so a caller can ask for a split
        // without ever calling Freeze. That is a split, not a freeze.
        using var package = Save(ws =>
        {
            ws.SheetView.SplitRow = 3;
            ws.SheetView.SplitColumn = 2;
        });

        var pane = PaneTag(package);
        await Assert.That(Attribute(pane, "state")).IsEqualTo("split");
        await Assert.That(Attribute(pane, "xSplit")).IsEqualTo("2");
        await Assert.That(Attribute(pane, "ySplit")).IsEqualTo("3");

        // An unfrozen split states its position in twentieths of a point, so the first-unfrozen-cell
        // anchor a freeze derives from split + 1 would name a cell hundreds of columns away.
        await Assert.That(Attribute(pane, "topLeftCell")).IsEqualTo("A1");

        using var wb = new XLWorkbook(package);
        await Assert.That(wb.Worksheet("S").SheetView.FreezePanes).IsFalse();
    }

    [Test]
    public async Task A_freeze_is_still_written_as_a_freeze()
    {
        using var package = Save(ws => ws.SheetView.Freeze(3, 2));

        var pane = PaneTag(package);
        await Assert.That(Attribute(pane, "state")).IsEqualTo("frozen");

        using var wb = new XLWorkbook(package);
        var view = wb.Worksheet("S").SheetView;
        await Assert.That(view.SplitRow).IsEqualTo(3);
        await Assert.That(view.SplitColumn).IsEqualTo(2);
        await Assert.That(view.FreezePanes).IsTrue();
    }

    /// <summary>
    /// <c>frozenSplit</c> is a pane frozen out of an existing manual split. The model carries a
    /// boolean, so it loads as frozen and saves as <c>frozen</c> — spec 29's normalisation, kept.
    /// </summary>
    [Test]
    public async Task A_frozen_split_loads_as_a_freeze()
    {
        using var package = Save(ws => ws.SheetView.Freeze(3, 2));
        RewriteSheet1(package, xml => xml.Replace("state=\"frozen\"", "state=\"frozenSplit\""));

        using (var wb = new XLWorkbook(package))
        {
            var view = wb.Worksheet("S").SheetView;
            await Assert.That(view.FreezePanes).IsTrue();
            await Assert.That(view.SplitRow).IsEqualTo(3);
            await Assert.That(view.SplitColumn).IsEqualTo(2);

            wb.SaveAs(package);
        }

        await Assert.That(Attribute(PaneTag(package), "state")).IsEqualTo("frozen");
    }

    /// <summary>
    /// An unfrozen split's position is not a line count, so it names no cell to scroll a pane to.
    /// Such a sheet scrolls through the view's own top-left, like a sheet with no pane at all.
    /// </summary>
    [Test]
    public async Task Focusing_a_cell_on_a_split_sheet_scrolls_the_view()
    {
        using var package = SplitPackage();
        using (var wb = new XLWorkbook(package))
        {
            var ws = wb.Worksheet("S");
            ws.FocusCell(ws.Cell("B2")!);
            wb.SaveAs(package);
        }

        await Assert.That(Attribute(Tag(package, "sheetView"), "topLeftCell")).IsEqualTo("B2");
        await Assert.That(Attribute(PaneTag(package), "topLeftCell")).IsEqualTo("A1");
    }

    /// <summary>
    /// The split grows and shrinks with an edit inside it because it counts frozen lines — which it
    /// only does for a freeze. Twentieths of a point are not lines, and no row edit moves them.
    /// </summary>
    [Test]
    public async Task A_row_or_column_edit_does_not_move_an_unfrozen_split()
    {
        using var package = SplitPackage();
        using var wb = new XLWorkbook(package);
        var view = wb.Worksheet("S").SheetView;

        view.Worksheet.Row(1).InsertRowsAbove(3);
        view.Worksheet.Column(1).InsertColumnsBefore(3);

        await Assert.That(view.SplitRow).IsEqualTo(3);
        await Assert.That(view.SplitColumn).IsEqualTo(2);

        // The case that loses the split bar outright: deleting every line "above" a 3-twip split.
        view.Worksheet.Rows(1, 100).Delete();

        await Assert.That(view.SplitRow).IsEqualTo(3);
    }

    /// <summary>
    /// Freezing one axis cannot reinterpret the other: the axis left alone is holding twentieths of
    /// a point, and reading 2 of those as two frozen columns is not a translation.
    /// </summary>
    [Test]
    public async Task Freezing_one_axis_of_an_unfrozen_split_drops_the_other()
    {
        using var package = SplitPackage();
        using var wb = new XLWorkbook(package);
        var view = wb.Worksheet("S").SheetView;

        view.FreezeRows(2);

        await Assert.That(view.FreezePanes).IsTrue();
        await Assert.That(view.SplitRow).IsEqualTo(2);
        await Assert.That(view.SplitColumn).IsEqualTo(0);
    }

    #region Helpers

    /// <summary>
    /// A package whose only sheet carries
    /// <c>&lt;pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="split"/&gt;</c>
    /// — written by freezing and then downgrading the state, so the rest of the pane is whatever
    /// XLibur itself would write.
    /// </summary>
    private static MemoryStream SplitPackage()
    {
        var package = Save(ws => ws.SheetView.Freeze(3, 2));
        RewriteSheet1(package, xml => xml.Replace("state=\"frozen\"", "state=\"split\""));
        return package;
    }

    private static MemoryStream Save(Action<IXLWorksheet> arrange)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = "x";
            arrange(ws);
            wb.SaveAs(ms);
        }

        return ms;
    }

    private static void RewriteSheet1(MemoryStream package, Func<string, string> transform)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals("xl/worksheets/sheet1.xml", StringComparison.OrdinalIgnoreCase));

        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();

        var rewritten = transform(xml);
        if (rewritten == xml)
            throw new InvalidOperationException(
                "The sheet1.xml transform changed nothing, so the test is not exercising what it says.");

        using var stream = entry.Open();
        stream.SetLength(0);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(rewritten);
    }

    private static string PaneTag(MemoryStream package) => Tag(package, "pane");

    /// <summary>
    /// Matches an element whatever prefix the serialiser gave it — see
    /// <see cref="WritePathAgreementTests"/>, which reads the same bytes for the same reason.
    /// </summary>
    private static string Tag(MemoryStream package, string name)
    {
        var match = Regex.Match(ReadSheet1(package), $"<(?:[A-Za-z_][\\w.-]*:)?{name}\\b[^>]*>");
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
