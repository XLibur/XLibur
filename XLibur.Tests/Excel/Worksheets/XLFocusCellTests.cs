#nullable enable
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Worksheets;

[TestFixture]
public class XLFocusCellTests
{
    // 4.1 - the low-level primitive is honored on a frozen sheet (A10 != the A3 split+1 default).
    [Test]
    public void PaneTopLeftCellAddress_OnFrozenSheet_IsEmitted()
    {
        using var ms = Save(ws =>
        {
            ws.SheetView.FreezeRows(2);
            ws.SheetView.PaneTopLeftCellAddress = ws.Cell("A10").Address;
        });

        InspectSheetView(ms, sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            Assert.That(pane.TopLeftCell!.Value, Is.EqualTo("A10"));
        });
    }

    // 4.2 - null default normalizes the pane to split+1 (regression guard for normalize-to-top).
    [Test]
    public void PaneTopLeftCellAddress_Unset_NormalizesToSplitPlusOne()
    {
        using var ms = Save(ws => ws.SheetView.FreezeRows(2));

        InspectSheetView(ms, sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            Assert.That(pane.TopLeftCell!.Value, Is.EqualTo("A3"));
        });
    }

    // 4.3 - without a frozen pane, the primitive has no effect (no <pane> emitted).
    [Test]
    public void PaneTopLeftCellAddress_WithoutPane_IsIgnored()
    {
        using var ms = Save(ws => ws.SheetView.PaneTopLeftCellAddress = ws.Cell("M50").Address);

        InspectSheetView(ms, sv =>
        {
            Assert.That(sv!.Elements<Pane>(), Is.Empty);
            Assert.That(sv.TopLeftCell, Is.Null);
        });
    }

    // 4.4 - FocusCell on a frozen sheet scrolls the pane and clears a residual horizontal scroll.
    [Test]
    public void FocusCell_OnFrozenSheet_ScrollsPaneAndClearsResidual()
    {
        using var ms = Save(ws =>
        {
            ws.SheetView.FreezeRows(2);
            ws.SheetView.TopLeftCellAddress = ws.Cell("G1").Address; // residual horizontal scroll
            ws.FocusCell("A3");
        });

        InspectSheetView(ms, sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            Assert.That(pane.TopLeftCell!.Value, Is.EqualTo("A3"));
            Assert.That(sv.TopLeftCell, Is.Null, "residual G1 should be cleared");

            var selection = sv.Elements<Selection>().First(s => s.Pane is not null);
            Assert.That(selection.ActiveCell!.Value, Is.EqualTo("A3"));
            Assert.That(selection.SequenceOfReferences!.InnerText, Is.EqualTo("A3"));
        });
    }

    // 4.5 - FocusCell on a non-frozen sheet sets the view top-left and emits no <pane>.
    [Test]
    public void FocusCell_OnNonFrozenSheet_SetsSheetViewTopLeft()
    {
        using var ms = Save(ws => ws.FocusCell("M50"));

        InspectSheetView(ms, sv =>
        {
            Assert.That(sv!.Elements<Pane>(), Is.Empty);
            Assert.That(sv.TopLeftCell!.Value, Is.EqualTo("M50"));

            var selection = sv.Elements<Selection>().First();
            Assert.That(selection.ActiveCell!.Value, Is.EqualTo("M50"));
            Assert.That(selection.SequenceOfReferences!.InnerText, Is.EqualTo("M50"));
        });
    }

    // 4.6 - SetActiveCell / SetActive never move the scroll position.
    [Test]
    public void SetActiveCell_DoesNotMoveScroll()
    {
        using var ms = Save(ws => ws.SetActiveCell("A3"));

        InspectSheetView(ms, sv =>
        {
            Assert.That(sv!.Elements<Pane>(), Is.Empty);
            Assert.That(sv.TopLeftCell, Is.Null);
            Assert.That(sv.Elements<Selection>().First().ActiveCell!.Value, Is.EqualTo("A3"));
        });
    }

    // 4.7 - focusing a cell inside the frozen band resets the pane to origin and names the owning pane.
    [Test]
    public void FocusCell_InFrozenRegion_ResetsPaneAndNamesOwningPane()
    {
        using var ms = Save(ws =>
        {
            ws.SheetView.FreezeRows(2);
            ws.FocusCell("A1");
        });

        InspectSheetView(ms, sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            Assert.That(pane.TopLeftCell!.Value, Is.EqualTo("A3"), "scrollable region reset to split+1");

            var selection = sv.Elements<Selection>().First(s => s.Pane is not null);
            Assert.That(selection.Pane!.Value, Is.EqualTo(PaneValues.TopLeft), "A1 lives in the frozen pane");
            Assert.That(selection.ActiveCell!.Value, Is.EqualTo("A1"));
        });
    }

    // 4.8 - freeze-shape matrix: orthogonal axis resets to origin per the single-axis cases.
    [TestCase("rows", "D10", "A10")]      // row-only freeze: column reset to A
    [TestCase("columns", "M5", "M1")]     // column-only freeze: row reset to 1
    [TestCase("both", "F8", "F8")]        // both-axis freeze: anchor both
    public void FocusCell_FreezeShapeMatrix(string freeze, string target, string expectedPane)
    {
        using var ms = Save(ws =>
        {
            switch (freeze)
            {
                case "rows": ws.SheetView.FreezeRows(2); break;
                case "columns": ws.SheetView.FreezeColumns(2); break;
                case "both": ws.SheetView.Freeze(2, 2); break;
            }

            ws.FocusCell(target);
        });

        InspectSheetView(ms, sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            Assert.That(pane.TopLeftCell!.Value, Is.EqualTo(expectedPane));
        });
    }

    // 4.9 - setting the pane address from a foreign worksheet throws.
    [Test]
    public void PaneTopLeftCellAddress_FromOtherWorksheet_Throws()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        Assert.Throws<ArgumentException>(() =>
            ws1.SheetView.PaneTopLeftCellAddress = ws2.Cell("A1").Address);
    }

    // 4.10 - an unresolvable address throws a descriptive ArgumentException, not a NullReferenceException.
    [Test]
    public void SetActiveCellAndFocusCell_InvalidAddress_Throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        Assert.Throws<ArgumentException>(() => ws.SetActiveCell("NotAName"));
        Assert.Throws<ArgumentException>(() => ws.FocusCell("NotAName"));
    }

    private static MemoryStream Save(Action<IXLWorksheet> configure)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            configure(ws);
            wb.SaveAs(ms);
        }

        return ms;
    }

    private static void InspectSheetView(MemoryStream ms, Action<SheetView?> assert)
    {
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var wsPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetView = wsPart.Worksheet!.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
        assert(sheetView);
    }
}
