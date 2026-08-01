#nullable enable
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Worksheets;

public class XLFocusCellTests
{
    // 4.1 - the low-level primitive is honored on a frozen sheet (A10 != the A3 split+1 default).
    [Test]
    public async Task PaneTopLeftCellAddress_OnFrozenSheet_IsEmitted()
    {
        using var ms = Save(ws =>
        {
            ws.SheetView.FreezeRows(2);
            ws.SheetView.PaneTopLeftCellAddress = ws.Cell("A10").Address;
        });

        await InspectSheetView(ms, async sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            await Assert.That(pane.TopLeftCell!.Value).IsEqualTo("A10");
        });
    }

    // 4.2 - null default normalizes the pane to split+1 (regression guard for normalize-to-top).
    [Test]
    public async Task PaneTopLeftCellAddress_Unset_NormalizesToSplitPlusOne()
    {
        using var ms = Save(ws => ws.SheetView.FreezeRows(2));

        await InspectSheetView(ms, async sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            await Assert.That(pane.TopLeftCell!.Value).IsEqualTo("A3");
        });
    }

    // 4.3 - without a frozen pane, the primitive has no effect (no <pane> emitted).
    [Test]
    public async Task PaneTopLeftCellAddress_WithoutPane_IsIgnored()
    {
        using var ms = Save(ws => ws.SheetView.PaneTopLeftCellAddress = ws.Cell("M50").Address);

        await InspectSheetView(ms, async sv =>
        {
            await Assert.That(sv!.Elements<Pane>()).IsEmpty();
            await Assert.That(sv.TopLeftCell).IsNull();
        });
    }

    // 4.4 - FocusCell on a frozen sheet scrolls the pane and clears a residual horizontal scroll.
    [Test]
    public async Task FocusCell_OnFrozenSheet_ScrollsPaneAndClearsResidual()
    {
        using var ms = Save(ws =>
        {
            ws.SheetView.FreezeRows(2);
            ws.SheetView.TopLeftCellAddress = ws.Cell("G1").Address; // residual horizontal scroll
            ws.FocusCell("A3");
        });

        await InspectSheetView(ms, async sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            await Assert.That(pane.TopLeftCell!.Value).IsEqualTo("A3");
            await Assert.That(sv.TopLeftCell).IsNull().Because("residual G1 should be cleared");

            var selection = sv.Elements<Selection>().First(s => s.Pane is not null);
            await Assert.That(selection.ActiveCell!.Value).IsEqualTo("A3");
            await Assert.That(selection.SequenceOfReferences!.InnerText).IsEqualTo("A3");
        });
    }

    // 4.5 - FocusCell on a non-frozen sheet sets the view top-left and emits no <pane>.
    [Test]
    public async Task FocusCell_OnNonFrozenSheet_SetsSheetViewTopLeft()
    {
        using var ms = Save(ws => ws.FocusCell("M50"));

        await InspectSheetView(ms, async sv =>
        {
            await Assert.That(sv!.Elements<Pane>()).IsEmpty();
            await Assert.That(sv.TopLeftCell!.Value).IsEqualTo("M50");

            var selection = sv.Elements<Selection>().First();
            await Assert.That(selection.ActiveCell!.Value).IsEqualTo("M50");
            await Assert.That(selection.SequenceOfReferences!.InnerText).IsEqualTo("M50");
        });
    }

    // 4.6 - SetActiveCell / SetActive never move the scroll position.
    [Test]
    public async Task SetActiveCell_DoesNotMoveScroll()
    {
        using var ms = Save(ws => ws.SetActiveCell("A3"));

        await InspectSheetView(ms, async sv =>
        {
            await Assert.That(sv!.Elements<Pane>()).IsEmpty();
            await Assert.That(sv.TopLeftCell).IsNull();
            await Assert.That(sv.Elements<Selection>().First().ActiveCell!.Value).IsEqualTo("A3");
        });
    }

    // 4.7 - focusing a cell inside the frozen band resets the pane to origin and names the owning pane.
    [Test]
    public async Task FocusCell_InFrozenRegion_ResetsPaneAndNamesOwningPane()
    {
        using var ms = Save(ws =>
        {
            ws.SheetView.FreezeRows(2);
            ws.FocusCell("A1");
        });

        await InspectSheetView(ms, async sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            await Assert.That(pane.TopLeftCell!.Value).IsEqualTo("A3").Because("scrollable region reset to split+1");

            var selection = sv.Elements<Selection>().First(s => s.Pane is not null);
            await Assert.That(selection.Pane!.Value).IsEqualTo(PaneValues.TopLeft).Because("A1 lives in the frozen pane");
            await Assert.That(selection.ActiveCell!.Value).IsEqualTo("A1");
        });
    }

    // 4.8 - freeze-shape matrix: orthogonal axis resets to origin per the single-axis cases.
    [Test]
    [Arguments("rows", "D10", "A10")]      // row-only freeze: column reset to A
    [Arguments("columns", "M5", "M1")]     // column-only freeze: row reset to 1
    [Arguments("both", "F8", "F8")]        // both-axis freeze: anchor both
    public async Task FocusCell_FreezeShapeMatrix(string freeze, string target, string expectedPane)
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

        await InspectSheetView(ms, async sv =>
        {
            var pane = sv!.Elements<Pane>().Single();
            await Assert.That(pane.TopLeftCell!.Value).IsEqualTo(expectedPane);
        });
    }

    // 4.9 - setting the pane address from a foreign worksheet throws.
    [Test]
    public async Task PaneTopLeftCellAddress_FromOtherWorksheet_Throws()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        await Assert.That(() =>
            ws1.SheetView.PaneTopLeftCellAddress = ws2.Cell("A1").Address).Throws<ArgumentException>();
    }

    // 4.10 - an unresolvable address throws a descriptive ArgumentException, not a NullReferenceException.
    [Test]
    public async Task SetActiveCellAndFocusCell_InvalidAddress_Throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(() => ws.SetActiveCell("NotAName")).Throws<ArgumentException>();
        await Assert.That(() => ws.FocusCell("NotAName")).Throws<ArgumentException>();
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

    // assert is Func<..., Task>: TUnit assertions are awaitable, and an async lambda
    // passed as an Action would be async void, silently swallowing failures.
    private static async Task InspectSheetView(MemoryStream ms, Func<SheetView?, Task> assert)
    {
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var wsPart = doc.WorkbookPart!.WorksheetParts.First();
        var sheetView = wsPart.Worksheet!.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
        await assert(sheetView);
    }
}
