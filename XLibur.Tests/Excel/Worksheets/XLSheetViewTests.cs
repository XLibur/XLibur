
using XLibur.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Worksheets;

public class XLSheetViewTests
{
    [Test]
    public async Task CopyWorksheetSheetViews()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();

        var ws1 = wb1.AddWorksheet("WS1");
        ws1.SheetView.TopLeftCellAddress = ws1.Cell("AZ2000").Address;

        var ws2 = ws1.CopyTo(wb2, "WS2");

        await Assert.That(ws2.SheetView.Worksheet).IsEqualTo(ws2);
        await Assert.That(ws2.SheetView.TopLeftCellAddress.ToString()).IsEqualTo("AZ2000");
    }

    [Test]
    public async Task InvalidTopLeftCell()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        await Assert.That(() => ws1.SheetView.TopLeftCellAddress = ws2.Cell("A1").Address).Throws<ArgumentException>();
    }

    #region Spec 38 regressions

    /// <summary>
    /// Defect: <see cref="XLWorksheet.CopyTo(string)"/> never touches
    /// <see cref="IXLWorksheet.ShowGridLines"/>, so a copy is seeded from the destination
    /// workbook's default (always on for a fresh workbook) instead of the source sheet's value.
    /// </summary>
    [Test]
    public async Task Copy_loses_gridlines()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("S1");
        ws1.ShowGridLines = false;

        var ws2 = ws1.CopyTo("S2");

        await Assert.That(ws2.ShowGridLines).IsFalse();
    }

    /// <summary>
    /// Defect: <see cref="XLSheetView"/>'s copy constructor only copies the pane fields; the zoom
    /// scales are left at whatever the parameterless constructor it chains to set them to (100).
    /// </summary>
    [Test]
    public async Task Copy_loses_zoom()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("S1");
        ws1.SheetView.ZoomScale = 150;

        var ws2 = ws1.CopyTo("S2");

        await Assert.That(ws2.SheetView.ZoomScale).IsEqualTo(150);
    }

    /// <summary>
    /// Defect: <c>sheetView/@view</c> is written on save but never read on load, so a workbook
    /// saved in Page Layout or Page Break Preview view reopens as Normal.
    /// </summary>
    [Test]
    public async Task ViewMode_lost_on_round_trip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S1");
            ws.SheetView.SetView(XLSheetViewOptions.PageLayout);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms);
        var reloaded = wb2.Worksheets.First();

        await Assert.That(reloaded.SheetView.View).IsEqualTo(XLSheetViewOptions.PageLayout);
    }

    /// <summary>
    /// Defect: because the copy never touches the boolean display flags, copying a sheet into a
    /// different workbook picks up that workbook's defaults rather than the source sheet's values.
    /// </summary>
    [Test]
    public async Task CrossWorkbookCopy_keeps_source_appearance_not_targets_defaults()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        wb2.ShowGridLines = false;

        var ws1 = wb1.AddWorksheet("S1");

        var ws2 = ws1.CopyTo(wb2, "S2");

        await Assert.That(ws2.ShowGridLines).IsTrue();
    }

    /// <summary>
    /// Sets every property in <see cref="XLViewProperties.All"/> to a non-default value, saves and
    /// reloads, and checks each one came back. Because the module owns the list, a property added
    /// later is covered here without a new test being written for it.
    /// </summary>
    [Test]
    public async Task AllViewProperties_survive_save_and_reload()
    {
        List<object> expected;
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S1");
            foreach (var property in XLViewProperties.All)
                property.SetNonDefault(ws);

            expected = XLViewProperties.All.Select(property => property.Get(ws)).ToList();
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloadedWorkbook = new XLWorkbook(ms);
        var reloaded = reloadedWorkbook.Worksheets.First();

        for (var i = 0; i < XLViewProperties.All.Count; i++)
        {
            var property = XLViewProperties.All[i];
            await Assert.That(property.Get(reloaded)).IsEqualTo(expected[i])
                .Because($"{property.Name} did not survive save/reload");
        }
    }

    /// <summary>
    /// The same property list, this time run against <see cref="XLWorksheet.CopyTo(string)"/>
    /// instead of a save/reload round trip — the mechanism the spec says must not diverge again.
    /// </summary>
    [Test]
    public async Task AllViewProperties_survive_copy()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("S1");
        foreach (var property in XLViewProperties.All)
            property.SetNonDefault(ws1);

        var ws2 = ws1.CopyTo("S2");

        foreach (var property in XLViewProperties.All)
        {
            await Assert.That(property.Get(ws2)).IsEqualTo(property.Get(ws1))
                .Because($"{property.Name} did not survive copy");
        }
    }

    #endregion Spec 38 regressions

    [Test]
    public async Task SheetViews()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").Address;
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.SheetView.TopLeftCellAddress.ToString()).IsEqualTo("AZ2000");

            ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000")
                .CellBelow()
                .CellRight()
                .Address;

            wb.Save();
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.SheetView.TopLeftCellAddress.ToString()).IsEqualTo("BA2001");
        }
    }
}
