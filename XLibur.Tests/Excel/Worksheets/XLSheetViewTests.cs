
using XLibur.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
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
    /// The one property excluded is <c>TabSelected</c>, which is selection rather than appearance;
    /// <see cref="Copy_does_not_duplicate_the_selected_tab"/> covers it from the other side.
    /// </summary>
    [Test]
    public async Task AllViewProperties_survive_copy()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("S1");
        foreach (var property in XLViewProperties.All)
            property.SetNonDefault(ws1);

        var ws2 = ws1.CopyTo("S2");

        foreach (var property in XLViewProperties.All.Where(p => p.SurvivesCopy))
        {
            await Assert.That(property.Get(ws2)).IsEqualTo(property.Get(ws1))
                .Because($"{property.Name} did not survive copy");
        }
    }

    /// <summary>
    /// <see cref="XLViewProperties.All"/>'s <c>TabColor</c> entry uses a plain RGB colour and copies
    /// within one workbook, so it cannot see a bug specific to a themed colour crossing into a
    /// workbook whose theme assigns different RGB values to the same theme slot. A themed
    /// <see cref="XLColor"/> carries a theme slot and a tint, not a baked RGB value — Excel resolves
    /// it against whichever workbook's theme contains it — so the correct behaviour on copy is to
    /// carry the (slot, tint) pair across unchanged, not to resolve and rebake it against either
    /// workbook's theme.
    /// </summary>
    [Test]
    public async Task Copy_preserves_themed_TabColor_across_workbooks_with_different_themes()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();

        // Give the destination workbook's theme a visibly different Accent1 than the source's, so
        // a bug that resolved the colour against the wrong workbook's theme would be observable.
        wb2.Theme.Accent1 = XLColor.FromArgb(10, 20, 30);

        var ws1 = wb1.AddWorksheet("S1");
        ws1.TabColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.25);

        var ws2 = ws1.CopyTo(wb2, "S2");

        await Assert.That(ws2.TabColor.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(ws2.TabColor.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(ws2.TabColor.ThemeTint).IsEqualTo(0.25);
        await Assert.That(ws2.TabColor).IsEqualTo(ws1.TabColor);
    }

    /// <summary>
    /// Selection is not appearance. <c>tabSelected</c> says "this is the tab the user is looking at",
    /// and a sheet loaded from a real file carries it on whichever sheet was active when Excel saved.
    /// Copying that sheet must not hand the copy the same claim: two sheets both carrying
    /// <c>tabSelected="1"</c> is how Excel encodes a <em>group</em>, so the reopened file shows
    /// "[Group]" in the title bar and the user's next edit lands on both sheets at once. The copy
    /// therefore starts unselected, and the source keeps the selection it had.
    /// </summary>
    [Test]
    public async Task Copy_does_not_duplicate_the_selected_tab()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws1 = wb.AddWorksheet("S1");
            ws1.TabSelected = true;

            var ws2 = ws1.CopyTo("S2");

            await Assert.That(ws2.TabSelected).IsFalse()
                .Because("the copy must not inherit the source's selection");
            await Assert.That(ws1.TabSelected).IsTrue()
                .Because("copying must not deselect the source either");

            wb.SaveAs(ms);
        }

        // The in-memory assertions above are the mechanism; this is the outcome that matters — a
        // saved package must never claim two selected tabs.
        await Assert.That(SelectedTabCount(ms)).IsEqualTo(1);
    }

    /// <summary>
    /// <see cref="IXLSheetView.ZoomScale"/> is the zoom of the view the sheet is currently in, and
    /// its setter mirrors the value onto the named scale for that view. Which named scale is not a
    /// free choice: <see cref="IXLSheetView.ZoomScalePageLayoutView"/> is Page Layout and
    /// <see cref="IXLSheetView.ZoomScaleSheetLayoutView"/> is Page Break Preview — per its own
    /// summary on <c>IXLSheetView</c>, and per ECMA-376 18.3.1.87, where <c>zoomScaleSheetLayoutView</c>
    /// is documented as "Zoom Scale Page Break Preview". The two arms were swapped, so setting the
    /// zoom of a page-layout sheet wrote the page-break-preview scale and left page layout at 100.
    /// </summary>
    [Test]
    public async Task ZoomScale_mirrors_onto_the_scale_named_by_the_current_view()
    {
        using var wb = new XLWorkbook();

        var pageLayout = wb.AddWorksheet("PL");
        pageLayout.SheetView.SetView(XLSheetViewOptions.PageLayout);
        pageLayout.SheetView.ZoomScale = 140;

        await Assert.That(pageLayout.SheetView.ZoomScalePageLayoutView).IsEqualTo(140)
            .Because("Page Layout's zoom belongs in zoomScalePageLayoutView");
        await Assert.That(pageLayout.SheetView.ZoomScaleSheetLayoutView).IsEqualTo(100)
            .Because("nothing set a Page Break Preview zoom on this sheet");

        var pageBreak = wb.AddWorksheet("PB");
        pageBreak.SheetView.SetView(XLSheetViewOptions.PageBreakPreview);
        pageBreak.SheetView.ZoomScale = 60;

        await Assert.That(pageBreak.SheetView.ZoomScaleSheetLayoutView).IsEqualTo(60)
            .Because("Page Break Preview's zoom belongs in zoomScaleSheetLayoutView");
        await Assert.That(pageBreak.SheetView.ZoomScalePageLayoutView).IsEqualTo(100)
            .Because("nothing set a Page Layout zoom on this sheet");

        var normal = wb.AddWorksheet("N");
        normal.SheetView.ZoomScale = 75;

        await Assert.That(normal.SheetView.ZoomScaleNormal).IsEqualTo(75);
    }

    /// <summary>How many worksheet parts in the package carry <c>tabSelected="1"</c>.</summary>
    private static int SelectedTabCount(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);

        return archive.Entries
            .Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase))
            .Count(e =>
            {
                using var reader = new StreamReader(e.Open());
                return Regex.IsMatch(reader.ReadToEnd(), "\\btabSelected=\"(1|true)\"");
            });
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
