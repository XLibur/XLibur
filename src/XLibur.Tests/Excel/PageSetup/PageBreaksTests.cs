using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.PageSetup;

public class PageBreaksTests
{
    [Test]
    public async Task RowBreaksShouldBeSorted()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.PageSetup.AddHorizontalPageBreak(12);
        sheet.PageSetup.AddHorizontalPageBreak(5);
        await Assert.That(sheet.PageSetup.RowBreaks).IsEquivalentTo([5, 10, 12], CollectionOrdering.Matching);
    }

    [Test]
    public async Task ColumnBreaksShouldBeSorted()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.PageSetup.AddVerticalPageBreak(12);
        sheet.PageSetup.AddVerticalPageBreak(5);
        await Assert.That(sheet.PageSetup.ColumnBreaks).IsEquivalentTo([5, 10, 12], CollectionOrdering.Matching);
    }

    [Test]
    public async Task RowBreaksShiftWhenInsertedRowAbove()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.Row(5).InsertRowsAbove(1);
        await Assert.That(sheet.PageSetup.RowBreaks[0]).IsEqualTo(11);
    }

    [Test]
    public async Task RowBreaksNotShiftWhenInsertedRowBelow()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.Row(15).InsertRowsAbove(1);
        await Assert.That(sheet.PageSetup.RowBreaks[0]).IsEqualTo(10);
    }

    [Test]
    public async Task ColumnBreaksShiftWhenInsertedColumnBefore()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.Column(5).InsertColumnsBefore(1);
        await Assert.That(sheet.PageSetup.ColumnBreaks[0]).IsEqualTo(11);
    }

    [Test]
    public async Task ColumnBreaksNotShiftWhenInsertedColumnAfter()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.Column(15).InsertColumnsBefore(1);
        await Assert.That(sheet.PageSetup.ColumnBreaks[0]).IsEqualTo(10);
    }

    [Test]
    public async Task PageBreaksWritePerpendicularAxisAsMax()
    {
        // brk@max is the extent perpendicular to the break: a row (horizontal) break
        // spans the full column width, a column (vertical) break spans the full row
        // height. Regression for ClosedXML issue #2842 — the row break wrote
        // max=1048576 (a row count), which makes Excel render a bogus vertical
        // scrollbar; the column break had the mirror-image defect (max=16384).
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet = wb.AddWorksheet("Sheet1");
            sheet.Cell("A1").Value = "x";
            sheet.PageSetup.AddHorizontalPageBreak(32);
            sheet.PageSetup.AddVerticalPageBreak(4);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.Single().Worksheet;

        var rowBreak = worksheet!.GetFirstChild<RowBreaks>()!.Elements<Break>().Single();
        await Assert.That(rowBreak.Id!.Value).IsEqualTo(32u);
        await Assert.That(rowBreak.Max!.Value).IsEqualTo(16383u); // last column, 0-based XFD

        var columnBreak = worksheet.GetFirstChild<ColumnBreaks>()!.Elements<Break>().Single();
        await Assert.That(columnBreak.Id!.Value).IsEqualTo(4u);
        await Assert.That(columnBreak.Max!.Value).IsEqualTo(1048575u); // last row, 0-based
    }

    [Test]
    public async Task Page_break_lists_cannot_be_edited_around_the_page_setup()
    {
        // AddHorizontalPageBreak and AddVerticalPageBreak keep the breaks sorted and free of
        // duplicates. A list the caller can edit directly lets both slip through.
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        await Assert.That(sheet.PageSetup.RowBreaks is ICollection<int> { IsReadOnly: false }).IsFalse();
        await Assert.That(sheet.PageSetup.ColumnBreaks is ICollection<int> { IsReadOnly: false }).IsFalse();
    }

    [Test]
    public async Task RemoveHorizontalPageBreak_removes_only_that_break()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        sheet.PageSetup.AddHorizontalPageBreak(5);
        sheet.PageSetup.AddHorizontalPageBreak(10);

        await Assert.That(sheet.PageSetup.RemoveHorizontalPageBreak(5)).IsTrue();
        await Assert.That(sheet.PageSetup.RemoveHorizontalPageBreak(7)).IsFalse();
        await Assert.That(sheet.PageSetup.RowBreaks).IsEquivalentTo([10]);
    }

    [Test]
    public async Task RemoveVerticalPageBreak_removes_only_that_break()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        sheet.PageSetup.AddVerticalPageBreak(5);
        sheet.PageSetup.AddVerticalPageBreak(10);

        await Assert.That(sheet.PageSetup.RemoveVerticalPageBreak(5)).IsTrue();
        await Assert.That(sheet.PageSetup.RemoveVerticalPageBreak(7)).IsFalse();
        await Assert.That(sheet.PageSetup.ColumnBreaks).IsEquivalentTo([10]);
    }

    [Test]
    public async Task ClearHorizontalPageBreaks_leaves_the_vertical_ones()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        sheet.PageSetup.AddHorizontalPageBreak(5);
        sheet.PageSetup.AddVerticalPageBreak(3);

        sheet.PageSetup.ClearHorizontalPageBreaks();

        await Assert.That(sheet.PageSetup.RowBreaks).IsEmpty();
        await Assert.That(sheet.PageSetup.ColumnBreaks).IsEquivalentTo([3]);
    }

    [Test]
    public async Task ClearVerticalPageBreaks_leaves_the_horizontal_ones()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        sheet.PageSetup.AddHorizontalPageBreak(5);
        sheet.PageSetup.AddVerticalPageBreak(3);

        sheet.PageSetup.ClearVerticalPageBreaks();

        await Assert.That(sheet.PageSetup.ColumnBreaks).IsEmpty();
        await Assert.That(sheet.PageSetup.RowBreaks).IsEquivalentTo([5]);
    }

    [Test]
    public async Task A_break_removed_from_a_loaded_sheet_is_not_written_back()
    {
        using var original = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet = wb.AddWorksheet("Sheet1");
            sheet.Cell("A1").Value = "x";
            sheet.PageSetup.AddHorizontalPageBreak(5);
            sheet.PageSetup.AddHorizontalPageBreak(10);
            wb.SaveAs(original);
        }

        original.Position = 0;
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Sheet1").PageSetup.RemoveHorizontalPageBreak(5);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var ids = doc.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<RowBreaks>()!
            .Elements<Break>().Select(b => b.Id!.Value);
        await Assert.That(ids).IsEquivalentTo([10u]);
    }
}
