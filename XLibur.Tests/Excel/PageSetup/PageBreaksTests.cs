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

        var rowBreak = worksheet.GetFirstChild<RowBreaks>()!.Elements<Break>().Single();
        await Assert.That(rowBreak.Id!.Value).IsEqualTo(32u);
        await Assert.That(rowBreak.Max!.Value).IsEqualTo(16383u); // last column, 0-based XFD

        var columnBreak = worksheet.GetFirstChild<ColumnBreaks>()!.Elements<Break>().Single();
        await Assert.That(columnBreak.Id!.Value).IsEqualTo(4u);
        await Assert.That(columnBreak.Max!.Value).IsEqualTo(1048575u); // last row, 0-based
    }
}
