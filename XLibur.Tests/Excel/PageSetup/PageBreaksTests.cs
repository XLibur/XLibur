using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.PageSetup;

[TestFixture]
public class PageBreaksTests
{
    [Test]
    public void RowBreaksShouldBeSorted()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.PageSetup.AddHorizontalPageBreak(12);
        sheet.PageSetup.AddHorizontalPageBreak(5);
        Assert.That(sheet.PageSetup.RowBreaks, Is.EqualTo([5, 10, 12]));
    }

    [Test]
    public void ColumnBreaksShouldBeSorted()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.PageSetup.AddVerticalPageBreak(12);
        sheet.PageSetup.AddVerticalPageBreak(5);
        Assert.That(sheet.PageSetup.ColumnBreaks, Is.EqualTo([5, 10, 12]));
    }

    [Test]
    public void RowBreaksShiftWhenInsertedRowAbove()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.Row(5).InsertRowsAbove(1);
        Assert.AreEqual(11, sheet.PageSetup.RowBreaks[0]);
    }

    [Test]
    public void RowBreaksNotShiftWhenInsertedRowBelow()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.Row(15).InsertRowsAbove(1);
        Assert.AreEqual(10, sheet.PageSetup.RowBreaks[0]);
    }

    [Test]
    public void ColumnBreaksShiftWhenInsertedColumnBefore()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.Column(5).InsertColumnsBefore(1);
        Assert.AreEqual(11, sheet.PageSetup.ColumnBreaks[0]);
    }

    [Test]
    public void ColumnBreaksNotShiftWhenInsertedColumnAfter()
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.Column(15).InsertColumnsBefore(1);
        Assert.AreEqual(10, sheet.PageSetup.ColumnBreaks[0]);
    }

    [Test]
    public void PageBreaksWritePerpendicularAxisAsMax()
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
        Assert.That(rowBreak.Id!.Value, Is.EqualTo(32u));
        Assert.That(rowBreak.Max!.Value, Is.EqualTo(16383u)); // last column, 0-based XFD

        var columnBreak = worksheet.GetFirstChild<ColumnBreaks>()!.Elements<Break>().Single();
        Assert.That(columnBreak.Id!.Value, Is.EqualTo(4u));
        Assert.That(columnBreak.Max!.Value, Is.EqualTo(1048575u)); // last row, 0-based
    }
}
