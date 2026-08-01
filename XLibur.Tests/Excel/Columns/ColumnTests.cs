using System.Linq;
using XLibur.Excel;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Columns;

public class ColumnTests
{
    [Test]
    public async Task ColumnUsed()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(2, 1).SetValue("Test");
        ws.Cell(3, 1).SetValue("Test");

        var fromColumn = ws.Column(1).ColumnUsed();
        await Assert.That(fromColumn.RangeAddress.ToStringRelative()).IsEqualTo("A2:A3");

        var fromRange = ws.Range("A1:A5").FirstColumn().ColumnUsed();
        await Assert.That(fromRange.RangeAddress.ToStringRelative()).IsEqualTo("A2:A3");
    }

    [Test]
    public async Task ColumnsUsedIsFast()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        var columnsUsed = ws.Row(1).AsRange().ColumnsUsed();
        await Assert.That(columnsUsed.Count()).IsEqualTo(1);
    }

    [Test]
    public async Task CopyColumn()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Test").Style.Font.SetBold();
        ws.FirstColumn().CopyTo(ws.Column(2));

        await Assert.That(ws.Cell("B1").Style.Font.Bold).IsTrue();
    }

    [Test]
    public async Task InsertingColumnsBefore1()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        var column1 = ws.Column(1);
        var column2 = ws.Column(2);
        var column3 = ws.Column(3);

        var columnIns = ws.Column(1).InsertColumnsBefore(1).First();

        await Assert.That(ws.Column(1).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Column(1).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Column(1).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);

        await Assert.That(ws.Column(2).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(2).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(2).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(3).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Column(3).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Column(3).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Column(4).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(4).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(4).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(3).Cell(2).GetText()).IsEqualTo("X");

        await Assert.That(columnIns.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(columnIns.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(columnIns.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);

        await Assert.That(column1.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column1.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column1.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(column2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(column2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(column3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column2.Cell(2).GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task InsertingColumnsBefore2()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        var column1 = ws.Column(1);
        var column2 = ws.Column(2);
        var column3 = ws.Column(3);

        var columnIns = ws.Column(2).InsertColumnsBefore(1).First();

        await Assert.That(ws.Column(1).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(1).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(1).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(2).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(2).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(2).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(3).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Column(3).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Column(3).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Column(4).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(4).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(4).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(3).Cell(2).GetText()).IsEqualTo("X");

        await Assert.That(columnIns.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(columnIns.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(columnIns.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column1.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column1.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column1.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(column2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(column2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(column3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column2.Cell(2).GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task InsertingColumnsBefore3()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Columns("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Column(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        var column1 = ws.Column(1);
        var column2 = ws.Column(2);
        var column3 = ws.Column(3);

        var columnIns = ws.Column(3).InsertColumnsBefore(1).First();

        await Assert.That(ws.Column(1).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(1).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(1).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(2).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Column(2).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Column(2).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Column(3).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Column(3).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Column(3).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Column(4).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(4).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Column(4).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Column(2).Cell(2).GetText()).IsEqualTo("X");

        await Assert.That(columnIns.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(columnIns.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(columnIns.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(column1.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column1.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column1.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(column2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(column2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(column3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(column2.Cell(2).GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task NoColumnsUsed()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var count = 0;

        foreach (var row in ws.ColumnsUsed())
            count++;

        foreach (var row in ws.Range("A1:C3").ColumnsUsed())
            count++;

        await Assert.That(count).IsEqualTo(0);
    }

    [Test]
    public async Task UngroupFromAll()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        ws.Columns(1, 2).Group();
        ws.Columns(1, 2).Ungroup(true);

        await Assert.That(ws.Column(1).OutlineLevel).IsEqualTo(0);
        await Assert.That(ws.Column(2).OutlineLevel).IsEqualTo(0);
    }

    [Test]
    public async Task LastColumnUsed()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "A1";
        ws.Cell("B1").Value = "B1";
        ws.Cell("A2").Value = "A2";
        var lastCoUsed = ws.LastColumnUsed().ColumnNumber();
        await Assert.That(lastCoUsed).IsEqualTo(2);
    }

    [Test]
    public async Task NegativeColumnNumberIsInvalid()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1") as XLWorksheet;

        var column = new XLColumn(ws, -1);

        await Assert.That(column.RangeAddress.IsValid).IsFalse();
    }

    [Test]
    public async Task AssignWorksheetColumnWidthWhenAllColumnsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        var columns = ws.Columns();

        columns.Width = 100;

        await Assert.That(ws.Column("G").Width).IsEqualTo(100).Within(XLHelper.Epsilon);
        await Assert.That(ws.ColumnWidth).IsEqualTo(100).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task PreserveWorksheetColumnWidthWhenNotAllColumnsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        var defaultColumnWidth = ws.ColumnWidth;
        var columns = ws.Columns(1, XLHelper.MaxColumnNumber);

        columns.Width = 100;

        await Assert.That(ws.Column("G").Width).IsEqualTo(100).Within(XLHelper.Epsilon);
        await Assert.That(ws.ColumnWidth).IsEqualTo(defaultColumnWidth).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task PreserveWorksheetColumnWidthWhenUsedColumnsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        ws.Cells("A1:E5").Value = "Not empty";
        var defaultColumnWidth = ws.ColumnWidth;
        var columns = ws.ColumnsUsed(XLCellsUsedOptions.Contents);

        columns.Width = 100;

        await Assert.That(ws.Column("C").Width).IsEqualTo(100).Within(XLHelper.Epsilon);
        await Assert.That(ws.Column("G").Width).IsEqualTo(defaultColumnWidth).Within(XLHelper.Epsilon);
        await Assert.That(ws.ColumnWidth).IsEqualTo(defaultColumnWidth).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task ColumnsCanBeInsertedWhenDocumentHasDefinedNameWithInvalidFormula()
    {
        // Issue: InsertColumnsBefore/After fails with a ParsingException when the workbook
        // has a defined name with an invalid formula.
        var wb = new XLWorkbook();
        wb.DefinedNames.Add("TestName", XLError.NameNotRecognized.ToDisplayString());
        var ws1 = wb.AddWorksheet();
        await Assert.That(() => ws1.FirstColumn().InsertColumnsAfter(1)).ThrowsNothing();
        var ws2 = wb.AddWorksheet();
        await Assert.That(() => ws2.FirstColumn().InsertColumnsBefore(1)).ThrowsNothing();

        await Assert.That(wb.Worksheets.Count).IsEqualTo(2);
    }
}
