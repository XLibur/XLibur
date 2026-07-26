using XLibur.Excel;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class UsedAndUnusedCellsTests
{
    private XLWorkbook workbook;

    [Before(HookType.Test)]
    public void SetupWorkbook()
    {
        workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = "A1";
        ws.Cell(1, 3).Value = "C1";
        ws.Cell(2, 2).Value = "B2";
        ws.Cell(4, 1).Value = "A4";
        ws.Cell(5, 2).Value = "B5";
        ws.Cell(6, 2).Style.Fill.BackgroundColor = XLColor.Red;
    }

    [Test]
    public async Task CountUsedCellsInRow()
    {
        var i = 0;
        var row = workbook.Worksheets.First().FirstRow();
        foreach (var cell in row.Cells()) // Cells() returns UnUsed cells by default
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(2);

        i = 0;
        row = workbook.Worksheets.First().FirstRow().RowBelow();
        foreach (var cell in row.Cells())
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(1);

        i = 0;
        row = workbook.Worksheets.First().LastRowUsed(XLCellsUsedOptions.All);
        await Assert.That(row.RowNumber()).IsEqualTo(6);
        foreach (var cell in row.Cells())
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(1);

        i = 0;
        row = workbook.Worksheets.First().LastRowUsed(XLCellsUsedOptions.All);
        await Assert.That(row.RowNumber()).IsEqualTo(6);
        foreach (var cell in row.CellsUsed())
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(0);
    }

    [Test]
    [Property("Description", "See 1443")]
    public async Task FirstRowUsedRegression()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Range("B3:F6").SetValue(100);

        await Assert.That(ws.FirstRowUsed(XLCellsUsedOptions.AllContents).RowNumber()).IsEqualTo(3);
    }

    [Test]
    public async Task CountAllCellsInRow()
    {
        var i = 0;
        var row = workbook.Worksheets.First().FirstRow();
        foreach (var cell in row.Cells(false)) // All cells in range between first and last cells used
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(3);

        i = 0;
        row = workbook.Worksheets.First().FirstRow().RowBelow(); //This row has no empty cells BETWEEN used cells
        foreach (var cell in row.Cells(false))
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(1);
    }

    [Test]
    public async Task CountUsedCellsInColumn()
    {
        var i = 0;
        var column = workbook.Worksheets.First().FirstColumn();
        foreach (var cell in column.Cells()) // Cells() returns UnUsed cells by default
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(2);

        i = 0;
        column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight();
        foreach (var cell in column.Cells())
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(1);

        i = 0;
        column = workbook.Worksheets.First().Column(2);
        foreach (var cell in column.Cells())
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(3);

        i = 0;
        column = workbook.Worksheets.First().Column(2);
        foreach (var cell in column.CellsUsed())
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(2);
    }

    [Test]
    public async Task CountAllCellsInColumn()
    {
        var i = 0;
        var column = workbook.Worksheets.First().FirstColumn();
        foreach (var cell in column.Cells(false)) // All cells in range between first and last cells used
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(4);

        i = 0;
        column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight(); //This column has no empty cells BETWEEN used cells
        foreach (var cell in column.Cells(false))
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(1);
    }

    [Test]
    public async Task CountCellsInWorksheet()
    {
        var ws = workbook.Worksheets.First();
        var i = 0;

        foreach (var cell in ws.Cells()) // All cells with content or formats
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(6);
    }

    [Test]
    public async Task CountUsedCellsInWorksheet()
    {
        var ws = workbook.Worksheets.First();
        var i = 0;

        foreach (var cell in ws.CellsUsed()) // Only used cells in worksheet
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(5);
    }

    [Test]
    public async Task CountAllCellsInWorksheet()
    {
        var ws = workbook.Worksheets.First();
        var i = 0;

        foreach (var cell in ws.Cells(false)) // All cells in range between first and last cells used (cartesian product of range)
        {
            i++;
        }
        await Assert.That(i).IsEqualTo(18);
    }

    [Test]
    public async Task GetCellsUsedNonRectangular()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("page1");

        sheet.Range("C1:E1").Value = "row1";
        sheet.Range("A2:E2").Value = "row2";

        var used = sheet.RangeUsed().RangeAddress.ToString(XLReferenceStyle.A1);

        await Assert.That(used).IsEqualTo("A1:E2");
    }

    [Test]
    [Arguments(true, "A1:D2", "A1")]
    [Arguments(true, "A2:D2", "A2")]
    [Arguments(true, "A1:D2", "A1", "B2")]
    [Arguments(true, "B2:D3", "C3")]
    [Arguments(true, "B2:F4", "F4")]
    [Arguments(false, "A1:D2", "A1")]
    [Arguments(false, "A2:D2", "A2")]
    [Arguments(false, "A1:D2", "A1", "B2")]
    [Arguments(false, "B2:D3", "C3")]
    [Arguments(false, "B2:F4", "F4")]
    public async Task RangeUsedIncludesMergedCells(bool includeFormatting, string expectedRange,
        params string[] cellsWithValues)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        foreach (var cellAddress in cellsWithValues)
        {
            ws.Cell(cellAddress).Value = "Not empty";
        }
        ws.Range("B2:D2").Merge();

        var options = includeFormatting
            ? XLCellsUsedOptions.All
            : XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges;
        var actual = ws.RangeUsed(options).RangeAddress;

        await Assert.That(actual.ToString()).IsEqualTo(expectedRange);
    }

    [Test]
    public async Task LastCellUsedPredicateConsidersMergedRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
        ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
        ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
        ws.Range("A1:C1").Merge();
        ws.Range("A2:C2").Merge();
        ws.Range("A3:C3").Merge();

        var actual = ws.LastCellUsed(XLCellsUsedOptions.All,
            c => c.Style.Fill.BackgroundColor == XLColor.Yellow);

        await Assert.That(actual.Address.ToString()).IsEqualTo("C2");
    }

    [Test]
    public async Task FirstCellUsedPredicateConsidersMergedRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
        ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
        ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
        ws.Range("A1:C1").Merge();
        ws.Range("A2:C2").Merge();
        ws.Range("A3:C3").Merge();

        var actual = ws.FirstCellUsed(XLCellsUsedOptions.All,
            c => c.Style.Fill.BackgroundColor == XLColor.Yellow);

        await Assert.That(actual.Address.ToString()).IsEqualTo("A2");
    }

    [Test]
    public async Task ApplyingDataValidationMakesCellNotEmpty()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B2:B12").CreateDataValidation()
            .Decimal.EqualOrGreaterThan(0);

        var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();

        await Assert.That(usedCells.Count).IsEqualTo(11);
        await Assert.That(usedCells.First().Address.ToString()).IsEqualTo("B2");
        await Assert.That(usedCells.Last().Address.ToString()).IsEqualTo("B12");
    }

    [Test]
    public async Task MergeMakesCellNotEmpty()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B2:B12").Merge();

        var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();

        await Assert.That(usedCells.Count).IsEqualTo(11);
        await Assert.That(usedCells.First().Address.ToString()).IsEqualTo("B2");
        await Assert.That(usedCells.Last().Address.ToString()).IsEqualTo("B12");
    }

    [Test]
    public async Task FirstCellUsedNotHangingOnLargeCFRules()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Gold);

        var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(0);
        await Assert.That(firstCell.Address.ToString()).IsEqualTo("A1");
    }

    [Test]
    public async Task LastCellUsedNotHangingOnLargeCFRules()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Gold);

        var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(0);
        await Assert.That(lastCell.Address.ToString()).IsEqualTo(XLHelper.LastCell);
    }

    [Test]
    public async Task FirstCellUsedNotHangingOnLargeDVRules()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.CreateDataValidation().WholeNumber.GreaterThan(0);

        var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(0);
        await Assert.That(firstCell.Address.ToString()).IsEqualTo("A1");
    }

    [Test]
    public async Task LastCellUsedNotHangingOnLargeDVRules()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.CreateDataValidation().WholeNumber.GreaterThan(0);

        var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(0);
        await Assert.That(lastCell.Address.ToString()).IsEqualTo(XLHelper.LastCell);
    }

    [Test]
    public async Task FirstCellUsedNotHangingOnLargeMergedRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Merge();

        var firstCell = ws.FirstCellUsed(XLCellsUsedOptions.All);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(0);
        await Assert.That(firstCell.Address.ToString()).IsEqualTo("A1");
    }

    [Test]
    public async Task LastCellUsedNotHangingOnLargeMergedRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Merge();

        var lastCell = ws.LastCellUsed(XLCellsUsedOptions.All);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(0);
        await Assert.That(lastCell.Address.ToString()).IsEqualTo(XLHelper.LastCell);
    }
}
