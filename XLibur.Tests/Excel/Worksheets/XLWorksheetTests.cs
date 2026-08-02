using System;
using System.IO;
using System.Linq;
using System.Reflection;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.Worksheets;
// ReSharper disable once InconsistentNaming
public class XLWorksheetTests
{
    private static readonly char[] IllegalWorksheetCharacters = "\0\u0003:\\/?*[]".ToCharArray();

    [Test]
    public async Task ColumnCountTime()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var start = DateTime.Now;
        ws.ColumnCount();
        var end = DateTime.Now;
        await Assert.That((end - start).TotalMilliseconds < 500).IsTrue();
    }

    [Test]
    public async Task CopyConditionalFormatsCount()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:C3").AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Range("A1:C3").Value = 1;
        var ws2 = ws.CopyTo("Sheet2");
        await Assert.That(ws2.ConditionalFormats.Count()).IsEqualTo(1);
    }

    [Test]
    public async Task CopyColumnVisibility()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Columns(10, 20).Hide();
        ws.CopyTo("Sheet2");
        await Assert.That(wb.Worksheet("Sheet2").Column(10).IsHidden).IsTrue();
    }

    [Test]
    public async Task CopyRowVisibility()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Rows(2, 5).Hide();
        ws.CopyTo("Sheet2");
        await Assert.That(wb.Worksheet("Sheet2").Row(4).IsHidden).IsTrue();
    }

    [Test]
    public async Task DeletingSheets1()
    {
        var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet3");
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet1", 1);

        wb.Worksheet("Sheet3").Delete();

        await Assert.That(wb.Worksheet(1).Name).IsEqualTo("Sheet1");
        await Assert.That(wb.Worksheet(2).Name).IsEqualTo("Sheet2");
        await Assert.That(wb.Worksheets.Count).IsEqualTo(2);
    }

    [Test]
    public async Task InsertingSheets1()
    {
        var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet3");

        await Assert.That(wb.Worksheet(1).Name).IsEqualTo("Sheet1");
        await Assert.That(wb.Worksheet(2).Name).IsEqualTo("Sheet2");
        await Assert.That(wb.Worksheet(3).Name).IsEqualTo("Sheet3");
    }

    [Test]
    public async Task InsertingSheets2()
    {
        var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet1", 1);
        wb.Worksheets.Add("Sheet3");

        await Assert.That(wb.Worksheet(1).Name).IsEqualTo("Sheet1");
        await Assert.That(wb.Worksheet(2).Name).IsEqualTo("Sheet2");
        await Assert.That(wb.Worksheet(3).Name).IsEqualTo("Sheet3");
    }

    [Test]
    public async Task InsertingSheets3()
    {
        var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet3");
        wb.Worksheets.Add("Sheet2", 1);
        wb.Worksheets.Add("Sheet1", 1);

        await Assert.That(wb.Worksheet(1).Name).IsEqualTo("Sheet1");
        await Assert.That(wb.Worksheet(2).Name).IsEqualTo("Sheet2");
        await Assert.That(wb.Worksheet(3).Name).IsEqualTo("Sheet3");
    }

    [Test]
    public async Task InsertingSheets4()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add();

        await Assert.That(ws1.Name).IsEqualTo("Sheet1");
        ws1.Name = "shEEt1";

        var ws2 = wb.Worksheets.Add();
        await Assert.That(ws2.Name).IsEqualTo("Sheet2");

        wb.Worksheets.Add("SHEET4");

        await Assert.That(wb.Worksheets.Add().Name).IsEqualTo("Sheet5");
        await Assert.That(wb.Worksheets.Add().Name).IsEqualTo("Sheet6");

        wb.Worksheets.Add(1);

        await Assert.That(wb.Worksheet(1).Name).IsEqualTo("Sheet7");
    }

    [Test]
    public async Task SheetIdIsNotReused()
    {
        using var wb = new XLWorkbook();
        var ws1 = (XLWorksheet)wb.AddWorksheet();
        var ws2 = (XLWorksheet)wb.AddWorksheet();
        var ws3 = (XLWorksheet)wb.AddWorksheet();

        await Assert.That(ws1.SheetId).IsEqualTo(ExpectedCellValue.From(1));
        await Assert.That(ws2.SheetId).IsEqualTo(2u);
        await Assert.That(ws3.SheetId).IsEqualTo(3u);

        ws3.Delete();
        var ws4 = (XLWorksheet)wb.AddWorksheet();
        await Assert.That(ws4.SheetId).IsEqualTo(4u);
    }

    [Test]
    public async Task AddingDuplicateSheetNameThrowsException()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");

        await Assert.That(() => wb.AddWorksheet("Sheet1")).Throws<ArgumentException>();

        // Sheet names are case-insensitive
        await Assert.That(() => wb.AddWorksheet("sheet1")).Throws<ArgumentException>();
    }

    [Test]
    public async Task MergedRanges()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("A1:B2").Merge();
        ws.Range("C1:D3").Merge();
        ws.Range("D2:E2").Merge();

        await Assert.That(ws.MergedRanges.Count).IsEqualTo(2);
        await Assert.That(ws.MergedRanges.First().RangeAddress.ToStringRelative()).IsEqualTo("A1:B2");
        await Assert.That(ws.MergedRanges.Last().RangeAddress.ToStringRelative()).IsEqualTo("D2:E2");

        await Assert.That(ws.Cell("A2").MergedRange().RangeAddress.ToStringRelative()).IsEqualTo("A1:B2");
        await Assert.That(ws.Cell("D2").MergedRange().RangeAddress.ToStringRelative()).IsEqualTo("D2:E2");

        await Assert.That(ws.Cell("Z10").MergedRange()).IsNull();
    }

    [Test]
    public async Task RowCountTime()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var start = DateTime.Now;
        ws.RowCount();
        var end = DateTime.Now;
        await Assert.That((end - start).TotalMilliseconds < 500).IsTrue();
    }

    [Test]
    public async Task SheetsWithCommas()
    {
        using var wb = new XLWorkbook();
        const string sourceSheetName = "Sheet1, Sheet3";
        var ws = wb.Worksheets.Add(sourceSheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("B2").Value = 3;

        ws = wb.Worksheets.Add("Formula");
        ws.FirstCell().FormulaA1 = string.Format("=SUM('{0}'!A1:A2,'{0}'!B1:B2)", sourceSheetName);

        var value = ws.FirstCell().Value;
        await Assert.That(value).IsEqualTo(6);
    }

    [Test]
    public async Task CanRenameWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet1");
        var ws2 = wb.AddWorksheet("Sheet2");

        ws1.Name = "New sheet name";
        await Assert.That(ws1.Name).IsEqualTo("New sheet name");

        ws2.Name = "sheet2";
        await Assert.That(ws2.Name).IsEqualTo("sheet2");

        await Assert.That(() => ws1.Name = "SHEET2").Throws<ArgumentException>();
    }

    [Test]
    public async Task TryGetWorksheet()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");

        await Assert.That(wb.Worksheets.TryGetWorksheet("Sheet1", out _)).IsTrue();
        await Assert.That(wb.Worksheets.TryGetWorksheet("sheet1", out _)).IsTrue();
        await Assert.That(wb.Worksheets.TryGetWorksheet("sHEeT1", out _)).IsTrue();
        await Assert.That(wb.Worksheets.TryGetWorksheet("Sheeeet2", out _)).IsFalse();

        await Assert.That(wb.TryGetWorksheet("Sheet1", out IXLWorksheet _)).IsTrue();
        await Assert.That(wb.TryGetWorksheet("sheet1", out IXLWorksheet _)).IsTrue();
        await Assert.That(wb.TryGetWorksheet("sHEeT1", out IXLWorksheet _)).IsTrue();
        await Assert.That(wb.TryGetWorksheet("Sheeeet2", out IXLWorksheet _)).IsFalse();
    }

    [Test]
    public async Task HideWorksheet()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("VisibleSheet");
            wb.Worksheets.Add("HiddenSheet").Hide();
            wb.SaveAs(ms);
        }

        // unhide the hidden sheet
        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Worksheet("VisibleSheet").Visibility).IsEqualTo(XLWorksheetVisibility.Visible);
            await Assert.That(wb.Worksheet("HiddenSheet").Visibility).IsEqualTo(XLWorksheetVisibility.Hidden);

            var ws = wb.Worksheet("HiddenSheet");
            ws.Unhide().Name = "NoAlsoVisible";

            await Assert.That(ws.Visibility).IsEqualTo(XLWorksheetVisibility.Visible);

            wb.Save();
        }

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Worksheet("VisibleSheet").Visibility).IsEqualTo(XLWorksheetVisibility.Visible);
            await Assert.That(wb.Worksheet("NoAlsoVisible").Visibility).IsEqualTo(XLWorksheetVisibility.Visible);
        }
    }

    [Test]
    public async Task CanCopySheetsWithAllAnchorTypes()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var copy1 = ws.CopyTo("Copy1");

        var ws2 = wb.Worksheets.Skip(1).First();
        var copy2 = ws2.CopyTo("Copy2");

        var ws3 = wb.Worksheets.Skip(2).First();
        var copy3 = ws3.CopyTo("Copy3");
        var copy4 = ws3.CopyTo("Copy4");

        await Assert.That(copy1.Pictures.Count).IsEqualTo(ws.Pictures.Count);
        await Assert.That(copy2.Pictures.Count).IsEqualTo(ws2.Pictures.Count);
        await Assert.That(copy3.Pictures.Count).IsEqualTo(ws3.Pictures.Count);
        await Assert.That(copy4.Pictures.Count).IsEqualTo(ws3.Pictures.Count);
    }

    [Test]
    public async Task CannotCopyDeletedWorksheet()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        var ws = wb.AddWorksheet("Sheet2");

        ws.Delete();
        await Assert.That(() => ws.CopyTo("Copy of Sheet2")).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task WorksheetNameCannotStartWithApostrophe()
    {
        var title = "'StartsWithApostrophe";
        Action addWorksheet = () =>
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add(title);
        };

        await Assert.That(addWorksheet).Throws<ArgumentException>();
    }

    [Test]
    public async Task WorksheetNameCannotEndWithApostrophe()
    {
        var title = "EndsWithApostrophe'";
        Action addWorksheet = () =>
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add(title);
        };

        await Assert.That(addWorksheet).Throws<ArgumentException>();
    }

    [Test]
    public async Task WorksheetNameCannotBeEmpty()
    {
        await Assert.That(() => new XLWorkbook().AddWorksheet(" ")).Throws<ArgumentException>();
    }

    [Test]
    [MethodDataSource(nameof(IllegalWorksheetCharacters))]
    public async Task WorksheetNameCannotContainIllegalCharacters(char c)
    {
        var proposedName = $"Sheet{c}Name";
        await Assert.That(() => new XLWorkbook().AddWorksheet(proposedName)).Throws<ArgumentException>();
    }

    [Test]
    public async Task WorksheetNameCanContainApostrophe()
    {
        var title = "With'Apostrophe";
        var savedTitle = "";
        Action saveAndOpenWorkbook = () =>
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add(title);
                wb.Worksheets.First().Cell(1, 1).FormulaA1 = $"{title}!A2";
                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                savedTitle = wb.Worksheets.First().Name;
            }
        };

        await Assert.That(saveAndOpenWorkbook).ThrowsNothing();
        await Assert.That(savedTitle).IsEqualTo(title);
    }

    [Test]
    public async Task CopyWorksheetPreservesContents()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Cell("A1").Value = "A1 value";
        ws1.Cell("A2").Value = 100;
        ws1.Cell("D4").Value = new DateTime(2018, 5, 1, 0, 0, 0, DateTimeKind.Unspecified);

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.Cell("A1").Value).IsEqualTo("A1 value");
        await Assert.That(ws2.Cell("A2").Value).IsEqualTo(100);
        await Assert.That(ws2.Cell("D4").Value).IsEqualTo(new DateTime(2018, 5, 1, 0, 0, 0, DateTimeKind.Unspecified));
    }

    [Test]
    public async Task CopyWorksheetPreservesFormulae()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Cell("A1").FormulaA1 = "10*10";
        ws1.Cell("A2").FormulaA1 = "A1 * 2";

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.Cell("A1").FormulaA1).IsEqualTo("10*10");
        await Assert.That(ws2.Cell("A2").FormulaA1).IsEqualTo("A1 * 2");
    }

    [Test]
    public async Task CopyWorksheetPreservesRowHeights()
    {
        using var wb1 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");
        using var wb2 = new XLWorkbook();
        ws1.RowHeight = 55;
        ws1.Row(2).Height = 0;
        ws1.Row(3).Height = 20;

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.RowHeight).IsEqualTo(ws1.RowHeight);
        for (var i = 1; i <= 3; i++)
        {
            await Assert.That(ws2.Row(i).Height).IsEqualTo(ws1.Row(i).Height);
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesColumnWidths()
    {
        using var wb1 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");
        using var wb2 = new XLWorkbook();
        ws1.ColumnWidth = 160;
        ws1.Column(2).Width = 0;
        ws1.Column(3).Width = 240;

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.ColumnWidth).IsEqualTo(ws1.ColumnWidth);
        for (var i = 1; i <= 3; i++)
        {
            await Assert.That(ws2.Column(i).Width).IsEqualTo(ws1.Column(i).Width);
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesMergedCells()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Range("A:A").Merge();
        ws1.Range("B1:C2").Merge();

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.MergedRanges.Count).IsEqualTo(ws1.MergedRanges.Count);
        for (var i = 0; i < ws1.MergedRanges.Count; i++)
        {
            await Assert.That(ws2.MergedRanges.ElementAt(i).RangeAddress.ToString()).IsEqualTo(ws1.MergedRanges.ElementAt(i).RangeAddress.ToString());
        }
    }

    [Test]
    public async Task Copy_sheet_across_workbooks_preserves_defined_names()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
        ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.DefinedNames.Count()).IsEqualTo(ws1.DefinedNames.Count());
        for (var i = 0; i < ws1.DefinedNames.Count(); i++)
        {
            var nr1 = ws1.DefinedNames.ElementAt(i);
            var nr2 = ws2.DefinedNames.ElementAt(i);
            await Assert.That(nr2.Ranges.ToString()).IsEqualTo(nr1.Ranges.ToString());
            await Assert.That(nr2.Scope).IsEqualTo(nr1.Scope);
            await Assert.That(nr2.Name).IsEqualTo(nr1.Name);
            await Assert.That(nr2.Visible).IsEqualTo(nr1.Visible);
            await Assert.That(nr2.Comment).IsEqualTo(nr1.Comment);
        }
    }

    [Test]
    public async Task Copying_sheet_inside_workbook_makes_copies_of_sheet_scoped_defined_names()
    {
        using var wb1 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Range("A1:A2").AddToNamed("GLOBAL", XLScope.Workbook);
        ws1.Ranges("B1:B2,D1:D2").AddToNamed("LOCAL", XLScope.Worksheet);

        var ws2 = ws1.CopyTo("Copy");

        await Assert.That(ws2.DefinedNames.Count()).IsEqualTo(ws1.DefinedNames.Count());
        for (var i = 0; i < ws1.DefinedNames.Count(); i++)
        {
            var nr1 = ws1.DefinedNames.ElementAt(i);
            var nr2 = ws2.DefinedNames.ElementAt(i);

            await Assert.That(nr2.Scope).IsEqualTo(XLScope.Worksheet);

            await Assert.That(nr2.Ranges.ToString()).IsEqualTo(nr1.Ranges.ToString());
            await Assert.That(nr2.Name).IsEqualTo(nr1.Name);
            await Assert.That(nr2.Visible).IsEqualTo(nr1.Visible);
            await Assert.That(nr2.Comment).IsEqualTo(nr1.Comment);
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesStyles()
    {
        using var ms = new MemoryStream();
        using var wb1 = new XLWorkbook();

        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws1.Range("A1:B2").Style.Font.FontSize = 25;
        ws1.Cell("C3").Style.Fill.BackgroundColor = XLColor.Red;
        ws1.Cell("C4").Style.Fill.BackgroundColor = XLColor.AliceBlue;
        ws1.Cell("C4").Value = "Non empty";

        using (var wb2 = new XLWorkbook())
        {
            var ws2 = ws1.CopyTo(wb2, "Copy");
            await AssertStylesAreEqual(ws1, ws2);
            wb2.SaveAs(ms);
        }

        using (var wb2 = new XLWorkbook(ms))
        {
            var ws2 = wb2.Worksheet("Copy");
            await AssertStylesAreEqual(ws1, ws2);
        }

        return;

        async Task AssertStylesAreEqual(IXLWorksheet ws1Assert, IXLWorksheet ws2)
        {
            await Assert.That((ws2.Style as XLStyle)!.Value).IsEqualTo((ws1Assert.Style as XLStyle)!.Value).Because("Worksheet styles differ");
            var cellsUsed = ws1Assert.Range(ws1Assert.FirstCell(), ws1Assert.LastCellUsed()).Cells();
            foreach (var cell in cellsUsed)
            {
                var style1 = (cell.Style as XLStyle)!.Value;
                var style2 = (ws2.Cell(cell.Address.ToString()).Style as XLStyle)!.Value;
                await Assert.That(style2).IsEqualTo(style1).Because($"Cell {cell.Address} styles differ");
            }
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesConditionalFormats()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Range("A:A").AddConditionalFormat()
            .WhenContains("0").Fill.SetBackgroundColor(XLColor.Red);
        var cf = ws1.Range("B1:C2").AddConditionalFormat();
        cf.Ranges.Add(ws1.Range("D4:D5"));
        cf.WhenEqualOrGreaterThan(100).Font.SetBold();

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.ConditionalFormats.Count()).IsEqualTo(ws1.ConditionalFormats.Count());
        for (var i = 0; i < ws1.ConditionalFormats.Count(); i++)
        {
            var original = ws1.ConditionalFormats.ElementAt(i);
            var copy = ws2.ConditionalFormats.ElementAt(i);
            await Assert.That(copy.Ranges.Count).IsEqualTo(original.Ranges.Count);
            for (var j = 0; j < original.Ranges.Count; j++)
            {
                await Assert.That(copy.Ranges.ElementAt(j).RangeAddress.ToString(XLReferenceStyle.A1, false)).IsEqualTo(original.Ranges.ElementAt(j).RangeAddress.ToString(XLReferenceStyle.A1, false));
            }

            await Assert.That((copy.Style as XLStyle)!.Value).IsEqualTo((original.Style as XLStyle)!.Value);
            await Assert.That(copy.Values.Single().Value.Value).IsEqualTo(original.Values.Single().Value.Value);
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesTables()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Cell("A2").Value = "Name";
        ws1.Cell("B2").Value = "Count";
        ws1.Cell("A3").Value = "John Smith";
        ws1.Cell("B3").Value = 50;
        ws1.Cell("A4").Value = "Ivan Ivanov";
        ws1.Cell("B4").Value = 40;
        var table1 = ws1.Range("A2:B4").CreateTable("Test_table_1");
        table1
            .SetShowAutoFilter(true)
            .SetShowTotalsRow(true)
            .SetEmphasizeFirstColumn(true)
            .SetShowColumnStripes(true)
            .SetShowRowStripes(true);
        table1.Theme = XLTableTheme.TableStyleDark8;
        table1.Field(1).TotalsRowFunction = XLTotalsRowFunction.Sum;

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.Tables.Count()).IsEqualTo(ws1.Tables.Count());
        for (var i = 0; i < ws1.Tables.Count(); i++)
        {
            var original = ws1.Tables.ElementAt(i);
            var copy = ws2.Tables.ElementAt(i);
            await Assert.That(copy.RangeAddress.ToString(XLReferenceStyle.A1, false)).IsEqualTo(original.RangeAddress.ToString(XLReferenceStyle.A1, false));
            await Assert.That(copy.Fields.Count()).IsEqualTo(original.Fields.Count());
            for (var j = 0; j < original.Fields.Count(); j++)
            {
                var originalField = original.Fields.ElementAt(j);
                var copyField = copy.Fields.ElementAt(j);
                await Assert.That(copyField.Name).IsEqualTo(originalField.Name);
                await Assert.That(copyField.TotalsRowFormulaA1).IsEqualTo(originalField.TotalsRowFormulaA1);
                await Assert.That(copyField.TotalsRowFunction).IsEqualTo(originalField.TotalsRowFunction);
            }

            await Assert.That(copy.Name).IsEqualTo(original.Name);
            await Assert.That(copy.ShowAutoFilter).IsEqualTo(original.ShowAutoFilter);
            await Assert.That(copy.ShowColumnStripes).IsEqualTo(original.ShowColumnStripes);
            await Assert.That(copy.ShowHeaderRow).IsEqualTo(original.ShowHeaderRow);
            await Assert.That(copy.ShowRowStripes).IsEqualTo(original.ShowRowStripes);
            await Assert.That(copy.ShowTotalsRow).IsEqualTo(original.ShowTotalsRow);
            await Assert.That((copy.Style as XLStyle)!.Value).IsEqualTo((original.Style as XLStyle)!.Value);
            await Assert.That(copy.Theme).IsEqualTo(original.Theme);
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesDataValidation()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        var dv1 = ws1.Range("A:A").CreateDataValidation();
        dv1.WholeNumber.EqualTo(2);
        dv1.ErrorStyle = XLErrorStyle.Warning;
        dv1.ErrorTitle = "Number out of range";
        dv1.ErrorMessage = "This cell only allows the number 2.";

        var dv2 = ws1.Ranges("B2:C3,D4:E5").CreateDataValidation();
        dv2.Decimal.GreaterThan(5);
        dv2.ErrorStyle = XLErrorStyle.Stop;
        dv2.ErrorTitle = "Decimal number out of range";
        dv2.ErrorMessage = "This cell only allows decimals greater than 5.";

        var dv3 = ws1.Cell("D1").CreateDataValidation();
        dv3.TextLength.EqualOrLessThan(10);
        dv3.ErrorStyle = XLErrorStyle.Information;
        dv3.ErrorTitle = "Text length out of range";
        dv3.ErrorMessage = "You entered more than 10 characters.";

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.DataValidations.Count()).IsEqualTo(ws1.DataValidations.Count());
        for (var i = 0; i < ws1.DataValidations.Count(); i++)
        {
            var original = ws1.DataValidations.ElementAt(i);
            var copy = ws2.DataValidations.ElementAt(i);

            var originalRanges = string.Join(",", original.Ranges.Select(r => r.RangeAddress.ToString()));
            var copyRanges = string.Join(",", original.Ranges.Select(r => r.RangeAddress.ToString()));

            await Assert.That(copyRanges).IsEqualTo(originalRanges);
            await Assert.That(copy.AllowedValues).IsEqualTo(original.AllowedValues);
            await Assert.That(copy.Operator).IsEqualTo(original.Operator);
            await Assert.That(copy.ErrorStyle).IsEqualTo(original.ErrorStyle);
            await Assert.That(copy.ErrorTitle).IsEqualTo(original.ErrorTitle);
            await Assert.That(copy.ErrorMessage).IsEqualTo(original.ErrorMessage);
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesPictures()
    {
        using var ms = new MemoryStream();
        using var imageStream = System.Reflection.Assembly.GetAssembly(typeof(XLibur.Examples.BasicTable))
            .GetManifestResourceStream("XLibur.Examples.Resources.SampleImage.jpg");
        using var wb1 = new XLWorkbook();

        var ws1 = wb1.Worksheets.Add("Original");

        ws1.AddPicture(imageStream, "MyPicture")
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(50, 50)
            .WithSize(200, 200);

        using (var wb2 = new XLWorkbook())
        {
            var ws2 = ws1.CopyTo(wb2, "Copy");
            await AssertPicturesAreEqual(ws1, ws2);
            wb2.SaveAs(ms);
        }

        using (var wb2 = new XLWorkbook(ms))
        {
            var ws2 = wb2.Worksheet("Copy");
            await AssertPicturesAreEqual(ws1, ws2);
        }

        async Task AssertPicturesAreEqual(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            await Assert.That(ws2.Pictures.Count).IsEqualTo(ws1.Pictures.Count);

            for (var i = 0; i < ws1.Pictures.Count; i++)
            {
                var original = ws1.Pictures.ElementAt(i);
                var copy = ws2.Pictures.ElementAt(i);
                await Assert.That(copy.Worksheet).IsEqualTo(ws2);

                await Assert.That(copy.Format).IsEqualTo(original.Format);
                await Assert.That(copy.Height).IsEqualTo(original.Height);
                await Assert.That(copy.Id).IsEqualTo(original.Id);
                await Assert.That(copy.Left).IsEqualTo(original.Left);
                await Assert.That(copy.Name).IsEqualTo(original.Name);
                await Assert.That(copy.Placement).IsEqualTo(original.Placement);
                await Assert.That(copy.Top).IsEqualTo(original.Top);
                await Assert.That(copy.TopLeftCell.Address.ToString()).IsEqualTo(original.TopLeftCell.Address.ToString());
                await Assert.That(copy.Width).IsEqualTo(original.Width);
                await Assert.That(copy.ImageStream.ToArray()).IsEquivalentTo(original.ImageStream.ToArray(), CollectionOrdering.Matching).Because("Image streams differ");
            }
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesPivotTables()
    {
        using var ms = new MemoryStream();
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);

        var ws1 = wb.Worksheet("pvt1");
        var copyOfws1 = ws1.CopyTo("CopyOfPvt1");

        await AssertPivotTablesAreEqual(ws1, copyOfws1);

        using (var wb2 = new XLWorkbook())
        {
            // We need to  copy the source too. Cross workbook references don't work yet.
            wb.Worksheet("PastrySalesData").CopyTo(wb2);
            var ws2 = ws1.CopyTo(wb2, "Copy");
            await AssertPivotTablesAreEqual(ws1, ws2);
            wb2.SaveAs(ms);
        }

        using (var wb2 = new XLWorkbook(ms))
        {
            var ws2 = wb2.Worksheet("Copy");
            await AssertPivotTablesAreEqual(ws1, ws2);
        }

        async Task AssertPivotTablesAreEqual(IXLWorksheet ws1Assert, IXLWorksheet ws2)
        {
            await Assert.That(ws2.PivotTables.Count()).IsEqualTo(ws1Assert.PivotTables.Count());

            var comparer = new PivotTableComparer();

            for (var i = 0; i < ws1Assert.PivotTables.Count(); i++)
            {
                var original = ws1Assert.PivotTables.ElementAt(i).CastTo<XLPivotTable>();
                var copy = ws2.PivotTables.ElementAt(i).CastTo<XLPivotTable>();

                await Assert.That(copy.Worksheet).IsEqualTo(ws2);

                await Assert.That(comparer.Equals(original, copy)).IsTrue();
            }
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesSelectedRanges()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.SelectedRanges.RemoveAll();
        ws1.SelectedRanges.Add(ws1.Range("E12:H20"));
        ws1.SelectedRanges.Add(ws1.Range("B:B"));
        ws1.SelectedRanges.Add(ws1.Range("3:6"));

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.SelectedRanges.Count).IsEqualTo(ws1.SelectedRanges.Count);
        for (var i = 0; i < ws1.SelectedRanges.Count; i++)
        {
            await Assert.That(ws2.SelectedRanges.ElementAt(i).RangeAddress.ToString()).IsEqualTo(ws1.SelectedRanges.ElementAt(i).RangeAddress.ToString());
        }
    }

    [Test]
    public async Task CopyWorksheetPreservesPageSetup()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.PageSetup.AddHorizontalPageBreak(15);
        ws1.PageSetup.AddVerticalPageBreak(5);
        ws1.PageSetup
            .SetBlackAndWhite()
            .SetCenterHorizontally()
            .SetCenterVertically()
            .SetFirstPageNumber(200)
            .SetPageOrientation(XLPageOrientation.Landscape)
            .SetPaperSize(XLPaperSize.A5Paper)
            .SetScale(89)
            .SetShowGridlines()
            .SetHorizontalDpi(200)
            .SetVerticalDpi(300)
            .SetPagesTall(5)
            .SetPagesWide(2)
            .SetColumnsToRepeatAtLeft(1, 3);
        ws1.PageSetup.PrintAreas.Clear();
        ws1.PageSetup.PrintAreas.Add("A1:Z200");
        ws1.PageSetup.Margins.SetBottom(5).SetTop(6).SetLeft(7).SetRight(8).SetFooter(9).SetHeader(10);
        ws1.PageSetup.Header.Left.AddText(XLHFPredefinedText.FullPath, XLHFOccurrence.AllPages);
        ws1.PageSetup.Footer.Right.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.OddPages);

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.PageSetup.FirstRowToRepeatAtTop).IsEqualTo(ws1.PageSetup.FirstRowToRepeatAtTop);
        await Assert.That(ws2.PageSetup.LastRowToRepeatAtTop).IsEqualTo(ws1.PageSetup.LastRowToRepeatAtTop);
        await Assert.That(ws2.PageSetup.FirstColumnToRepeatAtLeft).IsEqualTo(ws1.PageSetup.FirstColumnToRepeatAtLeft);
        await Assert.That(ws2.PageSetup.LastColumnToRepeatAtLeft).IsEqualTo(ws1.PageSetup.LastColumnToRepeatAtLeft);
        await Assert.That(ws2.PageSetup.PageOrientation).IsEqualTo(ws1.PageSetup.PageOrientation);
        await Assert.That(ws2.PageSetup.PagesWide).IsEqualTo(ws1.PageSetup.PagesWide);
        await Assert.That(ws2.PageSetup.PagesTall).IsEqualTo(ws1.PageSetup.PagesTall);
        await Assert.That(ws2.PageSetup.Scale).IsEqualTo(ws1.PageSetup.Scale);
        await Assert.That(ws2.PageSetup.HorizontalDpi).IsEqualTo(ws1.PageSetup.HorizontalDpi);
        await Assert.That(ws2.PageSetup.VerticalDpi).IsEqualTo(ws1.PageSetup.VerticalDpi);
        await Assert.That(ws2.PageSetup.FirstPageNumber).IsEqualTo(ws1.PageSetup.FirstPageNumber);
        await Assert.That(ws2.PageSetup.CenterHorizontally).IsEqualTo(ws1.PageSetup.CenterHorizontally);
        await Assert.That(ws2.PageSetup.CenterVertically).IsEqualTo(ws1.PageSetup.CenterVertically);
        await Assert.That(ws2.PageSetup.PaperSize).IsEqualTo(ws1.PageSetup.PaperSize);
        await Assert.That(ws2.PageSetup.Margins.Bottom).IsEqualTo(ws1.PageSetup.Margins.Bottom);
        await Assert.That(ws2.PageSetup.Margins.Top).IsEqualTo(ws1.PageSetup.Margins.Top);
        await Assert.That(ws2.PageSetup.Margins.Left).IsEqualTo(ws1.PageSetup.Margins.Left);
        await Assert.That(ws2.PageSetup.Margins.Right).IsEqualTo(ws1.PageSetup.Margins.Right);
        await Assert.That(ws2.PageSetup.Margins.Footer).IsEqualTo(ws1.PageSetup.Margins.Footer);
        await Assert.That(ws2.PageSetup.Margins.Header).IsEqualTo(ws1.PageSetup.Margins.Header);
        await Assert.That(ws2.PageSetup.ScaleHFWithDocument).IsEqualTo(ws1.PageSetup.ScaleHFWithDocument);
        await Assert.That(ws2.PageSetup.AlignHFWithMargins).IsEqualTo(ws1.PageSetup.AlignHFWithMargins);
        await Assert.That(ws2.PageSetup.ShowGridlines).IsEqualTo(ws1.PageSetup.ShowGridlines);
        await Assert.That(ws2.PageSetup.ShowRowAndColumnHeadings).IsEqualTo(ws1.PageSetup.ShowRowAndColumnHeadings);
        await Assert.That(ws2.PageSetup.BlackAndWhite).IsEqualTo(ws1.PageSetup.BlackAndWhite);
        await Assert.That(ws2.PageSetup.DraftQuality).IsEqualTo(ws1.PageSetup.DraftQuality);
        await Assert.That(ws2.PageSetup.PageOrder).IsEqualTo(ws1.PageSetup.PageOrder);
        await Assert.That(ws2.PageSetup.ShowComments).IsEqualTo(ws1.PageSetup.ShowComments);
        await Assert.That(ws2.PageSetup.PrintErrorValue).IsEqualTo(ws1.PageSetup.PrintErrorValue);

        await Assert.That(ws2.PageSetup.PrintAreas.Count()).IsEqualTo(ws1.PageSetup.PrintAreas.Count());

        await Assert.That(ws2.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages)).IsEqualTo(ws1.PageSetup.Header.Left.GetText(XLHFOccurrence.AllPages));
        await Assert.That(ws2.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages)).IsEqualTo(ws1.PageSetup.Footer.Right.GetText(XLHFOccurrence.OddPages));
    }

    [Test]
    public async Task CopyWorksheetPreservesSparklineGroups()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");
        var original = ws1.SparklineGroups.Add("A1:A10", "D1:Z10")
            .SetDateRange(ws1.Range("D11:Z11"))
            .SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues.Zero)
            .SetDisplayHidden(true)
            .SetLineWeight(1.5)
            .SetShowMarkers(XLSparklineMarkers.All)
            .SetStyle(XLSparklineTheme.Colorful3)
            .SetType(XLSparklineType.Column);

        original.HorizontalAxis
            .SetColor(XLColor.Blue)
            .SetRightToLeft(true)
            .SetVisible(true);

        original.VerticalAxis
            .SetManualMin(-100.0)
            .SetManualMax(100.0);

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.SparklineGroups.Count()).IsEqualTo(1);
        var copy = ws2.SparklineGroups.Single();

        await Assert.That(copy.Count()).IsEqualTo(original.Count());
        for (var i = 0; i < original.Count(); i++)
        {
            await Assert.That(copy.ElementAt(i).Location.Worksheet).IsSameReferenceAs(ws2);
            await Assert.That(copy.ElementAt(i).SourceData.Worksheet).IsSameReferenceAs(ws2);
            await Assert.That(copy.ElementAt(i).Location.Address.ToString()).IsEqualTo(original.ElementAt(i).Location.Address.ToString());
            await Assert.That(copy.ElementAt(i).SourceData.RangeAddress.ToString()).IsEqualTo(original.ElementAt(i).SourceData.RangeAddress.ToString());
        }

        await Assert.That(copy.DateRange.RangeAddress.ToString()).IsEqualTo(original.DateRange.RangeAddress.ToString());
        await Assert.That(copy.DateRange.Worksheet).IsSameReferenceAs(ws2);

        await Assert.That(copy.DisplayEmptyCellsAs).IsEqualTo(original.DisplayEmptyCellsAs);
        await Assert.That(copy.DisplayHidden).IsEqualTo(original.DisplayHidden);
        await Assert.That(copy.LineWeight).IsEqualTo(original.LineWeight).Within(XLHelper.Epsilon);
        await Assert.That(copy.ShowMarkers).IsEqualTo(original.ShowMarkers);
        await Assert.That(copy.Style).IsEqualTo(original.Style);
        await Assert.That(copy.Style).IsNotSameReferenceAs(original.Style);
        await Assert.That(copy.Type).IsEqualTo(original.Type);

        await Assert.That(copy.HorizontalAxis.Color).IsEqualTo(original.HorizontalAxis.Color);
        await Assert.That(copy.HorizontalAxis.DateAxis).IsEqualTo(original.HorizontalAxis.DateAxis);
        await Assert.That(copy.HorizontalAxis.IsVisible).IsEqualTo(original.HorizontalAxis.IsVisible);
        await Assert.That(copy.HorizontalAxis.RightToLeft).IsEqualTo(original.HorizontalAxis.RightToLeft);

        await Assert.That(copy.VerticalAxis.ManualMax).IsEqualTo(original.VerticalAxis.ManualMax);
        await Assert.That(copy.VerticalAxis.ManualMin).IsEqualTo(original.VerticalAxis.ManualMin);
        await Assert.That(copy.VerticalAxis.MaxAxisType).IsEqualTo(original.VerticalAxis.MaxAxisType);
        await Assert.That(copy.VerticalAxis.MinAxisType).IsEqualTo(original.VerticalAxis.MinAxisType);
    }

    [Test]
    public async Task CopyWorksheetChangesAbsoluteReferencesInFormulae()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Cell("A1").FormulaA1 = "10*10";
        ws1.Cell("A2").FormulaA1 = "Original!A1 * 3";

        var ws2 = ws1.CopyTo(wb2, "Copy");

        await Assert.That(ws2.Cell("A2").FormulaA1).IsEqualTo("Copy!A1 * 3");
    }

    [Test]
    public async Task CopyWorksheetWithinWorkbookChangesReferencesOnlyInTheCopy()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Original");

        ws.Cell("A1").Value = 100;
        ws.Cell("A2").FormulaA1 = "Original!A1 * 3";

        var copy = ws.CopyTo("Copy");

        await Assert.That(copy.Cell("A2").FormulaA1).IsEqualTo("Copy!A1 * 3");
        await Assert.That(copy.Cell("A2").Value).IsEqualTo(300);

        // The original keeps pointing at itself.
        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("Original!A1 * 3");
    }

    [Test]
    public async Task CopyWorksheetKeepsReferencesToOtherSheets()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Original");
        var other = wb.Worksheets.Add("Other");
        other.Cell("A1").Value = 7;

        ws.Cell("A1").Value = 100;
        ws.Cell("A2").FormulaA1 = "Original!A1 + Other!A1";

        var copy = ws.CopyTo("Copy");

        // Only the self-reference follows the copy, 'Other' still means the sheet it names.
        await Assert.That(copy.Cell("A2").FormulaA1).IsEqualTo("Copy!A1 + Other!A1");
        await Assert.That(copy.Cell("A2").Value).IsEqualTo(107);
    }

    [Test]
    public async Task CopyWorksheetChangesQuotedSheetReferences()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("My Sheet");

        ws.Cell("A1").Value = 100;
        ws.Cell("A2").FormulaA1 = "'My Sheet'!A1 * 3";

        var copy = ws.CopyTo("The Copy");

        await Assert.That(copy.Cell("A2").FormulaA1).IsEqualTo("'The Copy'!A1 * 3");
        await Assert.That(copy.Cell("A2").Value).IsEqualTo(300);
    }

    [Test]
    public async Task CopyWorksheetToSameNameInOtherWorkbookKeepsReferences()
    {
        using var wb1 = new XLWorkbook();
        using var wb2 = new XLWorkbook();
        var ws1 = wb1.Worksheets.Add("Original");

        ws1.Cell("A1").Value = 100;
        ws1.Cell("A2").FormulaA1 = "Original!A1 * 3";

        var ws2 = ws1.CopyTo(wb2);

        await Assert.That(ws2.Name).IsEqualTo("Original");
        await Assert.That(ws2.Cell("A2").FormulaA1).IsEqualTo("Original!A1 * 3");
    }

    [Test]
    public async Task Rename_sheets_changes_sheet_references_in_formulas()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Original");

        ws.Cell("A1").FormulaA1 = "10*10";
        ws.Cell("A2").FormulaA1 = "Original!A1 * 3";
        _ = ws.Cell("A2").Value;

        ws.Name = "Renamed";

        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("Renamed!A1 * 3");
        await Assert.That(ws.Cell("A2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(300);
    }

    [Test]
    // ReSharper disable once InconsistentNaming
    public async Task RangesFromDeletedWorksheetContainREF()
    {
        using var wb1 = new XLWorkbook();
        wb1.Worksheets.Add("Sheet1");
        var ws2 = wb1.Worksheets.Add("Sheet2");
        var range = ws2.Range("A1:B2");

        ws2.Delete();

        await Assert.That(range.RangeAddress.ToString()).IsEqualTo("#REF!A1:B2");
    }

    [Test]
    public async Task InvalidRowAndColumnIndices()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(() => ws.Row(-1)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => ws.Row(XLHelper.MaxRowNumber + 1)).Throws<ArgumentOutOfRangeException>();

        await Assert.That(() => ws.Column(-1)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => ws.Column(XLHelper.MaxColumnNumber + 1)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task InvalidSelectedRangeExcluded()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var range1 = ws.Range("B2:C2");
        var range2 = ws.Range("B4:C4");
        ws.SelectedRanges.Clear();

        ws.SelectedRanges.Add(range1);
        ws.SelectedRanges.Add(range2);

        ws.Row(4).Delete();

        await Assert.That(range2.RangeAddress.IsValid).IsFalse();
        await Assert.That(ws.SelectedRanges.Single()).IsEqualTo(range1);
    }

    [Test]
    public async Task InsertColumnsDoesNotIncreaseCellsCount()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetValue(1);
        ws.Cell("AAA50").SetValue(1);
        var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(originalCount);
    }

    [Test]
    public async Task InsertRowsDoesNotIncreaseCellsCount()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetValue(1);
        ws.Cell("AAA500").SetValue(1);
        var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

        ws.Row(1).InsertRowsAbove(1);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(originalCount);
    }

    [Test]
    public async Task InsertCellsBeforeDoesNotIncreaseCellsCount()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var a1 = ws.Cell("A1").SetValue(1);
        ws.Cell("AAA50").SetValue(1);
        var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

        a1.InsertCellsBefore(1);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(originalCount);
    }

    [Test]
    public async Task InsertCellsAboveDoesNotIncreaseCellsCount()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var a1 = ws.Cell("A1").SetValue(1);
        ws.Cell("AAA500").SetValue(1);
        var originalCount = ((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count();

        a1.InsertCellsAbove(1);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(originalCount);
    }

    [Test]
    public async Task CellsShiftedTooFarRightArePurged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var a1 = ws.Cell("A1").SetValue(1);
        ws.Cell(1, XLHelper.MaxColumnNumber).SetValue(1);
        ws.Cell(2, XLHelper.MaxColumnNumber).SetValue(1);

        a1.InsertCellsBefore(1);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(2);
        ws.Column(1).InsertColumnsBefore(1);
        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task CellsShiftedTooFarDownArePurged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var a1 = ws.Cell("A1").SetValue(1);
        ws.Cell(XLHelper.MaxRowNumber, 1).SetValue(1);
        ws.Cell(XLHelper.MaxRowNumber, 2).SetValue(1);

        a1.InsertCellsAbove(1);

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(2);
        ws.Row(1).InsertRowsAbove(1);
        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.GetCells().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task MaxColumnUsedUpdatedWhenColumnDeleted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").SetValue(1);
        ws.Cell(1, XLHelper.MaxColumnNumber).SetValue(1);

        ws.Column(XLHelper.MaxColumnNumber).Delete();

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.MaxColumnUsed).IsEqualTo(3);
    }

    [Test]
    public async Task MaxRowUsedUpdatedWhenRowDeleted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A3").SetValue(1);
        ws.Cell(XLHelper.MaxRowNumber, 1).SetValue(1);

        ws.Row(XLHelper.MaxRowNumber).Delete();

        await Assert.That(((XLWorksheet)ws).Internals.CellsCollection.MaxRowUsed).IsEqualTo(3);
    }

    [Test]
    public async Task ChangeColumnStyleFirst()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("ColumnFirst");

        ws.Column(2).Style.Font.SetBold(true);
        ws.Row(2).Style.Font.SetItalic(true);

        await Assert.That(ws.Cell("B2").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("B2").Style.Font.Italic).IsTrue();
    }

    [Test]
    public async Task ChangeRowStyleFirst()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("RowFirst");

        ws.Row(2).Style.Font.SetItalic(true);
        ws.Column(2).Style.Font.SetBold(true);

        await Assert.That(ws.Cell("B2").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("B2").Style.Font.Italic).IsTrue();
    }

    [Test]
    public async Task SelectedTabIsActive_WhenInsertBefore()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws1 = wb.AddWorksheet();
            ws1.TabSelected = true;
            wb.Worksheets.Add(1);
            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws1 = wb.Worksheets.First();
            var ws2 = wb.Worksheets.Last();

            await Assert.That(ws1.TabActive).IsFalse();
            await Assert.That(ws1.TabSelected).IsFalse();
            await Assert.That(ws2.TabActive).IsTrue();
            await Assert.That(ws2.TabSelected).IsTrue();
        }
    }

    [Test]
    [Arguments("noactive_noselected.xlsx")]
    [Arguments("noactive_twoselected.xlsx")]
    [Arguments("noactive_negativeId.xlsx")]
    public async Task FirstSheetIsActive_WhenNotSpecified(string fileName)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\NoActiveSheet\" + fileName));
        using var wb = new XLWorkbook(stream);
        await Assert.That(wb.Worksheets.First().TabActive).IsTrue();
        await Assert.That(wb.Worksheets.First().Visibility).IsEqualTo(XLWorksheetVisibility.Visible);
    }

    [Test]
    [Arguments(XLCellsUsedOptions.NormalFormats, 42)]
    [Arguments(XLCellsUsedOptions.Contents, 100)]
    public async Task FirstColumnUsed_ReturnsFirstColumnWithUsedCell(XLCellsUsedOptions options, int expectedColumn)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 42).Style.Fill.SetBackgroundColor(XLColor.Green);
        ws.Cell(1, 100).SetValue(5);

        var column = ws.FirstColumnUsed(options);
        await Assert.That(column!.ColumnNumber()).IsEqualTo(expectedColumn);
    }

    [Test]
    public async Task RecalculateAllFormulas_recalculates_all_formulas_in_sheet_and_leaves_rest_dirty()
    {
        using var wb = new XLWorkbook();
        var sut = wb.AddWorksheet("sut");
        var other = wb.AddWorksheet("other");

        other.Cell("A1").Value = 7;
        other.Cell("A2").FormulaA1 = "A1+3";
        await Assert.That(other.Cell("A2").Value).IsEqualTo(10.0);

        // Change the supporting value, but without recalculation of dependent
        // formula, thus the value stays the same.
        other.Cell("A1").Value = 5;

        await Assert.That(other.Cell("A2").NeedsRecalculation).IsTrue();
        await Assert.That(other.Cell("A2").CachedValue).IsEqualTo(10.0);

        // Tested formula depends on a dirty formula from other sheet.
        sut.Cell("A1").FormulaA1 = "other!A2+5";
        sut.Cell("A2").FormulaA1 = "1+2";

        await Assert.That(sut.Cell("A1").CachedValue).IsEqualTo(Blank.Value);
        await Assert.That(sut.Cell("A2").CachedValue).IsEqualTo(Blank.Value);

        sut.RecalculateAllFormulas();

        // Formulas in other sheets kept the value - not affected by recalculation of a sut sheet.
        await Assert.That(other.Cell("A2").NeedsRecalculation).IsTrue();
        await Assert.That(other.Cell("A2").CachedValue).IsEqualTo(10.0);

        // Formulas in test sheet were recalculated - they are affected by recalculation of a sut sheet.
        await Assert.That(sut.Cell("A1").NeedsRecalculation).IsFalse();
        await Assert.That(sut.Cell("A1").CachedValue).IsEqualTo(15.0);

        await Assert.That(sut.Cell("A2").NeedsRecalculation).IsFalse();
        await Assert.That(sut.Cell("A2").CachedValue).IsEqualTo(3.0);
    }

    [Test]
    public async Task Cell_returns_cell_at_address_or_workbook_scoped_named_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("test_range", ws.Range(2, 3, 5, 7)); // C2:G5

        var cellB4 = ws.Cell("B4");
        var firstCellOfRange = ws.Cell("test_range");

        await Assert.That(cellB4.Address.ToString()).IsEqualTo("B4");
        await Assert.That(firstCellOfRange.Address.ToString()).IsEqualTo("C2");
    }

    [Test]
    public async Task Cell_throws_exception_when_address_is_not_A1_address_or_workbook_scoped_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(() => _ = ws.Cell("XFF1")).Throws<ArgumentException>();
        await Assert.That(() => _ = ws.Cell("nonexistent_range")).Throws<ArgumentException>();
    }

    [Test]
    public async Task Range_returns_range_from_a1_address_or_named_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("book_range", ws.Range(2, 3, 5, 7)); // C2:G5
        ws.DefinedNames.Add("sheet_range", ws.Range(1, 2, 3, 4)); // B1:D3

        var singleCellRange = ws.Range("B4");
        var areaCellRange = ws.Range("B4:D7");
        var bookNamedRange = ws.Range("book_range");
        var sheetNamedRange = ws.Range("sheet_range");

        await Assert.That(singleCellRange.RangeAddress.ToString()).IsEqualTo("B4:B4");
        await Assert.That(areaCellRange.RangeAddress.ToString()).IsEqualTo("B4:D7");
        await Assert.That(bookNamedRange.RangeAddress.ToString()).IsEqualTo("$C$2:$G$5");
        await Assert.That(sheetNamedRange.RangeAddress.ToString()).IsEqualTo("$B$1:$D$3");
    }

    [Test]
    public async Task Range_throws_exception_when_address_is_not_A1_address_or_named_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(() => _ = ws.Range("DEAD1")).Throws<ArgumentException>();
        await Assert.That(() => _ = ws.Range("DEAD4:BEEF10")).Throws<ArgumentException>();
        await Assert.That(() => _ = ws.Range("nonexistent_range")).Throws<ArgumentException>();
    }

    [Test]
    public async Task EnumerateUsedCells_yields_only_non_blank_value_cells_in_row_major_order()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("C1").Value = "text";
        ws.Cell("B2").Value = 3.14;
        ws.Cell("E5").Value = true;
        // Cells touched by style but with no value should not be yielded.
        ws.Cell("D4").Style.Font.Bold = true;

        var yielded = new System.Collections.Generic.List<(int Row, int Col, XLCellValue Value)>();
        foreach (var cell in ws.EnumerateUsedCells())
            yielded.Add((cell.Row, cell.Column, cell.Value));

        await Assert.That(yielded.Count).IsEqualTo(4);
        await Assert.That((yielded[0].Row, yielded[0].Col)).IsEqualTo((1, 1));
        await Assert.That(yielded[0].Value.GetNumber()).IsEqualTo(1);
        await Assert.That((yielded[1].Row, yielded[1].Col)).IsEqualTo((1, 3));
        await Assert.That(yielded[1].Value.GetText()).IsEqualTo("text");
        await Assert.That((yielded[2].Row, yielded[2].Col)).IsEqualTo((2, 2));
        await Assert.That(yielded[2].Value.GetNumber()).IsEqualTo(3.14);
        await Assert.That((yielded[3].Row, yielded[3].Col)).IsEqualTo((5, 5));
        await Assert.That(yielded[3].Value.GetBoolean()).IsTrue();
    }

    [Test]
    public async Task EnumerateUsedCells_returns_formula_cached_value_after_evaluation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        ws.Cell("B1").FormulaA1 = "=A1*10";
        // Force the formula to evaluate so its cached value is populated.
        _ = ws.Cell("B1").Value;

        XLCellValue? formulaCellValue = null;
        foreach (var cell in ws.EnumerateUsedCells())
        {
            if (cell.Row == 1 && cell.Column == 2)
                formulaCellValue = cell.Value;
        }

        await Assert.That(formulaCellValue).IsNotNull();
        await Assert.That(formulaCellValue!.Value.GetNumber()).IsEqualTo(50);
    }

    [Test]
    public async Task EnumerateUsedCells_on_empty_sheet_yields_nothing_and_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var count = 0;
        foreach (var _ in ws.EnumerateUsedCells())
            count++;

        await Assert.That(count).IsEqualTo(0);
    }

    [Test]
    public async Task EnumerateUsedCells_matches_CellsUsed_value_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "hello";
        ws.Cell("B5").Value = 42;
        ws.Cell("C10").Value = 3.14;
        ws.Cell("D2").FormulaA1 = "=B5+1";
        _ = ws.Cell("D2").Value; // populate cached value

        var usedCellAddresses = ws.CellsUsed()
            .Select(c => (c.Address.RowNumber, c.Address.ColumnNumber))
            .OrderBy(t => t.RowNumber).ThenBy(t => t.ColumnNumber)
            .ToArray();
        var enumeratedAddresses = new System.Collections.Generic.List<(int, int)>();
        foreach (var cell in ws.EnumerateUsedCells())
            enumeratedAddresses.Add((cell.Row, cell.Column));

        await Assert.That(enumeratedAddresses.ToArray()).IsEquivalentTo(usedCellAddresses, CollectionOrdering.Matching);
    }
}
