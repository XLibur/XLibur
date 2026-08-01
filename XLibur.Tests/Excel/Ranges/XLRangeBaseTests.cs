using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class XLRangeBaseTests
{
    [Test]
    public async Task IsEmpty1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        var range = ws.Range("A1:B2");
        var actual = range.IsEmpty();
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        var range = ws.Range("A1:B2");
        var actual = range.IsEmpty(XLCellsUsedOptions.All);
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty3()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        var range = ws.Range("A1:B2");
        var actual = range.IsEmpty();
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty4()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        var range = ws.Range("A1:B2");
        var actual = range.IsEmpty(XLCellsUsedOptions.AllContents);
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty5()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        var range = ws.Range("A1:B2");
        var actual = range.IsEmpty(XLCellsUsedOptions.All);
        var expected = false;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty6()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Value = "X";
        var range = ws.Range("A1:B2");
        var actual = range.IsEmpty();
        var expected = false;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task SingleCell()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).Value = "Hello World!";
        wb.DefinedNames.Add("SingleCell", "Sheet1!$A$1");
        var range = wb.Range("SingleCell");
        await Assert.That(range.CellsUsed().Count()).IsEqualTo(1);
        await Assert.That(range.CellsUsed().Single().GetText()).IsEqualTo("Hello World!");
    }

    [Test]
    public async Task TableRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var rangeColumn = ws.Column(1).Column(1, 4);
        rangeColumn.Cell(1).Value = "FName";
        rangeColumn.Cell(2).Value = "John";
        rangeColumn.Cell(3).Value = "Hank";
        rangeColumn.Cell(4).Value = "Dagny";
        var table = rangeColumn.CreateTable();
        wb.DefinedNames.Add("FNameColumn", $"{table.Name}[FName]");

        var namedRange = wb.Range("FNameColumn");
        await Assert.That(namedRange.Cells().Count()).IsEqualTo(3);
        await Assert.That(namedRange.CellsUsed().Select(cell => cell.GetText()).SequenceEqual(["John", "Hank", "Dagny"])).IsTrue();
    }

    [Test]
    public async Task WsNamedCell()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
        await Assert.That(ws.Cell("TestCell").GetText()).IsEqualTo("Test");
    }

    [Test]
    public async Task WsNamedCells()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
        ws.Cell(2, 1).SetValue("B");
        var cells = ws.Cells("TestCell, A2");
        await Assert.That(cells.First().GetText()).IsEqualTo("Test");
        await Assert.That(cells.Last().GetText()).IsEqualTo("B");
    }

    [Test]
    public async Task WsNamedRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("A");
        ws.Cell(2, 1).SetValue("B");
        var original = ws.Range("A1:A2");
        original.AddToNamed("TestRange", XLScope.Worksheet);
        var named = ws.Range("TestRange");
        await Assert.That(named.RangeAddress.ToString()).IsEqualTo(original.RangeAddress.ToStringFixed());
    }

    [Test]
    public async Task WsNamedRanges()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("A");
        ws.Cell(2, 1).SetValue("B");
        ws.Cell(3, 1).SetValue("C");
        var original = ws.Range("A1:A2");
        original.AddToNamed("TestRange", XLScope.Worksheet);
        var namedRanges = ws.Ranges("TestRange, A3");
        await Assert.That(namedRanges.First().RangeAddress.ToString()).IsEqualTo(original.RangeAddress.ToStringFixed());
        await Assert.That(namedRanges.Last().RangeAddress.ToStringFixed()).IsEqualTo("$A$3:$A$3");
    }

    [Test]
    public async Task WsNamedRangesOneString()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.DefinedNames.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");
        var namedRanges = ws.Ranges("TestRange");

        await Assert.That(namedRanges.First().RangeAddress.ToStringFixed()).IsEqualTo("$A$1:$A$1");
        await Assert.That(namedRanges.Last().RangeAddress.ToStringFixed()).IsEqualTo("$A$3:$A$3");
    }

    //[Test]
    //public void WsNamedRangeLiteral()
    //{
    //    var wb = new XLWorkbook();
    //    var ws = wb.Worksheets.Add("Sheet1");
    //    ws.NamedRanges.Add("TestRange", "\"Hello\"");
    //    using (MemoryStream memoryStream = new MemoryStream())
    //    {
    //        wb.SaveAs(memoryStream, true);
    //        var wb2 = new XLWorkbook(memoryStream);
    //        var text = wb2.Worksheet("Sheet1").NamedRanges.First()
    //        memoryStream.Close();
    //    }

    //}

    [Test]
    public async Task GrowRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(ws.Cell("A1").AsRange().Grow().RangeAddress.ToString()).IsEqualTo("A1:B2");
        await Assert.That(ws.Cell("A2").AsRange().Grow().RangeAddress.ToString()).IsEqualTo("A1:B3");
        await Assert.That(ws.Cell("B1").AsRange().Grow().RangeAddress.ToString()).IsEqualTo("A1:C2");

        await Assert.That(ws.Cell("F5").AsRange().Grow().RangeAddress.ToString()).IsEqualTo("E4:G6");
        await Assert.That(ws.Cell("F5").AsRange().Grow(2).RangeAddress.ToString()).IsEqualTo("D3:H7");
        await Assert.That(ws.Cell("F5").AsRange().Grow(100).RangeAddress.ToString()).IsEqualTo("A1:DB105");
    }

    [Test]
    public async Task ShrinkRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(ws.Cell("A1").AsRange().Shrink()).IsNull();
        await Assert.That(ws.Range("B2:C3").Shrink()).IsNull();
        await Assert.That(ws.Range("B2:D4").Shrink().RangeAddress.ToString()).IsEqualTo("C3:C3");
        await Assert.That(ws.Range("A1:Z26").Shrink(10).RangeAddress.ToString()).IsEqualTo("K11:P16");

        // Grow and shrink back
        await Assert.That(ws.Cell("Z26").AsRange().Grow(10).Shrink(10).RangeAddress.ToString()).IsEqualTo("Z26:Z26");
    }

    [Test]
    public async Task Intersection()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(ws.Range("B9:I11").Intersection(ws.Range("D4:G16")).ToString()).IsEqualTo("D9:G11");
        await Assert.That(ws.Range("E9:I11").Intersection(ws.Range("D4:G16")).ToString()).IsEqualTo("E9:G11");
        await Assert.That(ws.Cell("E9").AsRange().Intersection(ws.Range("D4:G16")).ToString()).IsEqualTo("E9:E9");
        await Assert.That(ws.Range("D4:G16").Intersection(ws.Cell("E9").AsRange()).ToString()).IsEqualTo("E9:E9");

        var rangeAddress = (XLRangeAddress)ws.Cell("C3").AsRange().Intersection(ws.Cell("A1").AsRange());
        await Assert.That(rangeAddress.IsValid).IsFalse();

        rangeAddress = (XLRangeAddress)ws.Cell("A1").AsRange().Intersection(ws.Cell("C3").AsRange());
        await Assert.That(rangeAddress.IsValid).IsFalse();

        await Assert.That(ws.Range("A1:C3").Intersection(null)).IsNull();

        var otherWs = wb.AddWorksheet("Sheet2");
        await Assert.That(ws.Intersection(otherWs)).IsNull();
        await Assert.That(ws.Cell("A1").AsRange().Intersection(otherWs.Cell("A2").AsRange())).IsNull();
    }

    [Test]
    public async Task Union()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(ws.Range("B9:I11").Union(ws.Range("D4:G16")).Count()).IsEqualTo(64);
        await Assert.That(ws.Range("E9:I11").Union(ws.Range("D4:G16")).Count()).IsEqualTo(58);
        await Assert.That(ws.Cell("E9").AsRange().Union(ws.Range("D4:G16")).Count()).IsEqualTo(52);
        await Assert.That(ws.Range("D4:G16").Union(ws.Cell("E9").AsRange()).Count()).IsEqualTo(52);

        await Assert.That(ws.Cell("A1").AsRange().Union(ws.Cell("C3").AsRange()).Count()).IsEqualTo(2);

        await Assert.That(ws.Range("A1:C3").Union(null).Count()).IsEqualTo(9);

        var otherWs = wb.AddWorksheet("Sheet2");
        await Assert.That(ws.Union(otherWs).Any()).IsFalse();
        await Assert.That(ws.Cell("A1").AsRange().Union(otherWs.Cell("A2").AsRange()).Any()).IsFalse();
    }

    [Test]
    public async Task Difference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(ws.Range("B9:I11").Difference(ws.Range("D4:G16")).Count()).IsEqualTo(12);
        await Assert.That(ws.Range("E9:I11").Difference(ws.Range("D4:G16")).Count()).IsEqualTo(6);
        await Assert.That(ws.Cell("E9").AsRange().Difference(ws.Range("D4:G16")).Count()).IsEqualTo(0);
        await Assert.That(ws.Range("D4:G16").Difference(ws.Cell("E9").AsRange()).Count()).IsEqualTo(51);

        await Assert.That(ws.Cell("A1").AsRange().Difference(ws.Cell("C3").AsRange()).Count()).IsEqualTo(1);

        await Assert.That(ws.Range("A1:C3").Difference(null).Count()).IsEqualTo(9);

        var otherWs = wb.AddWorksheet("Sheet2");
        await Assert.That(ws.Difference(otherWs).Any()).IsFalse();
        await Assert.That(ws.Cell("A1").AsRange().Difference(otherWs.Cell("A2").AsRange()).Any()).IsFalse();
    }

    [Test]
    public async Task SurroundingCells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        await Assert.That(ws.FirstCell().AsRange().SurroundingCells().Count()).IsEqualTo(3);
        await Assert.That(ws.Cell("C3").AsRange().SurroundingCells().Count()).IsEqualTo(8);
        await Assert.That(ws.Range("C3:D6").AsRange().SurroundingCells().Count()).IsEqualTo(16);

        await Assert.That(ws.Range("C3:D6").AsRange().SurroundingCells(c => !c.IsEmpty()).Count()).IsEqualTo(0);
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeAbove1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:D7").AddConditionalFormat();
        ws.Range("B2:E3").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative()).IsEqualTo("C4:D7");
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeAbove2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:D7").AddConditionalFormat();
        ws.Range("C3:D3").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative()).IsEqualTo("C4:D7");
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeBelow1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:D7").AddConditionalFormat();
        ws.Range("B7:E8").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative()).IsEqualTo("C3:D6");
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeBelow2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:D7").AddConditionalFormat();
        ws.Range("C7:D7").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative()).IsEqualTo("C3:D6");
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeRowInMiddle()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:D7").AddConditionalFormat();
        ws.Range("C5:E5").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative()).IsEqualTo("C3:D4");
        await Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative()).IsEqualTo("C6:D7");
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeColumnInMiddle()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:G4").AddConditionalFormat();
        ws.Range("E2:E4").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative()).IsEqualTo("C3:D4");
        await Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative()).IsEqualTo("F3:G4");
    }

    [Test]
    public async Task ClearConditionalFormattingsWhenRangeContainsFormatWhole()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:G4").AddConditionalFormat();
        ws.Range("B2:G4").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task NoClearConditionalFormattingsWhenRangePartiallySuperimposed()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:G4").AddConditionalFormat();
        ws.Range("C2:D3").Clear(XLClearOptions.ConditionalFormats);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Count).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToStringRelative()).IsEqualTo("C3:G4");
    }

    [Test]
    public async Task RangesRemoveAllWithoutDispose()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var ranges = new XLRanges
        {
            ws.Range("A1:A2"),
            ws.Range("B1:B2")
        };
        var rangesCopy = ranges.ToList();

        ranges.RemoveAll(null, false);
        ws.FirstColumn().InsertColumnsBefore(1);

        await Assert.That(ranges.Count).IsEqualTo(0);
        // if ranges were not disposed they addresses should change
        await Assert.That(rangesCopy.First().RangeAddress.ToString()).IsEqualTo("B1:B2");
        await Assert.That(rangesCopy.Last().RangeAddress.ToString()).IsEqualTo("C1:C2");
    }

    [Test]
    public async Task RangesRemoveAllByCriteria()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var ranges = new XLRanges
        {
            ws.Range("A1:A2"),
            ws.Range("B1:B3"),
            ws.Range("C1:C4")
        };
        var otherRange = ws.Range("A3:D3");

        ranges.RemoveAll(r => r.Intersects(otherRange));

        await Assert.That(ranges.Count).IsEqualTo(1);
        await Assert.That(ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A2");
    }

    [Test]
    public async Task XLRangesReturnsRangesInDeterministicOrder()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var ws2 = wb.Worksheets.Add("Another sheet");

        var ranges = new XLRanges
        {
            ws2.Range("F1:F12"),
            ws1.Range("F12:F16"),
            ws1.Range("B1:F2"),
            ws2.Range("A13:B14"),
            ws2.Range("E1:E2"),
            ws1.Range("E1:H2"),
            ws1.Range("G2:G13"),
            ws1.Range("G20:G20")
        };

        var expectedRanges = new List<IXLRange>
        {
            ws1.Range("B1:F2"),
            ws1.Range("E1:H2"),
            ws1.Range("G2:G13"),
            ws1.Range("F12:F16"),
            ws1.Range("G20:G20"),

            ws2.Range("E1:E2"),
            ws2.Range("F1:F12"),
            ws2.Range("A13:B14"),
        };

        var actualRanges = ranges.ToList();

        await Assert.That(actualRanges.Count).IsEqualTo(expectedRanges.Count);
        for (var i = 0; i < actualRanges.Count; i++)
        {
            await Assert.That(actualRanges[i]).IsEqualTo(expectedRanges[i]);
        }
    }

    [Test]
    public async Task ClearRangeRemovesSparklines()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.SparklineGroups.Add("B1:B3", "C1:E3");

        ws.Range("B1:C1").Clear();
        ws.Range("B2:C2").Clear(XLClearOptions.Sparklines);

        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").HasSparkline).IsFalse();
        await Assert.That(ws.Cell("B2").HasSparkline).IsFalse();
        await Assert.That(ws.Cell("B3").HasSparkline).IsTrue();
    }

    [Test]
    [Arguments("B2:G7", "D4:E5", true, "B2:G3,B4:C5,D4:E5,F4:G5,B6:G7")]
    [Arguments("B2:G7", "D4:E5", false, "B2:G3,B4:C5,F4:G5,B6:G7")]
    [Arguments("B2:G7", "B2:G7", true, "B2:G7")]
    [Arguments("B2:G7", "B2:G7", false, "")]
    [Arguments("B2:G7", "A1:H8", true, "B2:G7")]
    [Arguments("B2:G7", "A1:H8", false, "")]
    [Arguments("B2:G7", "A1:B2", true, "B2:B2,C2:G2,B3:G7")]
    [Arguments("B2:G7", "A1:B2", false, "C2:G2,B3:G7")]
    [Arguments("B2:G7", "E4:J5", true, "B2:G3,B4:D5,E4:G5,B6:G7")]
    [Arguments("B2:G7", "E4:J5", false, "B2:G3,B4:D5,B6:G7")]
    [Arguments("B2:G7", "A11:H18", true, "B2:G7")]
    [Arguments("B2:G7", "A11:H18", false, "B2:G7")]
    [Arguments("B2:G7", "A1:H1", true, "B2:G7")]
    [Arguments("B2:G7", "A1:A12", true, "B2:G7")]
    [Arguments("B2:G7", "A8:H8", true, "B2:G7")]
    [Arguments("B2:G7", "H1:H8", true, "B2:G7")]
    public async Task CanSplitRange(string rangeAddress, string splitBy, bool includeIntersection, string expectedResult)
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range(rangeAddress) as XLRange;
        var splitter = ws.Range(splitBy);

        var result = range.Split(splitter.RangeAddress, includeIntersection);

        var actualAddresses = string.Join(",", result.Select(r => r.RangeAddress.ToString()));

        await Assert.That(actualAddresses).IsEqualTo(expectedResult);
    }

    [Test]
    public async Task Sorting_moves_values_and_fixes_formula_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var range = ws.Cell("A1").InsertData(new object[]
        {
            ("Price", "Amount", "Sales"),
            (7, 5, Blank.Value),
            (2, 14, Blank.Value),
            (32, 2, Blank.Value),
            (6, 9, Blank.Value)
        });
        ws.Cell("C2").FormulaA1 = "A2*B2 & \"(Cake)\""; // 35
        ws.Cell("C3").FormulaA1 = "A3*B3 & \"(Pie)\""; // 28
        ws.Cell("C4").FormulaA1 = "A4*B4 & \"(Waffle)\""; // 64
        ws.Cell("C5").FormulaA1 = "A5*B5 & \"(Shortcake)\""; // 54

        // Sort uses cached values - update them
        ws.RecalculateAllFormulas();

        range.Sort("3 DESC");

        await Assert.That(ws.Cell("A2").Value).IsEqualTo(32);
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("A4").Value).IsEqualTo(7);
        await Assert.That(ws.Cell("A5").Value).IsEqualTo(2);

        await Assert.That(ws.Cell("B2").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(9);
        await Assert.That(ws.Cell("B4").Value).IsEqualTo(5);
        await Assert.That(ws.Cell("B5").Value).IsEqualTo(14);

        // Formulas has been moved around and their coordinates fixed after move
        await Assert.That(ws.Cell("C2").FormulaA1).IsEqualTo("A2*B2 & \"(Waffle)\"");
        await Assert.That(ws.Cell("C3").FormulaA1).IsEqualTo("A3*B3 & \"(Shortcake)\"");
        await Assert.That(ws.Cell("C4").FormulaA1).IsEqualTo("A4*B4 & \"(Cake)\"");
        await Assert.That(ws.Cell("C5").FormulaA1).IsEqualTo("A5*B5 & \"(Pie)\"");
    }

    [Test]
    [Arguments("PY(4)", "_xlfn._xlws.PY(4)")]
    [Arguments("2 + CHISQ.INV(0.6,2)", "2 + _xlfn.CHISQ.INV(0.6,2)")]
    [Arguments("2 + _xlfn.CHISQ.INV(0.6,2)", "2 + _xlfn.CHISQ.INV(0.6,2)")]
    public async Task FormulaArrayA1_adds_prefix_to_future_functions(string formula, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:B2").FormulaArrayA1 = formula;
        var masterCellFormula = ws.Cell("A1").FormulaA1;
        await Assert.That(masterCellFormula).IsEqualTo(expected);
    }
}
