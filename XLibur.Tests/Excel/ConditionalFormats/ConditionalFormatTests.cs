using XLibur.Excel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using XLibur.Extensions;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.ConditionalFormats;

public class ConditionalFormatTests
{
    [Test]
    public async Task MaintainConditionalFormattingOrder()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ConditionalFormattingOrder\inputfile.xlsx"));
        using var ms = new MemoryStream();
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);
            wb.SaveAs(ms);
            return wb;
        }, @"Other\StyleReferenceFiles\ConditionalFormattingOrder\ConditionalFormattingOrder.xlsx");
    }


    [Test]
    [Arguments(true, 7)]
    [Arguments(false, 8)]
    public async Task SaveOptionAffectsConsolidationConditionalFormatRanges(bool consolidateConditionalFormatRanges, int expectedCount)
    {
        var options = new SaveOptions
        {
            ConsolidateConditionalFormatRanges = consolidateConditionalFormatRanges
        };

        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");

        ws.Range("D2:D3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("B2:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("E2:E6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);
        ws.Range("F2:F6").AddConditionalFormat().ColorScale().LowestValue(XLColor.Red).HighestValue(XLColor.Blue);
        ws.Range("G2:G7").AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XLColor.Blue);
        ws.Range("H2:H7").AddConditionalFormat().WhenIsUnique().Fill.SetBackgroundColor(XLColor.Blue);
        ws.Range("I2:I6").AddConditionalFormat().WhenContains("test");
        ws.Range("J2:J6").AddConditionalFormat().WhenContains("test");
        using var ms = new MemoryStream();
        wb.SaveAs(ms, options);
        var wb_saved = new XLWorkbook(ms);
        await Assert.That(wb_saved.Worksheet("Sheet").ConditionalFormats.Count()).IsEqualTo(expectedCount);
    }

    [Test]
    [Arguments(true, 1)]
    [Arguments(false, 2)]
    public async Task SaveOptionAffectsConsolidationDataValidationRanges(bool consolidateDataValidationRanges, int expectedCount)
    {
        var options = new SaveOptions
        {
            ConsolidateDataValidationRanges = consolidateDataValidationRanges
        };

        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Range("C2:C5").CreateDataValidation().Decimal.Between(1, 5);
        ws.Range("D2:D5").CreateDataValidation().Decimal.Between(1, 5);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, options);
        var wb_saved = new XLWorkbook(ms);
        await Assert.That(wb_saved.Worksheet("Sheet").DataValidations.Count()).IsEqualTo(expectedCount);
    }

    [Test]
    [Arguments("en-US")]
    [Arguments("fr-FR")]
    [Arguments("ru-RU")]
    public async Task SaveConditionalFormat_CultureIndependent(string culture)
    {
        using var ms = new MemoryStream();
        var expectedValue = 1.5;
        Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            var i = 1;
            ws.Cell(i++, 1).AddConditionalFormat().WhenEquals(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(i++, 1).AddConditionalFormat().WhenNotEquals(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(i++, 1).AddConditionalFormat().WhenGreaterThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(i++, 1).AddConditionalFormat().WhenLessThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(i++, 1).AddConditionalFormat().WhenEqualOrGreaterThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(i++, 1).AddConditionalFormat().WhenEqualOrLessThan(expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            ws.Cell(i++, 1).AddConditionalFormat().WhenBetween(expectedValue, expectedValue).Fill.SetBackgroundColor(XLColor.Red);
            // ReSharper disable once RedundantAssignment
            ws.Cell(i++, 1).AddConditionalFormat().WhenNotBetween(expectedValue, expectedValue).Fill.SetBackgroundColor(XLColor.Red);

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();

            var conditionalFormatValues = ws.ConditionalFormats
                .SelectMany(cf => cf.Values.Values)
                .Select(v => v.Value)
                .Distinct();

            await Assert.That(conditionalFormatValues.Count()).IsEqualTo(1);
            await Assert.That(conditionalFormatValues.Single()).IsEqualTo("1.5");
        }
    }

    [Test]
    public async Task CellIs_type_reads_only_required_formula_arguments()
    {
        // The CellIs uses formula tags as arguments. Some producers generate extra empty
        // formula tags and XLibur should be able to load CellIs conditional formatting
        // with such extra tags without an exception. The test file has been modified to
        // include extra formula tags and test checks that extra tags are ignored.
        await TestHelper.LoadAndAssert(async (_, ws) =>
        {
            await AssertFormulaArgs(ws, XLCFOperator.Between, "$D$2", "$E$2");
            await AssertFormulaArgs(ws, XLCFOperator.NotBetween, "$D$3", "$E$3");
            await AssertFormulaArgs(ws, XLCFOperator.GreaterThan, "$D$4");
            await AssertFormulaArgs(ws, XLCFOperator.LessThan, "$D$5");
            await AssertFormulaArgs(ws, XLCFOperator.Equal, "$D$6");
        }, @"Other\ConditionalFormats\Extra_formulas_CellIs_type.xlsx");

        static async Task AssertFormulaArgs(IXLWorksheet ws, XLCFOperator cfOperator, params string[] expectedFormulas)
        {
            var cf = ws.ConditionalFormats.Single(cf => cf.ConditionalFormatType == XLConditionalFormatType.CellIs && cf.Operator == cfOperator);
            await Assert.That(cf.Values.Count).IsEqualTo(expectedFormulas.Length);
            await Assert.That(cf.Values.Select(v => v.Value.Value)).IsEquivalentTo(expectedFormulas, CollectionOrdering.Matching);
        }
    }

    [Test]
    public async Task DataBar_Gradient_RoundTrips()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;
        ws.Cell("A3").Value = 30;

        ws.Range("A1:A3").AddConditionalFormat()
            .DataBar(XLColor.FromArgb(0xFF638EC6), showBarOnly: false, gradient: true)
            .LowestValue()
            .HighestValue();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf = wb2.Worksheet("Sheet1").ConditionalFormats.Single();

        await Assert.That(cf.Gradient).IsTrue();
        await Assert.That(cf.Colors[1].Color.ToHex()).IsEqualTo("FF638EC6");
    }

    [Test]
    public async Task DataBar_SolidFill_RoundTrips()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;
        ws.Cell("A3").Value = 30;

        ws.Range("A1:A3").AddConditionalFormat()
            .DataBar(XLColor.FromArgb(0xFF638EC6), showBarOnly: false, gradient: false)
            .LowestValue()
            .HighestValue();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf = wb2.Worksheet("Sheet1").ConditionalFormats.Single();

        await Assert.That(cf.Gradient).IsFalse();
        await Assert.That(cf.Colors[1].Color.ToHex()).IsEqualTo("FF638EC6");
    }

    [Test]
    public async Task Expression_type_skips_empty_formula_tags()
    {
        // The Expression uses formula tag as arguments. Some producers generate extra empty
        // formula tags and XLibur should be able to load Expression conditional formatting
        // with such extra tags without an exception. The test file has been modified to
        // include extra formula tags and test checks that extra tags are ignored.
        await TestHelper.LoadAndAssert(async (_, ws) =>
        {
            await AssertFormulaArgs(ws, "A1:A1", "$C$1=5");
            await AssertFormulaArgs(ws, "A2:A2", "$C$2=4");
        }, @"Other\ConditionalFormats\Extra_formulas_Expression_type.xlsx");

        static async Task AssertFormulaArgs(IXLWorksheet ws, string range, string expectedFormula)
        {
            var cf = ws.ConditionalFormats.Single(cf => cf.ConditionalFormatType == XLConditionalFormatType.Expression && cf.Range.RangeAddress.ToString() == range);
            await Assert.That(cf.Values.Count).IsEqualTo(1);
            await Assert.That(cf.Values[1].Value).IsEqualTo(expectedFormula);
        }
    }

    [Test]
    public async Task ContainsText_with_pipe_in_value_can_be_loaded()
    {
        // Issue #2754: conditional format with pipe character in cell value
        // should load without parsing errors.
        await TestHelper.LoadAndAssert(async (_, ws) =>
        {
            var cf = ws.ConditionalFormats
                .Single(cf => cf.ConditionalFormatType == XLConditionalFormatType.ContainsText);
            await Assert.That(cf.Values[1].Value).IsEqualTo("70|");
        }, @"Other\ConditionalFormats\ConditionalFormat_cellvalueequal_2754.xlsx");
    }

    [Test]
    public async Task ContainsText_with_pipe_in_value_round_trips()
    {
        // Issue #2754: conditional format with pipe character in value
        // should survive save and reload.
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\ConditionalFormats\ConditionalFormat_cellvalueequal_2754.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var ws = wb2.Worksheets.First();
        var cf = ws.ConditionalFormats
            .Single(cf => cf.ConditionalFormatType == XLConditionalFormatType.ContainsText);
        await Assert.That(cf.Values[1].Value).IsEqualTo("70|");
    }

    [Test]
    public async Task DataBar_FluentChain_Returns_ConditionalFormat()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;

        var cf = ws.Range("A1:A2").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();

        await Assert.That(cf).IsNotNull();
        await Assert.That(cf.ConditionalFormatType).IsEqualTo(XLConditionalFormatType.DataBar);
        await Assert.That(cf.Colors[1]).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task DataBar_Maximum_Returns_ConditionalFormat()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;

        var cf = ws.Range("A1:A1").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .Minimum(XLCFContentType.Number, 0)
            .Maximum(XLCFContentType.Number, 100);

        await Assert.That(cf).IsNotNull();
        await Assert.That(cf.ConditionalFormatType).IsEqualTo(XLConditionalFormatType.DataBar);
    }

    [Test]
    public async Task DataBar_Color_Can_Be_Changed_And_RoundTrips()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;

        var cf = ws.Range("A1:A2").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();

        cf.Colors[1] = XLColor.Blue;
        await Assert.That(cf.Colors[1]).IsEqualTo(XLColor.Blue);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf2 = wb2.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf2.Colors[1].Color.ToHex()).IsEqualTo(XLColor.Blue.Color.ToHex());
    }

    [Test]
    public async Task DataBar_ShowBarOnly_And_Gradient_Can_Be_Toggled_And_RoundTrip()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;

        var cf = ws.Range("A1:A2").AddConditionalFormat()
            .DataBar(XLColor.Red, showBarOnly: false, gradient: true)
            .LowestValue()
            .HighestValue();

        await Assert.That(cf.ShowBarOnly).IsFalse();
        await Assert.That(cf.Gradient).IsTrue();

        cf.ShowBarOnly = true;
        cf.Gradient = false;

        await Assert.That(cf.ShowBarOnly).IsTrue();
        await Assert.That(cf.Gradient).IsFalse();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf2 = wb2.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf2.ShowBarOnly).IsTrue();
        await Assert.That(cf2.Gradient).IsFalse();
    }

    [Test]
    public async Task DataBar_Gradient_Changed_To_Flat_RoundTrips()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;
        ws.Cell("A3").Value = 30;

        var cf = ws.Range("A1:A3").AddConditionalFormat()
            .DataBar(XLColor.FromArgb(0xFF638EC6), showBarOnly: false, gradient: true)
            .LowestValue()
            .HighestValue();

        await Assert.That(cf.Gradient).IsTrue();

        // Switch from gradient to flat fill
        cf.Gradient = false;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf2 = wb2.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf2.Gradient).IsFalse();
        await Assert.That(cf2.Colors[1].Color.ToHex()).IsEqualTo("FF638EC6");
    }

    [Test]
    public async Task DataBar_AxisPosition_Defaults_To_Automatic()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;

        var cf = ws.Range("A1:A1").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();

        await Assert.That(cf.BarAxisPosition).IsEqualTo(XLDataBarAxisPosition.Automatic);
        await Assert.That(cf.BarAxisColor).IsEqualTo(XLColor.Black);
    }

    [Test]
    [Arguments(XLDataBarAxisPosition.Automatic)]
    [Arguments(XLDataBarAxisPosition.Middle)]
    [Arguments(XLDataBarAxisPosition.None)]
    public async Task DataBar_AxisPosition_RoundTrips(XLDataBarAxisPosition position)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = -10;
        ws.Cell("A2").Value = 20;

        var cf = ws.Range("A1:A2").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();

        cf.BarAxisPosition = position;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf2 = wb2.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf2.BarAxisPosition).IsEqualTo(position);
    }

    [Test]
    public async Task DataBar_AxisColor_RoundTrips()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = -10;
        ws.Cell("A2").Value = 20;

        var cf = ws.Range("A1:A2").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();

        cf.BarAxisColor = XLColor.DarkBlue;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var wb2 = new XLWorkbook(ms);
        var cf2 = wb2.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf2.BarAxisColor.Color.ToHex()).IsEqualTo(XLColor.DarkBlue.Color.ToHex());
    }

    [Test]
    public async Task DataBar_Can_Be_Removed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;

        ws.Range("A1:A2").AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);

        ws.ConditionalFormats.Remove(cf => cf.ConditionalFormatType == XLConditionalFormatType.DataBar);

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(0);
    }
}
