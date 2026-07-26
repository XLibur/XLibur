using System.IO;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PivotTables.Style;

internal class XLPivotTableStyleFormatsTests
{
    [Test]
    public async Task Add_grand_row_total_styles()
    {
        await TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Price"),
                ("Cake", 9),
                ("Pie", 7),
                ("Cake", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            ptSheet.Column("A").Width = 15;
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.RowLabels.Add("Name");
            pt.Values.Add("Price", "Avg $").SetSummaryFormula(XLPivotSummary.Average);
            pt.Values.Add("Price", "Max $").SetSummaryFormula(XLPivotSummary.Maximum);

            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.All).Style
                .Font.SetFontSize(15)
                .Font.SetUnderline(XLFontUnderlineValues.Double);
            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.Label).Style
                .Font.SetFontColor(XLColor.Green);
            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.Data).Style
                .Font.SetFontColor(XLColor.Red);
        }, @"Other\PivotTable\Style\Add_grand_row_total_styles.xlsx");
    }

    [Test]
    public async Task Alignment_in_pivot_format_survives_round_trip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Price"),
                ("Cake", 9),
                ("Pie", 7),
                ("Cake", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.RowLabels.Add("Name");
            pt.Values.Add("Price");

            // Set various alignment properties on the grand total format
            pt.StyleFormats.RowGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.All).Style
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Alignment.SetVertical(XLAlignmentVerticalValues.Top)
                .Alignment.SetWrapText(true)
                .Alignment.SetTextRotation(45);

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var pt = (XLPivotTable)wb.Worksheet(2).PivotTables.PivotTable("pivot table");

            // Check DxfStyleValue on the loaded format — the internal representation
            // that proves alignment round-tripped through the DXF record.
            var format = pt.Formats[0];
            var alignment = format.DxfStyleValue.Alignment;
            await Assert.That(alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
            await Assert.That(alignment.Vertical).IsEqualTo(XLAlignmentVerticalValues.Top);
            await Assert.That(alignment.WrapText).IsEqualTo(true);
            await Assert.That(alignment.TextRotation).IsEqualTo(45);

            // Non-set properties remain at defaults
            await Assert.That(alignment.Indent).IsEqualTo(0);
            await Assert.That(alignment.ShrinkToFit).IsEqualTo(false);
            await Assert.That(alignment.ReadingOrder).IsEqualTo(XLAlignmentReadingOrderValues.ContextDependent);
        }
    }

    [Test]
    public async Task Add_grand_column_total_styles()
    {
        await TestHelper.CreateAndCompare(wb =>
        {
            var dataSheet = wb.AddWorksheet();
            var dataRange = dataSheet.Cell("A1").InsertData(new object[]
            {
                ("Name", "Month", "Price"),
                ("Cake", "Jan", 9),
                ("Pie", "Jan", 7),
                ("Cake", "Feb", 3),
            });

            var ptSheet = wb.AddWorksheet().SetTabActive();
            ptSheet.Column("A").Width = 15;
            var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");
            pt.RowLabels.Add("Name");
            pt.RowLabels.Add("Month");
            pt.Values.Add("Price");

            pt
                .SetShowGrandTotalsColumns(true)
                .SetShowGrandTotalsRows(false);

            pt.StyleFormats.ColumnGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.All).Style
                .Font.SetFontSize(15)
                .Font.SetUnderline(XLFontUnderlineValues.Double);
            pt.StyleFormats.ColumnGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.Label).Style
                .Font.SetFontColor(XLColor.Green);
            pt.StyleFormats.ColumnGrandTotalFormats
                .ForElement(XLPivotStyleFormatElement.Data).Style
                .Font.SetFontColor(XLColor.Red);
        }, @"Other\PivotTable\Style\Add_grand_column_total_styles.xlsx");
    }
}
