using System.Collections.Generic;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Tests.Infrastructure;

namespace XLibur.Report.Tests.Tags;

/// <summary>
/// The end-to-end conditional case: one range keeping only the rows that answer yes, and a second
/// dropped whole because the report was not asked for it.
/// </summary>
/// <remarks>
/// Written as a golden file because what <c>&lt;&lt;If&gt;&gt;</c> leaves behind is a shape as much
/// as a set of values — where the total lands, what happened to the headings of a range that
/// produced nothing, whether the text below it all moved by the right amount.
/// </remarks>
public class ConditionalReportGoldenTests
{
    private static List<SaleItem> Items() => new()
    {
        new() { Product = "Widget", Quantity = 2 },
        new() { Product = "Gadget", Quantity = 0 },
        new() { Product = "Doohickey", Quantity = 5 },
        new() { Product = "Sprocket", Quantity = 0 },
    };

    private static void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Orders");

        sheet.Cell("A1").Value = "Orders";
        sheet.Cell("A1").Style.Font.Bold = true;

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("B2").Value = "Qty";

        sheet.Cell("A3").Value = "{{ item.Product }}";
        sheet.Cell("B3").Value = "{{ item.Quantity }}";
        sheet.Cell("C3").Value = "<<If test=\"item.Quantity > 0\">>";

        sheet.Cell("A4").Value = "Total";
        sheet.Cell("B4").Value = "<<Sum>>";

        workbook.DefinedNames.Add("Ordered", sheet.Range("A3:C4"));

        sheet.Cell("A6").Value = "Backorders";
        sheet.Cell("A6").Style.Font.Italic = true;
        sheet.Cell("A7").Value = "{{ item.Product }}";
        sheet.Cell("C8").Value = "<<If test=\"ShowBackorders\">>";

        workbook.DefinedNames.Add("Backordered", sheet.Range("A7:C8"));

        sheet.Cell("A10").Value = "End of report";
    }

    private static void Bind(IXLTemplate template)
    {
        template.AddVariable("Ordered", Items());
        template.AddVariable("Backordered", Items());
        template.AddVariable("ShowBackorders", false);
    }

    private static void BuildExpected(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Orders");

        sheet.Cell("A1").Value = "Orders";
        sheet.Cell("A1").Style.Font.Bold = true;

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("B2").Value = "Qty";

        // The two rows whose quantity beat the test, in the order they arrived.
        sheet.Cell("A3").Value = "Widget";
        sheet.Cell("B3").Value = 2;
        sheet.Cell("A4").Value = "Doohickey";
        sheet.Cell("B4").Value = 5;

        // The options row survives because the total was written into it.
        sheet.Cell("A5").Value = "Total";
        sheet.Cell("B5").FormulaA1 = "SUBTOTAL(9,B3:B4)";

        // The backorders range answered no, so its rows and its options row are gone and its name
        // with them — exactly what an empty collection leaves. Its heading stays: a heading is not
        // part of the range.
        sheet.Cell("A7").Value = "Backorders";
        sheet.Cell("A7").Style.Font.Italic = true;

        sheet.Cell("A9").Value = "End of report";

        workbook.DefinedNames.Add("Ordered", sheet.Range("A3:C4"));
    }

    private static ReportFixture Fixture() => new("ConditionalReport", BuildTemplate, Bind, BuildExpected);

    [Test]
    public async Task ConditionalReportMatchesItsExpectation()
    {
        await Assert.That(() => GoldenFile.Verify(Fixture())).ThrowsNothing();
    }
}
