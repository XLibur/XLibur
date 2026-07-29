using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Tests.Infrastructure;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// The first end-to-end golden-file case: a header bound to a workbook variable, a repeating
/// block bound to a collection, a per-row formula, and a footer that has to end up below the
/// generated rows.
/// </summary>
public class SalesReportGoldenTests
{
    private static List<SaleItem> Items() => new()
    {
        new() { Product = "Widget", Quantity = 2, UnitPrice = 5m, SoldOn = new DateTime(2026, 1, 5) },
        new() { Product = "Gadget", Quantity = 3, UnitPrice = 10m, SoldOn = new DateTime(2026, 1, 6) },
        new() { Product = "Doohickey", Quantity = 1, UnitPrice = 7.5m, SoldOn = new DateTime(2026, 1, 7) },
    };

    private static void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Sales");

        sheet.Cell("A1").Value = "Sales for {{ Company }}";
        sheet.Cell("A1").Style.Font.Bold = true;

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("B2").Value = "Qty";
        sheet.Cell("C2").Value = "Unit price";
        sheet.Cell("D2").Value = "Total";
        sheet.Range("A2:D2").Style.Font.Italic = true;

        sheet.Cell("A3").Value = "{{ item.Product }}";
        sheet.Cell("B3").Value = "{{ item.Quantity }}";
        sheet.Cell("C3").Value = "{{ item.UnitPrice }}";
        sheet.Cell("D3").FormulaA1 = "B3*C3";

        // Row 4 is the options row: it carries no content here, so generation removes it.
        sheet.Cell("A5").Value = "End of report";

        workbook.DefinedNames.Add("Items", sheet.Range("A3:D4"));
    }

    private static void Bind(IXLTemplate template)
    {
        template.AddVariable("Company", "Contoso");
        template.AddVariable("Items", Items());
    }

    private static void BuildExpected(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Sales");

        sheet.Cell("A1").Value = "Sales for Contoso";
        sheet.Cell("A1").Style.Font.Bold = true;

        sheet.Cell("A2").Value = "Product";
        sheet.Cell("B2").Value = "Qty";
        sheet.Cell("C2").Value = "Unit price";
        sheet.Cell("D2").Value = "Total";
        sheet.Range("A2:D2").Style.Font.Italic = true;

        var rows = new (string Product, double Quantity, double UnitPrice)[]
        {
            ("Widget", 2, 5),
            ("Gadget", 3, 10),
            ("Doohickey", 1, 7.5),
        };

        for (var i = 0; i < rows.Length; i++)
        {
            var row = 3 + i;
            sheet.Cell(row, 1).Value = rows[i].Product;
            sheet.Cell(row, 2).Value = rows[i].Quantity;
            sheet.Cell(row, 3).Value = rows[i].UnitPrice;
            sheet.Cell(row, 4).FormulaA1 = $"B{row}*C{row}";
        }

        sheet.Cell("A6").Value = "End of report";

        // Generation re-points the name at the rows it produced.
        workbook.DefinedNames.Add("Items", sheet.Range("A3:D5"));
    }

    private static ReportFixture Fixture() => new("SalesReport", BuildTemplate, Bind, BuildExpected);

    [Test]
    public async Task SalesReportMatchesItsExpectation()
    {
        await Assert.That(() => GoldenFile.Verify(Fixture())).ThrowsNothing();
    }
}
