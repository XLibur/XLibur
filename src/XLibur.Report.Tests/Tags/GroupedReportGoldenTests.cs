using System.Collections.Generic;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Tests.Infrastructure;

namespace XLibur.Report.Tests.Tags;

/// <summary>
/// The end-to-end grouping case: two nested group levels, a sort deciding the order inside the
/// innermost group, merged labels on the outer one, a subtotal per group and a grand total over
/// the lot.
/// </summary>
/// <remarks>
/// Written as a golden file rather than as assertions because grouping's output is as much shape as
/// value — outline levels, row order, where the totals land — and the comparer checks all of it at
/// once, including the things nobody thought to assert.
/// </remarks>
public class GroupedReportGoldenTests
{
    private static List<SaleItem> Items() => new()
    {
        new() { Region = "North", Category = "Tools", Product = "Widget", Quantity = 2 },
        new() { Region = "South", Category = "Toys", Product = "Gadget", Quantity = 5 },
        new() { Region = "North", Category = "Toys", Product = "Doohickey", Quantity = 1 },
        new() { Region = "South", Category = "Tools", Product = "Sprocket", Quantity = 3 },
        new() { Region = "North", Category = "Tools", Product = "Anvil", Quantity = 4 },
        new() { Region = "South", Category = "Toys", Product = "Balloon", Quantity = 6 },
    };

    private static void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Sales");

        sheet.Cell("A1").Value = "Sales by region";
        sheet.Cell("A1").Style.Font.Bold = true;

        sheet.Cell("A2").Value = "Region";
        sheet.Cell("B2").Value = "Category";
        sheet.Cell("C2").Value = "Product";
        sheet.Cell("D2").Value = "Qty";
        sheet.Range("A2:D2").Style.Font.Italic = true;

        sheet.Cell("A3").Value = "{{ item.Region }}";
        sheet.Cell("B3").Value = "{{ item.Category }}";
        sheet.Cell("C3").Value = "{{ item.Product }}";
        sheet.Cell("D3").Value = "{{ item.Quantity }}";

        // Row 4 is the options row. Its styling is what every subtotal row takes.
        sheet.Cell("A4").Value = "<<Group merge>>";
        sheet.Cell("B4").Value = "<<Group>>";
        sheet.Cell("C4").Value = "<<Sort>>";
        sheet.Cell("D4").Value = "<<Sum>>";
        sheet.Range("A4:D4").Style.Font.Bold = true;

        sheet.Cell("A5").Value = "End of report";

        workbook.DefinedNames.Add("Items", sheet.Range("A3:D4"));
    }

    private static void Bind(IXLTemplate template) => template.AddVariable("Items", Items());

    private static void BuildExpected(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Sales");

        sheet.Cell("A1").Value = "Sales by region";
        sheet.Cell("A1").Style.Font.Bold = true;

        sheet.Cell("A2").Value = "Region";
        sheet.Cell("B2").Value = "Category";
        sheet.Cell("C2").Value = "Product";
        sheet.Cell("D2").Value = "Qty";
        sheet.Range("A2:D2").Style.Font.Italic = true;

        // Sorted by product, then ordered into groups: the sort survives inside each group.
        Data(sheet, 3, "North", "Tools", "Anvil", 4);
        Data(sheet, 4, "North", "Tools", "Widget", 2);
        Subtotal(sheet, 5, "B", "Tools Total", "D3:D4", outlineLevel: 1);
        Data(sheet, 6, "North", "Toys", "Doohickey", 1);
        Subtotal(sheet, 7, "B", "Toys Total", "D6:D6", outlineLevel: 1);
        Subtotal(sheet, 8, "A", "North Total", "D3:D7", outlineLevel: 0);

        Data(sheet, 9, "South", "Tools", "Sprocket", 3);
        Subtotal(sheet, 10, "B", "Tools Total", "D9:D9", outlineLevel: 1);
        Data(sheet, 11, "South", "Toys", "Balloon", 6);
        Data(sheet, 12, "South", "Toys", "Gadget", 5);
        Subtotal(sheet, 13, "B", "Toys Total", "D11:D12", outlineLevel: 1);
        Subtotal(sheet, 14, "A", "South Total", "D9:D13", outlineLevel: 0);

        // The options row survives because the grand total was written into it. SUBTOTAL spans the
        // group totals and ignores them, so the report total is not the data counted twice.
        sheet.Cell("D15").FormulaA1 = "SUBTOTAL(9,D3:D14)";
        sheet.Range("A15:D15").Style.Font.Bold = true;

        sheet.Cell("A16").Value = "End of report";

        // A region's label is written once and merged over its rows, its own total excluded.
        sheet.Range("A3:A7").Merge();
        sheet.Range("A3:A7").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Range("A9:A13").Merge();
        sheet.Range("A9:A13").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        workbook.DefinedNames.Add("Items", sheet.Range("A3:D14"));
    }

    private static void Data(IXLWorksheet sheet, int row, string region, string category, string product, double quantity)
    {
        sheet.Cell(row, 1).Value = region;
        sheet.Cell(row, 2).Value = category;
        sheet.Cell(row, 3).Value = product;
        sheet.Cell(row, 4).Value = quantity;
        sheet.Row(row).OutlineLevel = 2;
    }

    private static void Subtotal(IXLWorksheet sheet, int row, string labelColumn, string label, string totalled, int outlineLevel)
    {
        sheet.Cell(labelColumn + row).Value = label;
        sheet.Cell(row, 4).FormulaA1 = $"SUBTOTAL(9,{totalled})";
        sheet.Range(row, 1, row, 4).Style.Font.Bold = true;
        sheet.Row(row).OutlineLevel = outlineLevel;
    }

    private static ReportFixture Fixture() => new("GroupedReport", BuildTemplate, Bind, BuildExpected);

    [Test]
    public async Task GroupedReportMatchesItsExpectation()
    {
        await Assert.That(() => GoldenFile.Verify(Fixture())).ThrowsNothing();
    }
}
