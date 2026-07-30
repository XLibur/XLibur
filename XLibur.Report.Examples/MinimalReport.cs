using System.IO;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// The smallest template that does anything: a heading, one row repeated per item, and a defined name
/// tying the two together.
/// </summary>
/// <remarks>
/// There are only three things to learn here, and every other example is these three used harder.
/// <list type="number">
/// <item><c>{{ Company }}</c> in a cell is replaced by the variable of that name.</item>
/// <item>
/// A <em>defined name</em> over a block of rows, matching the name of a variable holding a collection,
/// makes that block repeat once per item. Inside it, <c>item</c> is the current one.
/// </item>
/// <item>
/// The block's <em>last</em> row is its options row: it carries tags and totals and is not repeated.
/// Empty, as here, it is removed. A one-row block has no options row and repeats entirely.
/// </item>
/// </list>
/// </remarks>
public sealed class MinimalReport : ReportExample
{
    public override string Name => "MinimalReport";

    public override string Summary => "A title, a repeated row, and the defined name that binds them.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Sales");

        sheet.Cell("A1").Value = "{{ Company }} — every line";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

        Headings(sheet, 3, "Product", "Quantity", "Unit price");

        // Row 4 is repeated once per sale. Row 5 is the options row: nothing in it, so it goes.
        sheet.Cell("A4").Value = "{{ item.Product }}";
        sheet.Cell("B4").Value = "{{ item.Quantity }}";
        sheet.Cell("C4").Value = "{{ item.UnitPrice }}";
        sheet.Cell("C4").Style.NumberFormat.Format = "#,##0.00";

        workbook.DefinedNames.Add("Sales", sheet.Range("A4:C5"));

        sheet.Columns(1, 3).Width = 16;
    }

    protected override void AddData(IXLTemplate template)
    {
        template.AddVariable("Company", "Contoso Ltd");
        template.AddVariable("Sales", SalesData.Sales());
    }

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("Sales");

        output.WriteLine($"  Title      {sheet.Cell("A1").GetFormattedString()}");
        output.WriteLine($"  Rows used  through {sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber}"
            + "   (the template's row 4, twelve times; row 5 was empty and went)");
        output.WriteLine();
    }
}
