using System.IO;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// Excel's own functions, called inside <c>{{ }}</c>.
/// </summary>
/// <remarks>
/// <para>
/// Every function XLibur's calculation engine implements is available in a template expression under
/// its upper-case Excel name. So a report author who knows Excel does not have to learn a second set
/// of names for the same things: <c>ROUND</c> is <c>ROUND</c>, and <c>IF</c> is <c>IF</c> even though
/// <c>if</c> is a keyword of the expression language.
/// </para>
/// <para>
/// The results are <em>typed</em>. <c>{{ SUM(...) }}</c> over decimals lands in a number cell that can
/// be formatted and totalled, not as text that looks like a number — which is the whole reason the
/// bridge exists rather than the engine formatting values itself.
/// </para>
/// <para>
/// This is different from writing a formula into the cell. A formula is evaluated by Excel every time
/// the file is opened; these are evaluated once, while the report is generated, and the workbook holds
/// the answer. Use whichever the report wants: <see cref="AnnualSalesReport"/> writes a real formula
/// for its line totals, because a reader who edits a quantity should see the total change.
/// </para>
/// </remarks>
public sealed class ExcelFunctionsInExpressions : ReportExample
{
    public override string Name => "ExcelFunctions";

    public override string Summary => "SUM, IF, ROUND and the rest of Excel's functions inside {{ }}.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Functions");

        sheet.Cell("A1").Value = "Excel functions in template expressions";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

        Summaries(sheet);
        PerRow(sheet);

        sheet.Column(1).Width = 34;
        sheet.Columns(2, 4).Width = 16;
    }

    /// <summary>
    /// Functions over the whole collection, outside any bound range. These are evaluated once, before
    /// anything is generated.
    /// </summary>
    private static void Summaries(IXLWorksheet sheet)
    {
        sheet.Cell("A3").Value = "Lines";
        sheet.Cell("B3").Value = "{{ COUNT(array.map Sales \"Quantity\") }}";

        sheet.Cell("A4").Value = "Total sold";
        sheet.Cell("B4").Value = "{{ SUM(array.map Sales \"Total\") }}";
        sheet.Cell("B4").Style.NumberFormat.Format = "#,##0.00";

        sheet.Cell("A5").Value = "Largest line";
        sheet.Cell("B5").Value = "{{ MAX(array.map Sales \"Total\") }}";
        sheet.Cell("B5").Style.NumberFormat.Format = "#,##0.00";

        sheet.Cell("A6").Value = "Average line, to the penny";
        sheet.Cell("B6").Value = "{{ ROUND(AVERAGE(array.map Sales \"Total\"), 2) }}";
        sheet.Cell("B6").Style.NumberFormat.Format = "#,##0.00";

        sheet.Cell("A7").Value = "Reads as a number, not as text";
        sheet.Cell("B7").FormulaA1 = "B4/B3";
        sheet.Cell("B7").Style.NumberFormat.Format = "#,##0.00";

        sheet.Range("A3:A7").Style.Font.SetItalic();

        sheet.Cell("A9").Value =
            "array.map is the expression language's, not Excel's: it pulls one property out of every "
            + "item so an Excel function has a list to work on.";
        sheet.Cell("A9").Style.Font.SetFontColor(XLColor.Gray);
    }

    /// <summary>Functions inside a bound range, per item.</summary>
    private static void PerRow(IXLWorksheet sheet)
    {
        Headings(sheet, 11, "Product", "Line total", "Size", "Rounded up");

        sheet.Cell("A12").Value = "{{ item.Product }}";
        sheet.Cell("B12").Value = "{{ item.Total }}";
        sheet.Cell("B12").Style.NumberFormat.Format = "#,##0.00";

        // IF works despite `if` being a keyword of the expression language: the bridge registers the
        // upper-case name, and the parser reads an upper-case identifier followed by ( as a call.
        sheet.Cell("C12").Value = "{{ IF(item.Total > 500, \"large\", IF(item.Total > 100, \"medium\", \"small\")) }}";

        sheet.Cell("D12").Value = "{{ CEILING(item.Total, 10) }}";
        sheet.Cell("D12").Style.NumberFormat.Format = "#,##0";

        sheet.Range("A13:D13").Style.Font.SetBold();
        sheet.Range("A13:D13").Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        sheet.Cell("B13").Value = "<<Sum>>";
        sheet.Cell("B13").Style.NumberFormat.Format = "#,##0.00";

        sheet.Range("A12:D13").Worksheet.Workbook.DefinedNames.Add(
            "Sales",
            sheet.Range("A12:D13"));
    }

    protected override void AddData(IXLTemplate template)
    {
        // One variable, read two ways: the defined name Sales repeats its items, and the expressions
        // outside the range read the same list as a whole.
        template.AddVariable("Sales", SalesData.Sales());
    }

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("Functions");

        // The data type is the thing to look at. A number cell can be formatted and divided; text that
        // looks like a number cannot, and B7 dividing B4 by B3 is what proves the difference.
        output.WriteLine($"  COUNT     {sheet.Cell("B3").Value}   ({sheet.Cell("B3").Value.Type})");
        output.WriteLine($"  SUM       {sheet.Cell("B4").GetFormattedString()}   ({sheet.Cell("B4").Value.Type})");
        output.WriteLine($"  MAX       {sheet.Cell("B5").GetFormattedString()}   ({sheet.Cell("B5").Value.Type})");
        output.WriteLine($"  ROUND     {sheet.Cell("B6").GetFormattedString()}   ({sheet.Cell("B6").Value.Type})");
        output.WriteLine($"  B4/B3     {sheet.Cell("B7").GetFormattedString()}   (a formula over them, so they are numbers)");
        output.WriteLine($"  IF, row 1 {sheet.Cell("C12").GetFormattedString()}");
        output.WriteLine();
    }
}
