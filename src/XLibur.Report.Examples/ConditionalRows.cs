using System.IO;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// Leaving rows out, and leaving a whole range out, with <c>&lt;&lt;If&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// <para>
/// The same tag means two things depending on where it is written, which is the only sensible reading
/// of each:
/// </para>
/// <list type="bullet">
/// <item>
/// in a <b>repeated</b> row it is about one item, so a falsy test drops that row — a filter written in
/// the template instead of in the code that supplies the data;
/// </item>
/// <item>
/// in the <b>options</b> row it is about the range, so a falsy test drops the lot, options row
/// included — an optional section.
/// </item>
/// </list>
/// <para>
/// Only <c>null</c> and <c>false</c> are false. Zero and the empty string are <b>true</b>, so a test
/// meaning "more than nothing" has to say so: <c>test="item.Quantity &gt; 0"</c>, not
/// <c>test="item.Quantity"</c>. This follows the expression language rather than inventing a third set
/// of truthiness rules, and is worth knowing before it surprises anyone.
/// </para>
/// </remarks>
public sealed class ConditionalRows : ReportExample
{
    public override string Name => "ConditionalRows";

    public override string Summary => "<<If>> drops a row in a repeated row, and a whole range in an options row.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Conditional");

        sheet.Cell("A1").Value = "Export lines only";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);
        sheet.Cell("A2").Value = "Every line is bound; the ones that are not exports are dropped as rows.";
        sheet.Cell("A2").Style.Font.SetItalic().Font.SetFontColor(XLColor.Gray);

        ExportsOnly(sheet);
        OptionalSection(sheet);

        sheet.Column(1).Width = 18;
        sheet.Columns(2, 3).Width = 14;
    }

    /// <summary>
    /// A per-item filter. The test sits in the repeated row, so it is asked once per item, with that
    /// item in scope.
    /// </summary>
    private static void ExportsOnly(IXLWorksheet sheet)
    {
        Headings(sheet, 4, "Product", "Region", "Line total");

        sheet.Cell("A5").Value = "{{ item.Product }}";
        sheet.Cell("B5").Value = "{{ item.Region }}";
        sheet.Cell("C5").Value = "{{ item.Total }}";
        sheet.Cell("C5").Style.NumberFormat.Format = "#,##0.00";

        // In the repeated row, so this is about one item. Dropped rows are gone before anything else
        // runs — the total below covers what survived, not what was bound.
        sheet.Cell("D5").Value = "<<If test=\"item.IsExport\">>";

        sheet.Cell("B6").Value = "Exported";
        sheet.Cell("C6").Value = "<<Sum>>";
        sheet.Cell("C6").Style.NumberFormat.Format = "#,##0.00";
        sheet.Range("A6:C6").Style.Font.SetBold();
        sheet.Range("A6:C6").Style.Border.SetTopBorder(XLBorderStyleValues.Thin);

        sheet.Workbook.DefinedNames.Add("Sales", sheet.Range("A5:D6"));

        // The helper column carrying the tag has served its purpose by the time the report is read.
        sheet.Cell("D6").Value = "<<Delete>>";
    }

    /// <summary>
    /// An optional section. The test sits in the options row, so it is asked once, about the range.
    /// </summary>
    private static void OptionalSection(IXLWorksheet sheet)
    {
        sheet.Cell("A9").Value = "Appendix — every line, printed only when asked for";
        sheet.Cell("A9").Style.Font.SetBold();

        Headings(sheet, 10, "Product", "Region", "Line total");

        sheet.Cell("A11").Value = "{{ item.Product }}";
        sheet.Cell("B11").Value = "{{ item.Region }}";
        sheet.Cell("C11").Value = "{{ item.Total }}";
        sheet.Cell("C11").Style.NumberFormat.Format = "#,##0.00";

        // In the options row, so this is about the whole range. ShowAppendix is false, so the whole
        // block goes — open the report and the headings above are all that is left.
        sheet.Cell("A12").Value = "<<If test=\"{{ ShowAppendix }}\">>";

        sheet.Workbook.DefinedNames.Add("Appendix", sheet.Range("A11:C12"));
    }

    protected override void AddData(IXLTemplate template)
    {
        var sales = SalesData.Sales();

        template.AddVariable("Sales", sales);
        template.AddVariable("Appendix", sales);
        template.AddVariable("ShowAppendix", false);
    }

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("Conditional");
        var exports = SalesData.Sales().Count(sale => sale.IsExport);

        output.WriteLine($"  Bound     12 lines");
        output.WriteLine($"  Generated {exports} rows   (the exports; the rest were dropped as rows)");
        output.WriteLine($"  Appendix  {(generated.DefinedNames.TryGetValue("Appendix", out _) ? "still there" : "gone, name and all")}");
        output.WriteLine($"  Total     {sheet.Cell(5 + exports, 3).GetFormattedString()}   (over what survived, not what was bound)");
        output.WriteLine();
    }
}
