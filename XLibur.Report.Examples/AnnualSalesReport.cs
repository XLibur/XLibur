using System;
using System.IO;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// The flagship: an annual sales report of the kind someone is actually asked for.
/// </summary>
/// <remarks>
/// <para>
/// A title band bound to workbook variables, a heading row, one row per sale, a per-row formula,
/// grouping by region with subtotals and an Excel outline, a sort inside each group, an autofilter,
/// number and date formats carried from the template, conditional colouring of the margin column, and
/// a grand total.
/// </para>
/// <para>
/// The thing to notice is how little of it is about generation. Nearly every line below is ordinary
/// XLibur formatting of a template that happens to have one data row; the report engine's whole
/// contribution is the six tags in the options row and the defined name. That is the point of the
/// design: a report is authored, not programmed.
/// </para>
/// <para>
/// The conditional format is worth opening the report to see. The template declares <b>one</b> rule
/// over <b>one</b> cell, and the report has one rule over the whole generated block — not one rule per
/// row, which is what ClosedXML.Report produces and what its issue #216 is about.
/// </para>
/// </remarks>
public sealed class AnnualSalesReport : ReportExample
{
    public override string Name => "AnnualSalesReport";

    public override string Summary =>
        "The full thing: bound title, repeated rows, per-row formulas, grouping with subtotals, "
        + "conditional colouring, grand total.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Annual sales");

        Title(sheet);
        Headings(sheet, 5, "Region", "Category", "Product", "Sold on", "Qty", "Unit price", "Line total");

        RepeatedRow(sheet);
        OptionsRow(sheet);

        // One rule, over the one template cell. Generation widens it over every generated row rather
        // than copying it per row.
        sheet.Range("G6:G6").AddConditionalFormat()
            .WhenGreaterThan(500)
            .Fill.SetBackgroundColor(XLColor.LightGreen);
        sheet.Range("G6:G6").AddConditionalFormat()
            .WhenLessThan(100)
            .Fill.SetBackgroundColor(XLColor.LightPink);

        // Rows 6 (repeated) and 7 (options). The name matches the variable holding the collection.
        workbook.DefinedNames.Add("Sales", sheet.Range("A6:G7"));

        sheet.Columns(1, 7).Width = 13;
        sheet.Column(3).Width = 18;
        sheet.SheetView.FreezeRows(5);
    }

    /// <summary>
    /// The title band. Ordinary cells with expressions in them: everything outside a bound range is
    /// evaluated once, against the workbook variables.
    /// </summary>
    private static void Title(IXLWorksheet sheet)
    {
        sheet.Cell("A1").Value = "{{ Company }}";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(18);

        sheet.Cell("A2").Value = "Annual sales — {{ Year }}";
        sheet.Cell("A2").Style.Font.SetFontSize(13).Font.SetFontColor(XLColor.FromArgb(70, 70, 70));

        sheet.Cell("A3").Value = "Generated {{ RunOn }} · green above {{ StrongLine }}, pink below {{ WeakLine }}";
        sheet.Cell("A3").Style.Font.SetItalic().Font.SetFontColor(XLColor.Gray);
        sheet.Cell("A3").Style.NumberFormat.Format = "@";

        sheet.Range("A1:G1").Merge();
        sheet.Range("A2:G2").Merge();
        sheet.Range("A3:G3").Merge();
    }

    /// <summary>The row that becomes one row per sale.</summary>
    private static void RepeatedRow(IXLWorksheet sheet)
    {
        sheet.Cell("A6").Value = "{{ item.Region }}";
        sheet.Cell("B6").Value = "{{ item.Category }}";
        sheet.Cell("C6").Value = "{{ item.Product }}";
        sheet.Cell("D6").Value = "{{ item.SoldOn }}";
        sheet.Cell("E6").Value = "{{ item.Quantity }}";
        sheet.Cell("F6").Value = "{{ item.UnitPrice }}";

        // A real formula, not an expression. The copies are re-pointed at their own rows by the core
        // library's insert-and-copy, so the report ends up with E7*F7, E8*F8 and so on — no different
        // from filling a formula down by hand in Excel.
        sheet.Cell("G6").FormulaA1 = "E6*F6";

        sheet.Cell("D6").Style.DateFormat.Format = "dd MMM yyyy";
        sheet.Range("F6:G6").Style.NumberFormat.Format = "#,##0.00";
    }

    /// <summary>
    /// The options row: six tags and a label, and the whole of what makes this a report rather than a
    /// table.
    /// </summary>
    /// <remarks>
    /// Its styling matters twice over. It is the grand-total row's styling, and it is what every
    /// group's subtotal row is given — the only way a template can say what a row that does not exist
    /// yet should look like.
    /// </remarks>
    private static void OptionsRow(IXLWorksheet sheet)
    {
        // Group by region, merging each region's label down its rows. <<Group>> takes the column's own
        // expression as its key, so there is no need to name item.Region twice.
        sheet.Cell("A7").Value = "<<Group merge>>";

        // Sort by product inside each group. Grouping's ordering is stable, so this survives it.
        sheet.Cell("C7").Value = "<<Sort>>";

        // Totals. Each appears in every region's subtotal row over that region's rows, and again here
        // over the lot. SUBTOTAL ignores nested SUBTOTALs, so the grand total does not count the data
        // twice.
        sheet.Cell("E7").Value = "<<Sum>>";
        sheet.Cell("G7").Value = "<<Sum>>";

        sheet.Cell("F7").Value = "Total";

        // Layout, last: the filter covers the heading row too, and the columns are fitted once the
        // totals are in them.
        sheet.Cell("B7").Value = "<<AutoFilter>><<ColsFit>>";

        sheet.Range("A7:G7").Style.Font.SetBold();
        sheet.Range("A7:G7").Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        sheet.Range("A7:G7").Style.Fill.SetBackgroundColor(XLColor.FromArgb(240, 240, 240));
        sheet.Cell("E7").Style.NumberFormat.Format = "#,##0";
        sheet.Cell("G7").Style.NumberFormat.Format = "#,##0.00";
        sheet.Cell("A7").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
    }

    protected override void AddData(IXLTemplate template)
    {
        template.AddVariable("Company", "Contoso Horticultural Supplies");
        template.AddVariable("Year", 2026);
        template.AddVariable("RunOn", new DateTime(2026, 7, 30, 0, 0, 0, DateTimeKind.Utc).ToString("d MMMM yyyy"));
        template.AddVariable("StrongLine", 500);
        template.AddVariable("WeakLine", 100);
        template.AddVariable("Sales", SalesData.Sales());
    }

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("Annual sales");

        output.WriteLine($"  Rows used            through {sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber}"
            + "   (the template had one data row)");
        output.WriteLine($"  Conditional formats  {sheet.ConditionalFormats.Count()}"
            + "   (the template declared 2 — one per row would be the #216 bug)");
        output.WriteLine($"  Merged ranges        {sheet.MergedRanges.Count}"
            + "   (three title lines, plus one per region)");
        output.WriteLine($"  Deepest outline      {sheet.Rows().Max(row => row.OutlineLevel)}"
            + "   (one group level)");
        output.WriteLine();
        output.WriteLine("  In Excel: the outline buttons down the left collapse to region subtotals and");
        output.WriteLine("  again to the grand total; each region's label is merged down its rows; and the");
        output.WriteLine("  grand total equals the sum of the twelve lines rather than double it.");
    }
}
