using System.IO;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// A range that repeats <em>across</em> instead of down: one column per item.
/// </summary>
/// <remarks>
/// <para>
/// Written by putting <c>&lt;&lt;Horizontal&gt;&gt;</c> in the range's last column. Everything then
/// turns ninety degrees: the last <b>column</b> is the options column, the columns before it are what
/// repeats, and a tag sits in a <b>row</b> and acts on that row. So a total goes in the options column
/// beside the row it totals, and the labels a reader reads down the side sit in a column outside the
/// range.
/// </para>
/// <para>
/// It suits a report with few items and many measures — a quarter or a region per column, which is how
/// a summary is usually laid out. <c>&lt;&lt;Group&gt;&gt;</c> and <c>&lt;&lt;AutoFilter&gt;&gt;</c> do
/// not apply across and say so rather than doing something surprising: a subtotal column labelled with
/// a group key is not a thing anyone asks for, and Excel filters rows.
/// </para>
/// </remarks>
public sealed class HorizontalReport : ReportExample
{
    public override string Name => "HorizontalReport";

    public override string Summary => "<<Horizontal>>: one column per item instead of one row.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("By region");

        sheet.Cell("A1").Value = "Sales by region — a column each";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

        // The labels live outside the range, in column A: they are not repeated.
        sheet.Cell("A3").Value = "Region";
        sheet.Cell("A4").Value = "Lines";
        sheet.Cell("A5").Value = "Units";
        sheet.Cell("A6").Value = "Total";
        sheet.Range("A3:A6").Style.Font.SetBold();
        sheet.Cell("A3").Style.Fill.SetBackgroundColor(XLColor.LightGray);

        // Column B is what repeats — one copy per region.
        sheet.Cell("B3").Value = "{{ item.Region }}";
        sheet.Cell("B4").Value = "{{ item.Lines }}";
        sheet.Cell("B5").Value = "{{ item.Units }}";
        sheet.Cell("B6").Value = "{{ item.Total }}";
        sheet.Cell("B3").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray);
        sheet.Cell("B6").Style.NumberFormat.Format = "#,##0.00";

        // Column C is the options column. <<Horizontal>> has to be findable before the range can be
        // read at all, which is why it goes in the last column specifically.
        sheet.Cell("C3").Value = "<<Horizontal>>All regions";
        sheet.Cell("C4").Value = "<<Sum>>";
        sheet.Cell("C5").Value = "<<Sum>>";
        sheet.Cell("C6").Value = "<<Sum>>";
        sheet.Range("C3:C6").Style.Font.SetBold();
        sheet.Range("C3:C6").Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
        sheet.Cell("C6").Style.NumberFormat.Format = "#,##0.00";

        workbook.DefinedNames.Add("Regions", sheet.Range("B3:C6"));

        sheet.Column(1).Width = 10;
        sheet.Columns(2, 6).Width = 14;
    }

    protected override void AddData(IXLTemplate template)
    {
        // Summarised in code rather than by the template: a horizontal range is for measures already
        // reduced to one figure each, not for raw lines.
        var byRegion = SalesData.Sales()
            .GroupBy(sale => sale.Region)
            .OrderBy(group => group.Key)
            .Select(group => new RegionSummary
            {
                Region = group.Key,
                Lines = group.Count(),
                Units = group.Sum(sale => sale.Quantity),
                Total = group.Sum(sale => sale.Total),
            })
            .ToList();

        template.AddVariable("Regions", byRegion);
    }

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("By region");
        var used = sheet.RangeUsed()!.RangeAddress;

        output.WriteLine($"  Columns used  through {used.LastAddress.ColumnLetter}"
            + "   (the template had one data column and one options column)");
        output.WriteLine($"  Row 3 reads   {string.Join(" | ", sheet.Row(3).CellsUsed().Select(cell => cell.GetFormattedString()))}");
        output.WriteLine($"  Row 6 reads   {string.Join(" | ", sheet.Row(6).CellsUsed().Select(cell => cell.GetFormattedString()))}");
        output.WriteLine();
        output.WriteLine("  The last figure on row 6 is the options column's <<Sum>>, totalling across the");
        output.WriteLine("  generated columns rather than down a row.");
        output.WriteLine();
    }

    /// <summary>One region's figures — one column of the report.</summary>
    private sealed class RegionSummary
    {
        public string Region { get; init; } = string.Empty;

        public int Lines { get; init; }

        public int Units { get; init; }

        public decimal Total { get; init; }
    }
}
