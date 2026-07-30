using System.IO;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// A chart in the template that plots the rows the report generated, not the one row it was drawn
/// over.
/// </summary>
/// <remarks>
/// <para>
/// Nothing in the template says so. The chart's series are drawn over the single template row, because
/// that is all there is to draw over, and the engine stretches them to cover what was generated. There
/// is no tag and nothing to remember: a template author draws the chart on the data and it follows.
/// </para>
/// <para>
/// This is the gap that ClosedXML.Report never closed — its documentation says charts are "not
/// supported", and its issues #123 and #351 are people finding out that a series quietly goes stale.
/// It needed a fix in the core library as well as the engine: setting a loaded chart's series reference
/// used to be a silent no-op, so a re-pointed series was dropped on save.
/// </para>
/// <para>
/// A picture, by contrast, needs nothing from anybody. Its anchor is a live range, so a full-row insert
/// carries it — which is why there is one below the range here, and why the engine contains no picture
/// code at all.
/// </para>
/// </remarks>
public sealed class ChartOverGeneratedRows : ReportExample
{
    public override string Name => "Chart";

    public override string Summary => "A chart series drawn over one template row plots every generated row.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Chart");

        sheet.Cell("A1").Value = "Line total by product";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

        Headings(sheet, 3, "Product", "Line total");

        sheet.Cell("A4").Value = "{{ item.Product }}";
        sheet.Cell("B4").Value = "{{ item.Total }}";
        sheet.Cell("B4").Style.NumberFormat.Format = "#,##0.00";

        // Row 5 is the options row and carries no tags, so it is removed. The range is A4:B5.
        workbook.DefinedNames.Add("Sales", sheet.Range("A4:B5"));

        // Drawn over the one row that exists. In the generated report these read B4:B15 and A4:A15.
        var chart = sheet.Charts.Add(XLChartType.ColumnClustered);
        chart.Title = "Line total by product";
        chart.Series.Add("Line total", "Chart!$B$4:$B$4", "Chart!$A$4:$A$4");
        chart.Position.SetColumn(4).SetRow(3);
        chart.SecondPosition.SetColumn(15).SetRow(25);

        sheet.Column(1).Width = 18;
        sheet.Column(2).Width = 14;
    }

    protected override void AddData(IXLTemplate template) =>
        template.AddVariable("Sales", SalesData.Sales());

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var series = generated.Worksheet("Chart").Charts.Single().Series.Single();

        output.WriteLine($"  Value series     {series.ValueReferences}      (drawn as Chart!$B$4:$B$4)");
        output.WriteLine($"  Category series  {series.CategoryReferences}      (drawn as Chart!$A$4:$A$4)");
        output.WriteLine();
        output.WriteLine("  In Excel: twelve columns with product names along the axis. One column would");
        output.WriteLine("  mean the series had not been re-pointed.");
    }
}
