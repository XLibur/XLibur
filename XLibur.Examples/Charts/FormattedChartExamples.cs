using XLibur.Excel;

namespace XLibur.Examples.Charts;

/// <summary>
/// Creates a workbook showing what the series formatting API can do: explicit fills and outlines,
/// markers, smoothing, and a series plotted against a secondary value axis.
/// </summary>
public class FormattedChartExamples : IXLExample
{
    public void Create(string filePath)
    {
        using var wb = new XLWorkbook();

        BrandedColumns(wb);
        LineStyles(wb);
        TwoScales(wb);
        HighlightOneSeries(wb);
        Labels(wb);

        wb.SaveAs(filePath);
    }

    /// <summary>
    /// The everyday case: give each series a colour from the corporate palette instead of accepting
    /// Excel's automatic theme colours.
    /// </summary>
    private static void BrandedColumns(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Branded columns");
        WriteQuarterlyData(ws);

        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.SetTitle("Revenue vs cost");

        var revenue = chart.Series.Add("Revenue", "'Branded columns'!$B$2:$B$5", "'Branded columns'!$A$2:$A$5");
        revenue.FillColor = XLColor.FromHtml("#4472C4");
        revenue.LineColor = XLColor.FromHtml("#203864");
        revenue.LineWidthPt = 1;

        // A muted grey keeps the comparison series in the background.
        var cost = chart.Series.Add("Cost", "'Branded columns'!$C$2:$C$5", "'Branded columns'!$A$2:$A$5");
        cost.FillColor = XLColor.FromHtml("#A5A5A5");

        chart.Position.SetColumn(6).SetRow(1);
        chart.SecondPosition.SetColumn(14).SetRow(17);
    }

    /// <summary>
    /// Markers and smoothing on line series.
    /// </summary>
    private static void LineStyles(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Line styles");
        WriteQuarterlyData(ws);

        var chart = ws.Charts.Add(XLChartType.Line);
        chart.SetTitle("Marker styles and smoothing");

        var straight = chart.Series.Add("Revenue", "'Line styles'!$B$2:$B$5", "'Line styles'!$A$2:$A$5");
        straight.LineColor = XLColor.FromHtml("#4472C4");
        straight.LineWidthPt = 2.25;
        straight.MarkerStyle = XLMarkerStyle.Circle;
        straight.MarkerSize = 8;
        straight.MarkerFillColor = XLColor.FromHtml("#FFFFFF");

        var smooth = chart.Series.Add("Cost", "'Line styles'!$C$2:$C$5", "'Line styles'!$A$2:$A$5");
        smooth.LineColor = XLColor.FromTheme(XLThemeColor.Accent6);
        smooth.LineWidthPt = 2.25;
        smooth.MarkerStyle = XLMarkerStyle.Diamond;
        smooth.MarkerSize = 8;
        smooth.Smooth = true;

        chart.Position.SetColumn(6).SetRow(1);
        chart.SecondPosition.SetColumn(14).SetRow(17);
    }

    /// <summary>
    /// A percentage next to values in the thousands only reads if it gets its own axis.
    /// </summary>
    private static void TwoScales(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Two scales");
        WriteQuarterlyData(ws);
        ws.Cell("D1").Value = "Margin %";
        for (var row = 2; row <= 5; row++)
            ws.Cell(row, 4).FormulaA1 = $"=(B{row}-C{row})/B{row}";
        ws.Range("D2:D5").Style.NumberFormat.Format = "0.0%";

        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.SetTitle("Revenue and margin");
        chart.Series.Add("Revenue", "'Two scales'!$B$2:$B$5", "'Two scales'!$A$2:$A$5");

        chart.SecondaryChartType = XLChartType.LineWithMarkers;
        var margin = chart.SecondarySeries.Add("Margin %", "'Two scales'!$D$2:$D$5", "'Two scales'!$A$2:$A$5");
        margin.UseSecondaryAxis = true;
        margin.LineColor = XLColor.FromHtml("#ED7D31");
        margin.LineWidthPt = 2.25;
        margin.MarkerStyle = XLMarkerStyle.Circle;
        margin.MarkerSize = 7;

        chart.Position.SetColumn(6).SetRow(1);
        chart.SecondPosition.SetColumn(15).SetRow(19);
    }

    /// <summary>
    /// Series on the primary chart type can go on the secondary axis too: the two currency series
    /// share the left axis while the growth percentage gets the right one, without the chart needing
    /// a secondary chart type. Only the series worth looking at first is given an explicit colour.
    /// </summary>
    private static void HighlightOneSeries(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Highlight");
        ws.Cell("A1").Value = "Region";
        ws.Cell("B1").Value = "This year";
        ws.Cell("C1").Value = "Last year";
        ws.Cell("D1").Value = "Growth %";

        string[] regions = ["North", "South", "East", "West"];
        double[] thisYear = [42_000, 38_500, 51_000, 29_750];
        double[] lastYear = [39_000, 40_100, 44_500, 31_000];

        for (var i = 0; i < regions.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = regions[i];
            ws.Cell(i + 2, 2).Value = thisYear[i];
            ws.Cell(i + 2, 3).Value = lastYear[i];
            ws.Cell(i + 2, 4).FormulaA1 = $"=(B{i + 2}-C{i + 2})/C{i + 2}";
        }

        ws.Range("B2:C5").Style.NumberFormat.Format = "$ #,##0";
        ws.Range("D2:D5").Style.NumberFormat.Format = "0.0%";
        ws.Columns("A", "D").AdjustToContents();

        var chart = ws.Charts.Add(XLChartType.ColumnClustered);
        chart.SetTitle("Year on year by region");

        chart.Series.Add("This year", "Highlight!$B$2:$B$5", "Highlight!$A$2:$A$5").FillColor =
            XLColor.FromHtml("#4472C4");

        // No colour at all: Excel picks the next theme colour, which is the right default for a
        // series nobody is meant to look at first.
        chart.Series.Add("Last year", "Highlight!$C$2:$C$5", "Highlight!$A$2:$A$5");

        // Growth is a percentage, so it needs the secondary axis even though it is the same
        // (column) chart type as the other two series.
        var growth = chart.Series.Add("Growth %", "Highlight!$D$2:$D$5", "Highlight!$A$2:$A$5");
        growth.UseSecondaryAxis = true;
        growth.FillColor = XLColor.FromHtml("#70AD47");

        chart.Position.SetColumn(6).SetRow(1);
        chart.SecondPosition.SetColumn(15).SetRow(19);
    }

    /// <summary>
    /// Data labels: chart-wide defaults, a per-series override, and the percentage labels a pie chart
    /// usually wants.
    /// </summary>
    private static void Labels(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Labels");
        WriteQuarterlyData(ws);

        var columns = ws.Charts.Add(XLChartType.ColumnClustered);
        columns.SetTitle("Labelled columns");
        columns.Series.Add("Revenue", "Labels!$B$2:$B$5", "Labels!$A$2:$A$5");
        var cost = columns.Series.Add("Cost", "Labels!$C$2:$C$5", "Labels!$A$2:$A$5");

        // Every series gets a value above its column...
        columns.DataLabels.ShowValue = true;
        columns.DataLabels.NumberFormat = "$ #,##0";
        columns.DataLabels.Position = XLDataLabelPosition.OutsideEnd;

        // ...except this one, which puts its label inside and names the series as well.
        cost.DataLabels.ShowValue = true;
        cost.DataLabels.ShowSeriesName = true;
        cost.DataLabels.NumberFormat = "$ #,##0";
        cost.DataLabels.Position = XLDataLabelPosition.InsideEnd;

        columns.Position.SetColumn(6).SetRow(1);
        columns.SecondPosition.SetColumn(15).SetRow(19);

        var pie = ws.Charts.Add(XLChartType.Pie);
        pie.SetTitle("Revenue share");
        var share = pie.Series.Add("Revenue", "Labels!$B$2:$B$5", "Labels!$A$2:$A$5");
        share.DataLabels.ShowCategoryName = true;
        share.DataLabels.ShowPercentage = true;
        share.DataLabels.Position = XLDataLabelPosition.BestFit;

        pie.Position.SetColumn(6).SetRow(21);
        pie.SecondPosition.SetColumn(15).SetRow(39);
    }

    private static void WriteQuarterlyData(IXLWorksheet ws)
    {
        ws.Cell("A1").Value = "Quarter";
        ws.Cell("B1").Value = "Revenue";
        ws.Cell("C1").Value = "Cost";

        string[] quarters = ["Q1", "Q2", "Q3", "Q4"];
        double[] revenue = [30_000, 45_000, 28_000, 50_000];
        double[] cost = [21_000, 29_000, 20_500, 31_000];

        for (var i = 0; i < quarters.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = quarters[i];
            ws.Cell(i + 2, 2).Value = revenue[i];
            ws.Cell(i + 2, 3).Value = cost[i];
        }

        ws.Range("B2:C5").Style.NumberFormat.Format = "$ #,##0";
        ws.Columns("A", "C").AdjustToContents();
    }
}
