using XLibur.Excel;

#if ENGINE_PINATA
XLibur.Fonts.PinataLabs.PinataLabsFontBootstrap.Register();
#elif ENGINE_SIXLABORS
XLibur.Excel.LoadOptions.DefaultFontEngine = new XLibur.Fonts.SixLabors.SixLaborsFontEngine("Microsoft Sans Serif");
#elif ENGINE_SKIA
XLibur.Fonts.SkiaSharp.SkiaSharpFontBootstrap.Register();
#endif

var failures = new List<string>();

void Check(string name, object? actual, object? expected)
{
    if (!Equals(actual, expected))
        failures.Add($"{name}: expected <{expected}> but was <{actual}>");
}

var path = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "xlibur-aot-smoke.xlsx");

using (var wb = new XLWorkbook())
{
    var ws = wb.Worksheets.Add("Data");
    ws.Cell("A1").Value = "Quarter";
    ws.Cell("B1").Value = "Sales";
    ws.Cell("C1").Value = "Date";
    string[] quarters = ["Q1", "Q2", "Q3", "Q4"];
    double[] sales = [30.5, 45, 28, 50];
    for (var i = 0; i < 4; i++)
    {
        ws.Cell(i + 2, 1).Value = quarters[i];
        ws.Cell(i + 2, 2).Value = sales[i];
        ws.Cell(i + 2, 3).Value = new DateTime(2026, i * 3 + 1, 15);
    }

    ws.Cell("B7").FormulaA1 = "SUM(B2:B5)";
    ws.Cell("B8").FormulaA1 = "TEXT(B2,\"0.00\")&\"!\"";
    ws.Cell("B9").Value = true;
    ws.Cell("B10").Value = 1234567.891;
    ws.Cell("B10").Style.NumberFormat.Format = "#,##0.00";
    ws.Cell("A1").Style.Font.Bold = true;
    ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.LightGray;
    ws.Range("A1:C5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
    ws.Column(1).AdjustToContents();
    ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
    ws.PageSetup.PaperSize = XLPaperSize.A4Paper;

    Check("recalc SUM", ws.Cell("B7").Value.GetNumber(), 153.5);
    Check("recalc TEXT", ws.Cell("B8").Value.GetText(), "30.50!");
    Check("formatted", ws.Cell("B10").GetFormattedString(), "1,234,567.89");

    var tbl = ws.Range("A1:C5").CreateTable("SalesTable");
    Check("table name", tbl.Name, "SalesTable");

    var chart = ws.Charts.Add(XLChartType.ColumnClustered);
    chart.SetTitle("Sales");
    chart.Series.Add("Sales", "'Data'!$B$2:$B$5", "'Data'!$A$2:$A$5");
    chart.Position.SetColumn(4).SetRow(1);
    chart.SecondPosition.SetColumn(10).SetRow(15);

    wb.SaveAs(path);
}

using (var wb = new XLWorkbook(path))
{
    var ws = wb.Worksheet("Data");
    Check("A2", ws.Cell("A2").Value.GetText(), "Q1");
    Check("B3", ws.Cell("B3").Value.GetNumber(), 45.0);
    Check("C2", ws.Cell("C2").Value.GetDateTime(), new DateTime(2026, 1, 15));
    Check("B9", ws.Cell("B9").Value.GetBoolean(), true);
    Check("B7 cached", ws.Cell("B7").Value.GetNumber(), 153.5);
    Check("bold", ws.Cell("A1").Style.Font.Bold, true);
    Check("landscape", ws.PageSetup.PageOrientation, XLPageOrientation.Landscape);
    Check("paper", ws.PageSetup.PaperSize, XLPaperSize.A4Paper);
    Check("table", ws.Tables.Count(), 1);
    Check("chart", ws.Charts.Count, 1);
    Check("border", ws.Cell("A1").Style.Border.TopBorder, XLBorderStyleValues.Thin);
}

if (failures.Count > 0)
{
    foreach (var f in failures) Console.Error.WriteLine("FAIL " + f);
    return 1;
}

Console.WriteLine("OK");
return 0;
