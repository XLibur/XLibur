using XLibur.Excel;

namespace XLibur.Examples.Charts;

/// <summary>
/// Creates an example workbook containing every XLChartType (one chart per type, grouped by family).
/// </summary>
public class ChartExamples : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();

        // ── Helper: write shared sample data to a sheet ──
        void WriteData2Col(IXLWorksheet ws, string[] cats, double[] vals)
        {
            ws.Cell("A1").Value = "Category";
            ws.Cell("B1").Value = "Value";
            for (var i = 0; i < cats.Length; i++)
            {
                ws.Cell(i + 2, 1).Value = cats[i];
                ws.Cell(i + 2, 2).Value = vals[i];
            }
            ws.Columns("A", "B").AdjustToContents();
        }

        void WriteData3Col(IXLWorksheet ws, string[] cats, double[] v1, double[] v2)
        {
            ws.Cell("A1").Value = "Category";
            ws.Cell("B1").Value = "Series 1";
            ws.Cell("C1").Value = "Series 2";
            for (var i = 0; i < cats.Length; i++)
            {
                ws.Cell(i + 2, 1).Value = cats[i];
                ws.Cell(i + 2, 2).Value = v1[i];
                ws.Cell(i + 2, 3).Value = v2[i];
            }
            ws.Columns("A", "C").AdjustToContents();
        }

        IXLChart AddChart(IXLWorksheet ws, XLChartType type, string title, string valRef, string catRef, int row, int col = 0, int width = 5)
        {
            var c = ws.Charts.Add(type);
            c.SetTitle(title);
            c.Series.Add("Data", valRef, catRef);
            c.Position.SetColumn(col).SetRow(row);
            c.SecondPosition.SetColumn(col + width).SetRow(row + 14);
            return c;
        }

        IXLChart AddChart2S(IXLWorksheet ws, XLChartType type, string title,
            string v1, string v2, string cat, int row, int col = 0, int width = 5)
        {
            var c = ws.Charts.Add(type);
            c.SetTitle(title);
            c.Series.Add("S1", v1, cat);
            c.Series.Add("S2", v2, cat);
            c.Position.SetColumn(col).SetRow(row);
            c.SecondPosition.SetColumn(col + width).SetRow(row + 14);
            return c;
        }

        var cats4 = new[] { "Q1", "Q2", "Q3", "Q4" };
        var vals4a = new[] { 30.0, 45, 28, 50 };
        var vals4b = new[] { 20.0, 35, 40, 25 };

        // ════════════════════════════════════════════════════════════════
        // Sheet 1: Bar/Column (6 types)
        // ════════════════════════════════════════════════════════════════
        var ws1 = wb.Worksheets.Add("Bar Column");
        WriteData3Col(ws1, cats4, vals4a, vals4b);
        var sn = "'" + ws1.Name + "'";
        var cat = $"{sn}!$A$2:$A$5";
        var v1 = $"{sn}!$B$2:$B$5";
        var v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(ws1, XLChartType.ColumnClustered,          "ColumnClustered",          v1, v2, cat, 7,  0);
        AddChart2S(ws1, XLChartType.ColumnStacked,            "ColumnStacked",            v1, v2, cat, 7,  6);
        AddChart2S(ws1, XLChartType.ColumnStacked100Percent,  "ColumnStacked100%",        v1, v2, cat, 22, 0);
        AddChart2S(ws1, XLChartType.BarClustered,             "BarClustered",             v1, v2, cat, 22, 6);
        AddChart2S(ws1, XLChartType.BarStacked,               "BarStacked",               v1, v2, cat, 37, 0);
        AddChart2S(ws1, XLChartType.BarStacked100Percent,     "BarStacked100%",           v1, v2, cat, 37, 6);

        // ════════════════════════════════════════════════════════════════
        // Sheet 2: Bar/Column 3D (7 types)
        // ════════════════════════════════════════════════════════════════
        var ws2 = wb.Worksheets.Add("Bar Column 3D");
        WriteData3Col(ws2, cats4, vals4a, vals4b);
        sn = "'" + ws2.Name + "'";
        cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(ws2, XLChartType.ColumnClustered3D,          "ColumnClustered3D",          v1, v2, cat, 7,  0);
        AddChart2S(ws2, XLChartType.ColumnStacked3D,            "ColumnStacked3D",            v1, v2, cat, 7,  6);
        AddChart2S(ws2, XLChartType.ColumnStacked100Percent3D,  "ColumnStacked100%3D",        v1, v2, cat, 22, 0);
        AddChart (ws2, XLChartType.Column3D,                    "Column3D",                   v1,     cat, 22, 6);
        AddChart2S(ws2, XLChartType.BarClustered3D,             "BarClustered3D",             v1, v2, cat, 37, 0);
        AddChart2S(ws2, XLChartType.BarStacked3D,               "BarStacked3D",               v1, v2, cat, 37, 6);
        AddChart2S(ws2, XLChartType.BarStacked100Percent3D,     "BarStacked100%3D",           v1, v2, cat, 52, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 3: Line (7 types)
        // ════════════════════════════════════════════════════════════════
        var ws3 = wb.Worksheets.Add("Line");
        WriteData3Col(ws3, cats4, vals4a, vals4b);
        sn = ws3.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(ws3, XLChartType.Line,                              "Line",                    v1, v2, cat, 7,  0);
        AddChart2S(ws3, XLChartType.LineStacked,                       "LineStacked",             v1, v2, cat, 7,  6);
        AddChart2S(ws3, XLChartType.LineStacked100Percent,             "LineStacked100%",         v1, v2, cat, 22, 0);
        AddChart2S(ws3, XLChartType.LineWithMarkers,                   "LineWithMarkers",         v1, v2, cat, 22, 6);
        AddChart2S(ws3, XLChartType.LineWithMarkersStacked,            "LineMarkersStacked",      v1, v2, cat, 37, 0);
        AddChart2S(ws3, XLChartType.LineWithMarkersStacked100Percent,  "LineMarkers100%",         v1, v2, cat, 37, 6);
        AddChart2S(ws3, XLChartType.Line3D,                            "Line3D",                  v1, v2, cat, 52, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 4: Area (6 types)
        // ════════════════════════════════════════════════════════════════
        var ws4 = wb.Worksheets.Add("Area");
        WriteData3Col(ws4, cats4, vals4a, vals4b);
        sn = ws4.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(ws4, XLChartType.Area,                    "Area",               v1, v2, cat, 7,  0);
        AddChart2S(ws4, XLChartType.AreaStacked,             "AreaStacked",        v1, v2, cat, 7,  6);
        AddChart2S(ws4, XLChartType.AreaStacked100Percent,   "AreaStacked100%",    v1, v2, cat, 22, 0);
        AddChart2S(ws4, XLChartType.Area3D,                  "Area3D",             v1, v2, cat, 22, 6);
        AddChart2S(ws4, XLChartType.AreaStacked3D,           "AreaStacked3D",      v1, v2, cat, 37, 0);
        AddChart2S(ws4, XLChartType.AreaStacked100Percent3D, "AreaStacked100%3D",  v1, v2, cat, 37, 6);

        // ════════════════════════════════════════════════════════════════
        // Sheet 5: Pie & Doughnut (8 types)
        // ════════════════════════════════════════════════════════════════
        var ws5 = wb.Worksheets.Add("Pie Doughnut");
        WriteData2Col(ws5, cats4, vals4a);
        sn = "'" + ws5.Name + "'";
        cat = $"{sn}!$A$2:$A$5"; var v = $"{sn}!$B$2:$B$5";

        AddChart(ws5, XLChartType.Pie,              "Pie",              v, cat, 7,  0);
        AddChart(ws5, XLChartType.Pie3D,            "Pie3D",            v, cat, 7,  6);
        AddChart(ws5, XLChartType.PieExploded,      "PieExploded",      v, cat, 22, 0);
        AddChart(ws5, XLChartType.PieExploded3D,    "PieExploded3D",    v, cat, 22, 6);
        AddChart(ws5, XLChartType.PieToPie,         "PieToPie",         v, cat, 37, 0);
        AddChart(ws5, XLChartType.PieToBar,         "PieToBar",         v, cat, 37, 6);
        AddChart(ws5, XLChartType.Doughnut,         "Doughnut",         v, cat, 52, 0);
        AddChart(ws5, XLChartType.DoughnutExploded, "DoughnutExploded", v, cat, 52, 6);

        // ════════════════════════════════════════════════════════════════
        // Sheet 6: Radar (3 types)
        // ════════════════════════════════════════════════════════════════
        var ws6 = wb.Worksheets.Add("Radar");
        var skills = new[] { "C#", "SQL", "DevOps", "Testing", "Design" };
        var skillVals = new[] { 9.0, 7, 6, 8, 4 };
        WriteData2Col(ws6, skills, skillVals);
        sn = ws6.Name; cat = $"{sn}!$A$2:$A$6"; v = $"{sn}!$B$2:$B$6";

        AddChart(ws6, XLChartType.Radar,            "Radar",            v, cat, 8,  0);
        AddChart(ws6, XLChartType.RadarWithMarkers, "RadarWithMarkers", v, cat, 8,  6);
        AddChart(ws6, XLChartType.RadarFilled,      "RadarFilled",      v, cat, 23, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 7: Scatter / XY (5 types)
        // ════════════════════════════════════════════════════════════════
        var ws7 = wb.Worksheets.Add("Scatter");
        ws7.Cell("A1").Value = "X"; ws7.Cell("B1").Value = "Y";
        double[] xs = [1, 2.5, 4, 5.5, 7]; double[] ys = [2.3, 3.1, 5.8, 6.2, 8.9];
        for (var i = 0; i < xs.Length; i++) { ws7.Cell(i + 2, 1).Value = xs[i]; ws7.Cell(i + 2, 2).Value = ys[i]; }
        ws7.Columns("A", "B").AdjustToContents();
        sn = ws7.Name; cat = $"{sn}!$A$2:$A$6"; v = $"{sn}!$B$2:$B$6";

        AddChart(ws7, XLChartType.XYScatterMarkers,                  "ScatterMarkers",         v, cat, 8,  0);
        AddChart(ws7, XLChartType.XYScatterStraightLinesWithMarkers, "StraightLinesMarkers",   v, cat, 8,  6);
        AddChart(ws7, XLChartType.XYScatterStraightLinesNoMarkers,   "StraightLinesNoMarkers", v, cat, 23, 0);
        AddChart(ws7, XLChartType.XYScatterSmoothLinesWithMarkers,   "SmoothLinesMarkers",     v, cat, 23, 6);
        AddChart(ws7, XLChartType.XYScatterSmoothLinesNoMarkers,     "SmoothLinesNoMarkers",   v, cat, 38, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 8: Bubble (2 types)
        // ════════════════════════════════════════════════════════════════
        var ws8 = wb.Worksheets.Add("Bubble");
        ws8.Cell("A1").Value = "X"; ws8.Cell("B1").Value = "Y";
        ws8.Cell("A2").Value = 100; ws8.Cell("B2").Value = 15;
        ws8.Cell("A3").Value = 200; ws8.Cell("B3").Value = 30;
        ws8.Cell("A4").Value = 150; ws8.Cell("B4").Value = 10;
        ws8.Cell("A5").Value = 300; ws8.Cell("B5").Value = 50;
        ws8.Columns("A", "B").AdjustToContents();
        sn = ws8.Name; cat = $"{sn}!$A$2:$A$5"; v = $"{sn}!$B$2:$B$5";

        AddChart(ws8, XLChartType.Bubble,   "Bubble",   v, cat, 7,  0);
        AddChart(ws8, XLChartType.Bubble3D, "Bubble3D", v, cat, 7,  6);

        // ════════════════════════════════════════════════════════════════
        // Sheet 9: Stock (4 types)
        // ════════════════════════════════════════════════════════════════
        var ws9 = wb.Worksheets.Add("Stock");
        ws9.Cell("A1").Value = "Date"; ws9.Cell("B1").Value = "High"; ws9.Cell("C1").Value = "Low"; ws9.Cell("D1").Value = "Close";
        string[] days = ["Mon", "Tue", "Wed", "Thu"];
        double[] highs = [105, 108, 110, 112], lows = [98, 100, 99, 103], closes = [102, 104, 107, 109];
        for (var i = 0; i < days.Length; i++)
        {
            ws9.Cell(i + 2, 1).Value = days[i];
            ws9.Cell(i + 2, 2).Value = highs[i];
            ws9.Cell(i + 2, 3).Value = lows[i];
            ws9.Cell(i + 2, 4).Value = closes[i];
        }
        ws9.Columns("A", "D").AdjustToContents();
        sn = ws9.Name;

        foreach (var (type, title, row) in new[] {
            (XLChartType.StockHighLowClose,           "StockHLC",  7),
            (XLChartType.StockOpenHighLowClose,       "StockOHLC", 22),
            (XLChartType.StockVolumeHighLowClose,     "StockVHLC", 37),
            (XLChartType.StockVolumeOpenHighLowClose, "StockVOHLC",52) })
        {
            var c = ws9.Charts.Add(type);
            c.SetTitle(title);
            c.Series.Add("High", $"{sn}!$B$2:$B$5", $"{sn}!$A$2:$A$5");
            c.Series.Add("Low", $"{sn}!$C$2:$C$5", $"{sn}!$A$2:$A$5");
            c.Series.Add("Close", $"{sn}!$D$2:$D$5", $"{sn}!$A$2:$A$5");
            c.Position.SetColumn(0).SetRow(row);
            c.SecondPosition.SetColumn(8).SetRow(row + 14);
        }

        // ════════════════════════════════════════════════════════════════
        // Sheet 10: Surface (4 types)
        // ════════════════════════════════════════════════════════════════
        var ws10 = wb.Worksheets.Add("Surface");
        ws10.Cell("B1").Value = "C1"; ws10.Cell("C1").Value = "C2"; ws10.Cell("D1").Value = "C3";
        ws10.Cell("A2").Value = "R1"; ws10.Cell("A3").Value = "R2"; ws10.Cell("A4").Value = "R3";
        ws10.Cell("B2").Value = 10; ws10.Cell("C2").Value = 20; ws10.Cell("D2").Value = 30;
        ws10.Cell("B3").Value = 25; ws10.Cell("C3").Value = 15; ws10.Cell("D3").Value = 35;
        ws10.Cell("B4").Value = 30; ws10.Cell("C4").Value = 40; ws10.Cell("D4").Value = 20;
        ws10.Columns("A", "D").AdjustToContents();
        sn = ws10.Name; cat = $"{sn}!$A$2:$A$4";

        foreach (var (type, title, row, col) in new[] {
            (XLChartType.Surface,                 "Surface",                7,  0),
            (XLChartType.SurfaceWireframe,        "SurfaceWireframe",       7,  6),
            (XLChartType.SurfaceContour,          "SurfaceContour",         22, 0),
            (XLChartType.SurfaceContourWireframe, "SurfaceContourWireframe",22, 6) })
        {
            var c = ws10.Charts.Add(type);
            c.SetTitle(title);
            c.Series.Add("C1", $"{sn}!$B$2:$B$4", cat);
            c.Series.Add("C2", $"{sn}!$C$2:$C$4", cat);
            c.Series.Add("C3", $"{sn}!$D$2:$D$4", cat);
            c.Position.SetColumn(col).SetRow(row);
            c.SecondPosition.SetColumn(col + 5).SetRow(row + 14);
        }

        // ════════════════════════════════════════════════════════════════
        // Sheet 11: Combo (Bar + Line)
        // ════════════════════════════════════════════════════════════════
        var ws11 = wb.Worksheets.Add("Combo");
        WriteData3Col(ws11, cats4, vals4a, vals4b);
        sn = ws11.Name;
        var combo = ws11.Charts.Add(XLChartType.ColumnClustered);
        combo.SetTitle("Combo: Columns + Line");
        combo.Series.Add("Columns", $"{sn}!$B$2:$B$5", $"{sn}!$A$2:$A$5");
        combo.SecondaryChartType = XLChartType.Line;
        combo.SecondarySeries.Add("Line", $"{sn}!$C$2:$C$5", $"{sn}!$A$2:$A$5");
        combo.Position.SetColumn(0).SetRow(7);
        combo.SecondPosition.SetColumn(10).SetRow(24);

        // ════════════════════════════════════════════════════════════════
        // Sheet 12: Cone (7 types)
        // ════════════════════════════════════════════════════════════════
        var ws12 = wb.Worksheets.Add("Cone");
        WriteData3Col(ws12, cats4, vals4a, vals4b);
        sn = ws12.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart (ws12, XLChartType.Cone,                          "Cone",                   v1,     cat, 7,  0);
        AddChart (ws12, XLChartType.ConeClustered,                 "ConeClustered",           v1,     cat, 7,  6);
        AddChart (ws12, XLChartType.ConeStacked,                   "ConeStacked",             v1,     cat, 22, 0);
        AddChart (ws12, XLChartType.ConeStacked100Percent,         "ConeStacked100%",         v1,     cat, 22, 6);
        AddChart (ws12, XLChartType.ConeHorizontalClustered,       "ConeHorizClustered",      v1,     cat, 37, 0);
        AddChart2S(ws12, XLChartType.ConeHorizontalStacked,        "ConeHorizStacked",        v1, v2, cat, 37, 6);
        AddChart2S(ws12, XLChartType.ConeHorizontalStacked100Percent,"ConeHoriz100%",         v1, v2, cat, 52, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 13: Cylinder (7 types)
        // ════════════════════════════════════════════════════════════════
        var ws13 = wb.Worksheets.Add("Cylinder");
        WriteData3Col(ws13, cats4, vals4a, vals4b);
        sn = ws13.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart (ws13, XLChartType.Cylinder,                          "Cylinder",                v1,     cat, 7,  0);
        AddChart (ws13, XLChartType.CylinderClustered,                 "CylinderClustered",       v1,     cat, 7,  6);
        AddChart (ws13, XLChartType.CylinderStacked,                   "CylinderStacked",         v1,     cat, 22, 0);
        AddChart (ws13, XLChartType.CylinderStacked100Percent,         "CylinderStacked100%",     v1,     cat, 22, 6);
        AddChart (ws13, XLChartType.CylinderHorizontalClustered,       "CylHorizClustered",       v1,     cat, 37, 0);
        AddChart2S(ws13, XLChartType.CylinderHorizontalStacked,        "CylHorizStacked",         v1, v2, cat, 37, 6);
        AddChart2S(ws13, XLChartType.CylinderHorizontalStacked100Percent,"CylHoriz100%",          v1, v2, cat, 52, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 14: Pyramid (7 types)
        // ════════════════════════════════════════════════════════════════
        var ws14 = wb.Worksheets.Add("Pyramid");
        WriteData3Col(ws14, cats4, vals4a, vals4b);
        sn = ws14.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart (ws14, XLChartType.Pyramid,                          "Pyramid",                v1,     cat, 7,  0);
        AddChart (ws14, XLChartType.PyramidClustered,                 "PyramidClustered",       v1,     cat, 7,  6);
        AddChart (ws14, XLChartType.PyramidStacked,                   "PyramidStacked",         v1,     cat, 22, 0);
        AddChart (ws14, XLChartType.PyramidStacked100Percent,         "PyramidStacked100%",     v1,     cat, 22, 6);
        AddChart (ws14, XLChartType.PyramidHorizontalClustered,       "PyrHorizClustered",      v1,     cat, 37, 0);
        AddChart2S(ws14, XLChartType.PyramidHorizontalStacked,        "PyrHorizStacked",        v1, v2, cat, 37, 6);
        AddChart2S(ws14, XLChartType.PyramidHorizontalStacked100Percent,"PyrHoriz100%",         v1, v2, cat, 52, 0);

        // ════════════════════════════════════════════════════════════════
        // Sheet 15: Extended — Waterfall & Funnel
        // ════════════════════════════════════════════════════════════════
        var ws15 = wb.Worksheets.Add("Waterfall Funnel");
        ws15.Cell("A1").Value = "Step"; ws15.Cell("B1").Value = "Amount";
        string[] steps = ["Start", "Sales", "Returns", "Costs", "End"];
        double[] amounts = [1000, 500, -150, -300, 1050];
        for (var i = 0; i < steps.Length; i++) { ws15.Cell(i + 2, 1).Value = steps[i]; ws15.Cell(i + 2, 2).Value = amounts[i]; }
        ws15.Columns("A", "B").AdjustToContents();
        sn = "'" + ws15.Name + "'";

        var wf = ws15.Charts.Add(XLChartType.Waterfall);
        wf.SetTitle("Waterfall");
        wf.Series.Add("Amount", $"{sn}!$B$2:$B$6", $"{sn}!$A$2:$A$6");
        wf.Position.SetColumn(0).SetRow(8); wf.SecondPosition.SetColumn(6).SetRow(24);

        var fn = ws15.Charts.Add(XLChartType.Funnel);
        fn.SetTitle("Funnel");
        fn.Series.Add("Amount", $"{sn}!$B$2:$B$6", $"{sn}!$A$2:$A$6");
        fn.Position.SetColumn(7).SetRow(8); fn.SecondPosition.SetColumn(13).SetRow(24);

        // ════════════════════════════════════════════════════════════════
        // Sheet 16: Extended — Sunburst & Treemap (hierarchical)
        // ════════════════════════════════════════════════════════════════
        var ws16 = wb.Worksheets.Add("Sunburst Treemap");
        ws16.Cell("A1").Value = "Branch"; ws16.Cell("B1").Value = "Category"; ws16.Cell("C1").Value = "Item"; ws16.Cell("D1").Value = "Value";
        var hierData = new[] {
            ("Food", "Fruit", "Apple", 30), ("Food", "Fruit", "Banana", 25),
            ("Food", "Veg", "Carrot", 15), ("Food", "Veg", "Peas", 10),
            ("Drink", "Hot", "Coffee", 35), ("Drink", "Hot", "Tea", 20),
            ("Drink", "Cold", "Juice", 18) };
        for (var i = 0; i < hierData.Length; i++)
        {
            ws16.Cell(i + 2, 1).Value = hierData[i].Item1;
            ws16.Cell(i + 2, 2).Value = hierData[i].Item2;
            ws16.Cell(i + 2, 3).Value = hierData[i].Item3;
            ws16.Cell(i + 2, 4).Value = hierData[i].Item4;
        }
        ws16.Columns("A", "D").AdjustToContents();
        sn = "'" + ws16.Name + "'";

        var sb = ws16.Charts.Add(XLChartType.Sunburst);
        sb.SetTitle("Sunburst");
        sb.Series.Add("Value", $"{sn}!$D$2:$D$8", $"{sn}!$A$2:$C$8");
        sb.Position.SetColumn(0).SetRow(10); sb.SecondPosition.SetColumn(7).SetRow(28);

        var tm = ws16.Charts.Add(XLChartType.Treemap);
        tm.SetTitle("Treemap");
        tm.Series.Add("Value", $"{sn}!$D$2:$D$8", $"{sn}!$A$2:$C$8");
        tm.Position.SetColumn(8).SetRow(10); tm.SecondPosition.SetColumn(15).SetRow(28);

        // ════════════════════════════════════════════════════════════════
        // Sheet 17: Extended — Box & Whisker
        // ════════════════════════════════════════════════════════════════
        var ws17 = wb.Worksheets.Add("BoxWhisker");
        ws17.Cell("A1").Value = "Group"; ws17.Cell("B1").Value = "Value";
        string[] groups = ["A", "A", "A", "A", "B", "B", "B", "B"];
        double[] bwVals = [12, 15, 18, 22, 8, 14, 20, 25];
        for (var i = 0; i < groups.Length; i++) { ws17.Cell(i + 2, 1).Value = groups[i]; ws17.Cell(i + 2, 2).Value = bwVals[i]; }
        ws17.Columns("A", "B").AdjustToContents();
        sn = ws17.Name;

        var bw = ws17.Charts.Add(XLChartType.BoxWhisker);
        bw.SetTitle("Box & Whisker");
        bw.Series.Add("Value", $"{sn}!$B$2:$B$9", $"{sn}!$A$2:$A$9");
        bw.Position.SetColumn(0).SetRow(11); bw.SecondPosition.SetColumn(9).SetRow(26);

        wb.SaveAs(filePath);
    }
}
