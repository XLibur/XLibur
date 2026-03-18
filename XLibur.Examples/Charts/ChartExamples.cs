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

        IXLWorksheet ws = null!;

        IXLChart AddChart(XLChartType type, string title, string valRef, string catRef, (int Row, int Col) pos)
        {
            var c = ws.Charts.Add(type);
            c.SetTitle(title);
            c.Series.Add("Data", valRef, catRef);
            c.Position.SetColumn(pos.Col).SetRow(pos.Row);
            c.SecondPosition.SetColumn(pos.Col + 5).SetRow(pos.Row + 14);
            return c;
        }

        IXLChart AddChart2S(XLChartType type, string title,
            string v1, string v2, string cat, (int Row, int Col) pos)
        {
            var c = ws.Charts.Add(type);
            c.SetTitle(title);
            c.Series.Add("S1", v1, cat);
            c.Series.Add("S2", v2, cat);
            c.Position.SetColumn(pos.Col).SetRow(pos.Row);
            c.SecondPosition.SetColumn(pos.Col + 5).SetRow(pos.Row + 14);
            return c;
        }

        var cats4 = new[] { "Q1", "Q2", "Q3", "Q4" };
        var vals4a = new[] { 30.0, 45, 28, 50 };
        var vals4b = new[] { 20.0, 35, 40, 25 };

        // ════════════════════════════════════════════════════════════════
        // Sheet 1: Bar/Column (6 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Bar Column");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        var sn = "'" + ws.Name + "'";
        var cat = $"{sn}!$A$2:$A$5";
        var v1 = $"{sn}!$B$2:$B$5";
        var v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(XLChartType.ColumnClustered, "ColumnClustered", v1, v2, cat, (7, 0));
        AddChart2S(XLChartType.ColumnStacked, "ColumnStacked", v1, v2, cat, (7, 6));
        AddChart2S(XLChartType.ColumnStacked100Percent, "ColumnStacked100%", v1, v2, cat, (22, 0));
        AddChart2S(XLChartType.BarClustered, "BarClustered", v1, v2, cat, (22, 6));
        AddChart2S(XLChartType.BarStacked, "BarStacked", v1, v2, cat, (37, 0));
        AddChart2S(XLChartType.BarStacked100Percent, "BarStacked100%", v1, v2, cat, (37, 6));

        // ════════════════════════════════════════════════════════════════
        // Sheet 2: Bar/Column 3D (7 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Bar Column 3D");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = "'" + ws.Name + "'";
        cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(XLChartType.ColumnClustered3D, "ColumnClustered3D", v1, v2, cat, (7, 0));
        AddChart2S(XLChartType.ColumnStacked3D, "ColumnStacked3D", v1, v2, cat, (7, 6));
        AddChart2S(XLChartType.ColumnStacked100Percent3D, "ColumnStacked100%3D", v1, v2, cat, (22, 0));
        AddChart(XLChartType.Column3D, "Column3D", v1, cat, (22, 6));
        AddChart2S(XLChartType.BarClustered3D, "BarClustered3D", v1, v2, cat, (37, 0));
        AddChart2S(XLChartType.BarStacked3D, "BarStacked3D", v1, v2, cat, (37, 6));
        AddChart2S(XLChartType.BarStacked100Percent3D, "BarStacked100%3D", v1, v2, cat, (52, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 3: Line (7 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Line");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = ws.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(XLChartType.Line, "Line", v1, v2, cat, (7, 0));
        AddChart2S(XLChartType.LineStacked, "LineStacked", v1, v2, cat, (7, 6));
        AddChart2S(XLChartType.LineStacked100Percent, "LineStacked100%", v1, v2, cat, (22, 0));
        AddChart2S(XLChartType.LineWithMarkers, "LineWithMarkers", v1, v2, cat, (22, 6));
        AddChart2S(XLChartType.LineWithMarkersStacked, "LineMarkersStacked", v1, v2, cat, (37, 0));
        AddChart2S(XLChartType.LineWithMarkersStacked100Percent, "LineMarkers100%", v1, v2, cat, (37, 6));
        AddChart2S(XLChartType.Line3D, "Line3D", v1, v2, cat, (52, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 4: Area (6 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Area");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = ws.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart2S(XLChartType.Area, "Area", v1, v2, cat, (7, 0));
        AddChart2S(XLChartType.AreaStacked, "AreaStacked", v1, v2, cat, (7, 6));
        AddChart2S(XLChartType.AreaStacked100Percent, "AreaStacked100%", v1, v2, cat, (22, 0));
        AddChart2S(XLChartType.Area3D, "Area3D", v1, v2, cat, (22, 6));
        AddChart2S(XLChartType.AreaStacked3D, "AreaStacked3D", v1, v2, cat, (37, 0));
        AddChart2S(XLChartType.AreaStacked100Percent3D, "AreaStacked100%3D", v1, v2, cat, (37, 6));

        // ════════════════════════════════════════════════════════════════
        // Sheet 5: Pie & Doughnut (8 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Pie Doughnut");
        WriteData2Col(ws, cats4, vals4a);
        sn = "'" + ws.Name + "'";
        cat = $"{sn}!$A$2:$A$5"; var v = $"{sn}!$B$2:$B$5";

        AddChart(XLChartType.Pie, "Pie", v, cat, (7, 0));
        AddChart(XLChartType.Pie3D, "Pie3D", v, cat, (7, 6));
        AddChart(XLChartType.PieExploded, "PieExploded", v, cat, (22, 0));
        AddChart(XLChartType.PieExploded3D, "PieExploded3D", v, cat, (22, 6));
        AddChart(XLChartType.PieToPie, "PieToPie", v, cat, (37, 0));
        AddChart(XLChartType.PieToBar, "PieToBar", v, cat, (37, 6));
        AddChart(XLChartType.Doughnut, "Doughnut", v, cat, (52, 0));
        AddChart(XLChartType.DoughnutExploded, "DoughnutExploded", v, cat, (52, 6));

        // ════════════════════════════════════════════════════════════════
        // Sheet 6: Radar (3 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Radar");
        var skills = new[] { "C#", "SQL", "DevOps", "Testing", "Design" };
        var skillVals = new[] { 9.0, 7, 6, 8, 4 };
        WriteData2Col(ws, skills, skillVals);
        sn = ws.Name; cat = $"{sn}!$A$2:$A$6"; v = $"{sn}!$B$2:$B$6";

        AddChart(XLChartType.Radar, "Radar", v, cat, (8, 0));
        AddChart(XLChartType.RadarWithMarkers, "RadarWithMarkers", v, cat, (8, 6));
        AddChart(XLChartType.RadarFilled, "RadarFilled", v, cat, (23, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 7: Scatter / XY (5 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Scatter");
        ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = "Y";
        double[] xs = [1, 2.5, 4, 5.5, 7]; double[] ys = [2.3, 3.1, 5.8, 6.2, 8.9];
        for (var i = 0; i < xs.Length; i++) { ws.Cell(i + 2, 1).Value = xs[i]; ws.Cell(i + 2, 2).Value = ys[i]; }
        ws.Columns("A", "B").AdjustToContents();
        sn = ws.Name; cat = $"{sn}!$A$2:$A$6"; v = $"{sn}!$B$2:$B$6";

        AddChart(XLChartType.XYScatterMarkers, "ScatterMarkers", v, cat, (8, 0));
        AddChart(XLChartType.XYScatterStraightLinesWithMarkers, "StraightLinesMarkers", v, cat, (8, 6));
        AddChart(XLChartType.XYScatterStraightLinesNoMarkers, "StraightLinesNoMarkers", v, cat, (23, 0));
        AddChart(XLChartType.XYScatterSmoothLinesWithMarkers, "SmoothLinesMarkers", v, cat, (23, 6));
        AddChart(XLChartType.XYScatterSmoothLinesNoMarkers, "SmoothLinesNoMarkers", v, cat, (38, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 8: Bubble (2 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Bubble");
        ws.Cell("A1").Value = "X"; ws.Cell("B1").Value = "Y";
        ws.Cell("A2").Value = 100; ws.Cell("B2").Value = 15;
        ws.Cell("A3").Value = 200; ws.Cell("B3").Value = 30;
        ws.Cell("A4").Value = 150; ws.Cell("B4").Value = 10;
        ws.Cell("A5").Value = 300; ws.Cell("B5").Value = 50;
        ws.Columns("A", "B").AdjustToContents();
        sn = ws.Name; cat = $"{sn}!$A$2:$A$5"; v = $"{sn}!$B$2:$B$5";

        AddChart(XLChartType.Bubble, "Bubble", v, cat, (7, 0));
        AddChart(XLChartType.Bubble3D, "Bubble3D", v, cat, (7, 6));

        // ════════════════════════════════════════════════════════════════
        // Sheet 9: Stock (4 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Stock");
        ws.Cell("A1").Value = "Date"; ws.Cell("B1").Value = "High"; ws.Cell("C1").Value = "Low"; ws.Cell("D1").Value = "Close";
        string[] days = ["Mon", "Tue", "Wed", "Thu"];
        double[] highs = [105, 108, 110, 112], lows = [98, 100, 99, 103], closes = [102, 104, 107, 109];
        for (var i = 0; i < days.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = days[i];
            ws.Cell(i + 2, 2).Value = highs[i];
            ws.Cell(i + 2, 3).Value = lows[i];
            ws.Cell(i + 2, 4).Value = closes[i];
        }
        ws.Columns("A", "D").AdjustToContents();
        sn = ws.Name;

        foreach (var (type, title, row) in new[] {
            (XLChartType.StockHighLowClose,           "StockHLC",  7),
            (XLChartType.StockOpenHighLowClose,       "StockOHLC", 22),
            (XLChartType.StockVolumeHighLowClose,     "StockVHLC", 37),
            (XLChartType.StockVolumeOpenHighLowClose, "StockVOHLC",52) })
        {
            var c = ws.Charts.Add(type);
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
        ws = wb.Worksheets.Add("Surface");
        ws.Cell("B1").Value = "C1"; ws.Cell("C1").Value = "C2"; ws.Cell("D1").Value = "C3";
        ws.Cell("A2").Value = "R1"; ws.Cell("A3").Value = "R2"; ws.Cell("A4").Value = "R3";
        ws.Cell("B2").Value = 10; ws.Cell("C2").Value = 20; ws.Cell("D2").Value = 30;
        ws.Cell("B3").Value = 25; ws.Cell("C3").Value = 15; ws.Cell("D3").Value = 35;
        ws.Cell("B4").Value = 30; ws.Cell("C4").Value = 40; ws.Cell("D4").Value = 20;
        ws.Columns("A", "D").AdjustToContents();
        sn = ws.Name; cat = $"{sn}!$A$2:$A$4";

        foreach (var (type, title, row, col) in new[] {
            (XLChartType.Surface,                 "Surface",                7,  0),
            (XLChartType.SurfaceWireframe,        "SurfaceWireframe",       7,  6),
            (XLChartType.SurfaceContour,          "SurfaceContour",         22, 0),
            (XLChartType.SurfaceContourWireframe, "SurfaceContourWireframe",22, 6) })
        {
            var c = ws.Charts.Add(type);
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
        ws = wb.Worksheets.Add("Combo");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = ws.Name;
        var combo = ws.Charts.Add(XLChartType.ColumnClustered);
        combo.SetTitle("Combo: Columns + Line");
        combo.Series.Add("Columns", $"{sn}!$B$2:$B$5", $"{sn}!$A$2:$A$5");
        combo.SecondaryChartType = XLChartType.Line;
        combo.SecondarySeries.Add("Line", $"{sn}!$C$2:$C$5", $"{sn}!$A$2:$A$5");
        combo.Position.SetColumn(0).SetRow(7);
        combo.SecondPosition.SetColumn(10).SetRow(24);

        // ════════════════════════════════════════════════════════════════
        // Sheet 12: Cone (7 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Cone");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = ws.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart(XLChartType.Cone, "Cone", v1, cat, (7, 0));
        AddChart(XLChartType.ConeClustered, "ConeClustered", v1, cat, (7, 6));
        AddChart(XLChartType.ConeStacked, "ConeStacked", v1, cat, (22, 0));
        AddChart(XLChartType.ConeStacked100Percent, "ConeStacked100%", v1, cat, (22, 6));
        AddChart(XLChartType.ConeHorizontalClustered, "ConeHorizClustered", v1, cat, (37, 0));
        AddChart2S(XLChartType.ConeHorizontalStacked, "ConeHorizStacked", v1, v2, cat, (37, 6));
        AddChart2S(XLChartType.ConeHorizontalStacked100Percent, "ConeHoriz100%", v1, v2, cat, (52, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 13: Cylinder (7 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Cylinder");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = ws.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart(XLChartType.Cylinder, "Cylinder", v1, cat, (7, 0));
        AddChart(XLChartType.CylinderClustered, "CylinderClustered", v1, cat, (7, 6));
        AddChart(XLChartType.CylinderStacked, "CylinderStacked", v1, cat, (22, 0));
        AddChart(XLChartType.CylinderStacked100Percent, "CylinderStacked100%", v1, cat, (22, 6));
        AddChart(XLChartType.CylinderHorizontalClustered, "CylHorizClustered", v1, cat, (37, 0));
        AddChart2S(XLChartType.CylinderHorizontalStacked, "CylHorizStacked", v1, v2, cat, (37, 6));
        AddChart2S(XLChartType.CylinderHorizontalStacked100Percent, "CylHoriz100%", v1, v2, cat, (52, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 14: Pyramid (7 types)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Pyramid");
        WriteData3Col(ws, cats4, vals4a, vals4b);
        sn = ws.Name; cat = $"{sn}!$A$2:$A$5"; v1 = $"{sn}!$B$2:$B$5"; v2 = $"{sn}!$C$2:$C$5";

        AddChart(XLChartType.Pyramid, "Pyramid", v1, cat, (7, 0));
        AddChart(XLChartType.PyramidClustered, "PyramidClustered", v1, cat, (7, 6));
        AddChart(XLChartType.PyramidStacked, "PyramidStacked", v1, cat, (22, 0));
        AddChart(XLChartType.PyramidStacked100Percent, "PyramidStacked100%", v1, cat, (22, 6));
        AddChart(XLChartType.PyramidHorizontalClustered, "PyrHorizClustered", v1, cat, (37, 0));
        AddChart2S(XLChartType.PyramidHorizontalStacked, "PyrHorizStacked", v1, v2, cat, (37, 6));
        AddChart2S(XLChartType.PyramidHorizontalStacked100Percent, "PyrHoriz100%", v1, v2, cat, (52, 0));

        // ════════════════════════════════════════════════════════════════
        // Sheet 15: Extended — Waterfall & Funnel
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Waterfall Funnel");
        ws.Cell("A1").Value = "Step"; ws.Cell("B1").Value = "Amount";
        string[] steps = ["Start", "Sales", "Returns", "Costs", "End"];
        double[] amounts = [1000, 500, -150, -300, 1050];
        for (var i = 0; i < steps.Length; i++) { ws.Cell(i + 2, 1).Value = steps[i]; ws.Cell(i + 2, 2).Value = amounts[i]; }
        ws.Columns("A", "B").AdjustToContents();
        sn = "'" + ws.Name + "'";

        var wf = ws.Charts.Add(XLChartType.Waterfall);
        wf.SetTitle("Waterfall");
        wf.Series.Add("Amount", $"{sn}!$B$2:$B$6", $"{sn}!$A$2:$A$6");
        wf.Position.SetColumn(0).SetRow(8); wf.SecondPosition.SetColumn(6).SetRow(24);

        var fn = ws.Charts.Add(XLChartType.Funnel);
        fn.SetTitle("Funnel");
        fn.Series.Add("Amount", $"{sn}!$B$2:$B$6", $"{sn}!$A$2:$A$6");
        fn.Position.SetColumn(7).SetRow(8); fn.SecondPosition.SetColumn(13).SetRow(24);

        // ════════════════════════════════════════════════════════════════
        // Sheet 16: Extended — Sunburst & Treemap (hierarchical)
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("Sunburst Treemap");
        ws.Cell("A1").Value = "Branch"; ws.Cell("B1").Value = "Category"; ws.Cell("C1").Value = "Item"; ws.Cell("D1").Value = "Value";
        var hierData = new[] {
            ("Food", "Fruit", "Apple", 30), ("Food", "Fruit", "Banana", 25),
            ("Food", "Veg", "Carrot", 15), ("Food", "Veg", "Peas", 10),
            ("Drink", "Hot", "Coffee", 35), ("Drink", "Hot", "Tea", 20),
            ("Drink", "Cold", "Juice", 18) };
        for (var i = 0; i < hierData.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = hierData[i].Item1;
            ws.Cell(i + 2, 2).Value = hierData[i].Item2;
            ws.Cell(i + 2, 3).Value = hierData[i].Item3;
            ws.Cell(i + 2, 4).Value = hierData[i].Item4;
        }
        ws.Columns("A", "D").AdjustToContents();
        sn = "'" + ws.Name + "'";

        var sb = ws.Charts.Add(XLChartType.Sunburst);
        sb.SetTitle("Sunburst");
        sb.Series.Add("Value", $"{sn}!$D$2:$D$8", $"{sn}!$A$2:$C$8");
        sb.Position.SetColumn(0).SetRow(10); sb.SecondPosition.SetColumn(7).SetRow(28);

        var tm = ws.Charts.Add(XLChartType.Treemap);
        tm.SetTitle("Treemap");
        tm.Series.Add("Value", $"{sn}!$D$2:$D$8", $"{sn}!$A$2:$C$8");
        tm.Position.SetColumn(8).SetRow(10); tm.SecondPosition.SetColumn(15).SetRow(28);

        // ════════════════════════════════════════════════════════════════
        // Sheet 17: Extended — Box & Whisker
        // ════════════════════════════════════════════════════════════════
        ws = wb.Worksheets.Add("BoxWhisker");
        ws.Cell("A1").Value = "Group"; ws.Cell("B1").Value = "Value";
        string[] groups = ["A", "A", "A", "A", "B", "B", "B", "B"];
        double[] bwVals = [12, 15, 18, 22, 8, 14, 20, 25];
        for (var i = 0; i < groups.Length; i++) { ws.Cell(i + 2, 1).Value = groups[i]; ws.Cell(i + 2, 2).Value = bwVals[i]; }
        ws.Columns("A", "B").AdjustToContents();
        sn = ws.Name;

        var bw = ws.Charts.Add(XLChartType.BoxWhisker);
        bw.SetTitle("Box & Whisker");
        bw.Series.Add("Value", $"{sn}!$B$2:$B$9", $"{sn}!$A$2:$A$9");
        bw.Position.SetColumn(0).SetRow(11); bw.SecondPosition.SetColumn(9).SetRow(26);

        wb.SaveAs(filePath);
    }
}
