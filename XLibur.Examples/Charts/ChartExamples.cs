using XLibur.Excel;

namespace XLibur.Examples.Charts;

public class ChartExamples : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();

        // --- Sheet 1: Column Clustered chart with two series ---
        var ws1 = wb.Worksheets.Add("Sales");

        // Headers
        ws1.Cell("A1").Value = "Quarter";
        ws1.Cell("B1").Value = "North";
        ws1.Cell("C1").Value = "South";

        // Data
        ws1.Cell("A2").Value = "Q1";
        ws1.Cell("A3").Value = "Q2";
        ws1.Cell("A4").Value = "Q3";
        ws1.Cell("A5").Value = "Q4";

        ws1.Cell("B2").Value = 12000;
        ws1.Cell("B3").Value = 15000;
        ws1.Cell("B4").Value = 18000;
        ws1.Cell("B5").Value = 21000;

        ws1.Cell("C2").Value = 9000;
        ws1.Cell("C3").Value = 11000;
        ws1.Cell("C4").Value = 14000;
        ws1.Cell("C5").Value = 17000;

        ws1.Columns("A", "C").AdjustToContents();

        // Create a column clustered chart comparing North vs South
        var chart1 = ws1.Charts.Add(XLChartType.ColumnClustered);
        chart1.SetTitle("Regional Sales by Quarter");
        chart1.Series.Add("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
        chart1.Series.Add("South", "Sales!$C$2:$C$5", "Sales!$A$2:$A$5");
        chart1.Position.SetColumn(0).SetRow(7);
        chart1.SecondPosition.SetColumn(8).SetRow(22);

        // --- Sheet 2: Bar chart (horizontal) ---
        var ws2 = wb.Worksheets.Add("Products");

        ws2.Cell("A1").Value = "Product";
        ws2.Cell("B1").Value = "Units Sold";

        ws2.Cell("A2").Value = "Widget A";
        ws2.Cell("A3").Value = "Widget B";
        ws2.Cell("A4").Value = "Widget C";
        ws2.Cell("A5").Value = "Widget D";
        ws2.Cell("A6").Value = "Widget E";

        ws2.Cell("B2").Value = 340;
        ws2.Cell("B3").Value = 520;
        ws2.Cell("B4").Value = 180;
        ws2.Cell("B5").Value = 410;
        ws2.Cell("B6").Value = 290;

        ws2.Columns("A", "B").AdjustToContents();

        // Create a horizontal bar chart
        var chart2 = ws2.Charts.Add(XLChartType.BarClustered);
        chart2.SetTitle("Product Comparison");
        chart2.Series.Add("Units Sold", "Products!$B$2:$B$6", "Products!$A$2:$A$6");
        chart2.Position.SetColumn(0).SetRow(8);
        chart2.SecondPosition.SetColumn(8).SetRow(22);

        // --- Sheet 3: Chart without a title ---
        var ws3 = wb.Worksheets.Add("Expenses");

        ws3.Cell("A1").Value = "Month";
        ws3.Cell("B1").Value = "Amount";

        ws3.Cell("A2").Value = "Jan";
        ws3.Cell("A3").Value = "Feb";
        ws3.Cell("A4").Value = "Mar";

        ws3.Cell("B2").Value = 4200;
        ws3.Cell("B3").Value = 3800;
        ws3.Cell("B4").Value = 4500;

        ws3.Columns("A", "B").AdjustToContents();

        // A simple chart with no title
        var chart3 = ws3.Charts.Add(XLChartType.ColumnClustered);
        chart3.Series.Add("Amount", "Expenses!$B$2:$B$4", "Expenses!$A$2:$A$4");
        chart3.Position.SetColumn(0).SetRow(5);
        chart3.SecondPosition.SetColumn(7).SetRow(18);

        // --- Sheet 4: Pie chart ---
        var ws4 = wb.Worksheets.Add("Market Share");

        ws4.Cell("A1").Value = "Company";
        ws4.Cell("B1").Value = "Share %";

        ws4.Cell("A2").Value = "Acme Corp";
        ws4.Cell("A3").Value = "Globex";
        ws4.Cell("A4").Value = "Initech";
        ws4.Cell("A5").Value = "Umbrella";

        ws4.Cell("B2").Value = 35;
        ws4.Cell("B3").Value = 28;
        ws4.Cell("B4").Value = 22;
        ws4.Cell("B5").Value = 15;

        ws4.Columns("A", "B").AdjustToContents();

        var chart4 = ws4.Charts.Add(XLChartType.Pie);
        chart4.SetTitle("Market Share");
        chart4.Series.Add("Share %", "'Market Share'!$B$2:$B$5", "'Market Share'!$A$2:$A$5");
        chart4.Position.SetColumn(0).SetRow(7);
        chart4.SecondPosition.SetColumn(8).SetRow(22);

        // --- Sheet 5: Stacked bar chart ---
        var ws5 = wb.Worksheets.Add("Staff Hours");

        ws5.Cell("A1").Value = "Department";
        ws5.Cell("B1").Value = "Regular";
        ws5.Cell("C1").Value = "Overtime";

        ws5.Cell("A2").Value = "Engineering";
        ws5.Cell("A3").Value = "Sales";
        ws5.Cell("A4").Value = "Support";
        ws5.Cell("A5").Value = "Marketing";

        ws5.Cell("B2").Value = 160;
        ws5.Cell("B3").Value = 140;
        ws5.Cell("B4").Value = 150;
        ws5.Cell("B5").Value = 130;

        ws5.Cell("C2").Value = 40;
        ws5.Cell("C3").Value = 25;
        ws5.Cell("C4").Value = 35;
        ws5.Cell("C5").Value = 10;

        ws5.Columns("A", "C").AdjustToContents();

        var chart5 = ws5.Charts.Add(XLChartType.BarStacked);
        chart5.SetTitle("Staff Hours by Department");
        chart5.Series.Add("Regular", "'Staff Hours'!$B$2:$B$5", "'Staff Hours'!$A$2:$A$5");
        chart5.Series.Add("Overtime", "'Staff Hours'!$C$2:$C$5", "'Staff Hours'!$A$2:$A$5");
        chart5.Position.SetColumn(0).SetRow(7);
        chart5.SecondPosition.SetColumn(9).SetRow(22);

        // --- Sheet 6: Line chart ---
        var ws6 = wb.Worksheets.Add("Trends");

        ws6.Cell("A1").Value = "Month";
        ws6.Cell("B1").Value = "Revenue";
        ws6.Cell("C1").Value = "Costs";

        ws6.Cell("A2").Value = "Jan";
        ws6.Cell("A3").Value = "Feb";
        ws6.Cell("A4").Value = "Mar";
        ws6.Cell("A5").Value = "Apr";
        ws6.Cell("A6").Value = "May";
        ws6.Cell("A7").Value = "Jun";

        ws6.Cell("B2").Value = 50000;
        ws6.Cell("B3").Value = 52000;
        ws6.Cell("B4").Value = 48000;
        ws6.Cell("B5").Value = 55000;
        ws6.Cell("B6").Value = 60000;
        ws6.Cell("B7").Value = 63000;

        ws6.Cell("C2").Value = 42000;
        ws6.Cell("C3").Value = 44000;
        ws6.Cell("C4").Value = 41000;
        ws6.Cell("C5").Value = 43000;
        ws6.Cell("C6").Value = 46000;
        ws6.Cell("C7").Value = 47000;

        ws6.Columns("A", "C").AdjustToContents();

        var chart6 = ws6.Charts.Add(XLChartType.LineWithMarkers);
        chart6.SetTitle("Revenue vs Costs");
        chart6.Series.Add("Revenue", "Trends!$B$2:$B$7", "Trends!$A$2:$A$7");
        chart6.Series.Add("Costs", "Trends!$C$2:$C$7", "Trends!$A$2:$A$7");
        chart6.Position.SetColumn(0).SetRow(9);
        chart6.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 7: Combo chart (columns + line overlay) ---
        var ws7 = wb.Worksheets.Add("Combo");

        ws7.Cell("A1").Value = "Quarter";
        ws7.Cell("B1").Value = "Units";
        ws7.Cell("C1").Value = "Avg Price";

        ws7.Cell("A2").Value = "Q1";
        ws7.Cell("A3").Value = "Q2";
        ws7.Cell("A4").Value = "Q3";
        ws7.Cell("A5").Value = "Q4";

        ws7.Cell("B2").Value = 800;
        ws7.Cell("B3").Value = 950;
        ws7.Cell("B4").Value = 1100;
        ws7.Cell("B5").Value = 1250;

        ws7.Cell("C2").Value = 12.5;
        ws7.Cell("C3").Value = 11.8;
        ws7.Cell("C4").Value = 13.2;
        ws7.Cell("C5").Value = 14.0;

        ws7.Columns("A", "C").AdjustToContents();

        // Primary: column chart for units; Secondary: line chart for average price
        var chart7 = ws7.Charts.Add(XLChartType.ColumnClustered);
        chart7.SetTitle("Units Sold & Avg Price");
        chart7.Series.Add("Units", "Combo!$B$2:$B$5", "Combo!$A$2:$A$5");
        chart7.SecondaryChartType = XLChartType.Line;
        chart7.SecondarySeries.Add("Avg Price", "Combo!$C$2:$C$5", "Combo!$A$2:$A$5");
        chart7.Position.SetColumn(0).SetRow(7);
        chart7.SecondPosition.SetColumn(10).SetRow(24);

        // --- Sheet 8: Radar chart ---
        var ws8 = wb.Worksheets.Add("Skills");

        ws8.Cell("A1").Value = "Skill";
        ws8.Cell("B1").Value = "Alice";
        ws8.Cell("C1").Value = "Bob";

        ws8.Cell("A2").Value = "C#";
        ws8.Cell("A3").Value = "SQL";
        ws8.Cell("A4").Value = "DevOps";
        ws8.Cell("A5").Value = "Testing";
        ws8.Cell("A6").Value = "Design";

        ws8.Cell("B2").Value = 9;
        ws8.Cell("B3").Value = 7;
        ws8.Cell("B4").Value = 6;
        ws8.Cell("B5").Value = 8;
        ws8.Cell("B6").Value = 4;

        ws8.Cell("C2").Value = 6;
        ws8.Cell("C3").Value = 9;
        ws8.Cell("C4").Value = 8;
        ws8.Cell("C5").Value = 5;
        ws8.Cell("C6").Value = 7;

        ws8.Columns("A", "C").AdjustToContents();

        var chart8 = ws8.Charts.Add(XLChartType.Radar);
        chart8.SetTitle("Skill Comparison");
        chart8.Series.Add("Alice", "Skills!$B$2:$B$6", "Skills!$A$2:$A$6");
        chart8.Series.Add("Bob", "Skills!$C$2:$C$6", "Skills!$A$2:$A$6");
        chart8.Position.SetColumn(0).SetRow(8);
        chart8.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 9: Scatter (XY) chart ---
        var ws9 = wb.Worksheets.Add("Scatter");

        ws9.Cell("A1").Value = "X";
        ws9.Cell("B1").Value = "Y";
        ws9.Cell("A2").Value = 1.0;
        ws9.Cell("A3").Value = 2.5;
        ws9.Cell("A4").Value = 4.0;
        ws9.Cell("A5").Value = 5.5;
        ws9.Cell("A6").Value = 7.0;
        ws9.Cell("B2").Value = 2.3;
        ws9.Cell("B3").Value = 3.1;
        ws9.Cell("B4").Value = 5.8;
        ws9.Cell("B5").Value = 6.2;
        ws9.Cell("B6").Value = 8.9;

        ws9.Columns("A", "B").AdjustToContents();

        var chart9 = ws9.Charts.Add(XLChartType.XYScatterMarkers);
        chart9.SetTitle("Scatter Plot");
        chart9.Series.Add("Measurements", "Scatter!$B$2:$B$6", "Scatter!$A$2:$A$6");
        chart9.Position.SetColumn(0).SetRow(8);
        chart9.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 10: Stock (High-Low-Close) chart ---
        var ws10 = wb.Worksheets.Add("Stock");

        ws10.Cell("A1").Value = "Date";
        ws10.Cell("B1").Value = "High";
        ws10.Cell("C1").Value = "Low";
        ws10.Cell("D1").Value = "Close";

        ws10.Cell("A2").Value = "Mon";
        ws10.Cell("A3").Value = "Tue";
        ws10.Cell("A4").Value = "Wed";
        ws10.Cell("A5").Value = "Thu";

        ws10.Cell("B2").Value = 105; ws10.Cell("C2").Value = 98;  ws10.Cell("D2").Value = 102;
        ws10.Cell("B3").Value = 108; ws10.Cell("C3").Value = 100; ws10.Cell("D3").Value = 104;
        ws10.Cell("B4").Value = 110; ws10.Cell("C4").Value = 99;  ws10.Cell("D4").Value = 107;
        ws10.Cell("B5").Value = 112; ws10.Cell("C5").Value = 103; ws10.Cell("D5").Value = 109;

        ws10.Columns("A", "D").AdjustToContents();

        var chart10 = ws10.Charts.Add(XLChartType.StockHighLowClose);
        chart10.SetTitle("Stock Prices");
        chart10.Series.Add("High", "Stock!$B$2:$B$5", "Stock!$A$2:$A$5");
        chart10.Series.Add("Low", "Stock!$C$2:$C$5", "Stock!$A$2:$A$5");
        chart10.Series.Add("Close", "Stock!$D$2:$D$5", "Stock!$A$2:$A$5");
        chart10.Position.SetColumn(0).SetRow(7);
        chart10.SecondPosition.SetColumn(9).SetRow(22);

        // --- Sheet 11: Surface chart ---
        var ws11 = wb.Worksheets.Add("Surface");

        ws11.Cell("A1").Value = "";
        ws11.Cell("B1").Value = "Col1";
        ws11.Cell("C1").Value = "Col2";
        ws11.Cell("D1").Value = "Col3";
        ws11.Cell("A2").Value = "Row1";
        ws11.Cell("A3").Value = "Row2";
        ws11.Cell("A4").Value = "Row3";

        ws11.Cell("B2").Value = 10; ws11.Cell("C2").Value = 20; ws11.Cell("D2").Value = 30;
        ws11.Cell("B3").Value = 25; ws11.Cell("C3").Value = 15; ws11.Cell("D3").Value = 35;
        ws11.Cell("B4").Value = 30; ws11.Cell("C4").Value = 40; ws11.Cell("D4").Value = 20;

        ws11.Columns("A", "D").AdjustToContents();

        var chart11 = ws11.Charts.Add(XLChartType.Surface);
        chart11.SetTitle("Surface Data");
        chart11.Series.Add("Col1", "Surface!$B$2:$B$4", "Surface!$A$2:$A$4");
        chart11.Series.Add("Col2", "Surface!$C$2:$C$4", "Surface!$A$2:$A$4");
        chart11.Series.Add("Col3", "Surface!$D$2:$D$4", "Surface!$A$2:$A$4");
        chart11.Position.SetColumn(0).SetRow(6);
        chart11.SecondPosition.SetColumn(9).SetRow(22);

        // --- Sheet 12: Waterfall chart (extended) ---
        var ws12 = wb.Worksheets.Add("Waterfall");

        ws12.Cell("A1").Value = "Category";
        ws12.Cell("B1").Value = "Amount";
        ws12.Cell("A2").Value = "Start";
        ws12.Cell("A3").Value = "Sales";
        ws12.Cell("A4").Value = "Returns";
        ws12.Cell("A5").Value = "Costs";
        ws12.Cell("A6").Value = "End";
        ws12.Cell("B2").Value = 1000;
        ws12.Cell("B3").Value = 500;
        ws12.Cell("B4").Value = -150;
        ws12.Cell("B5").Value = -300;
        ws12.Cell("B6").Value = 1050;

        ws12.Columns("A", "B").AdjustToContents();

        var chart12 = ws12.Charts.Add(XLChartType.Waterfall);
        chart12.SetTitle("Waterfall Analysis");
        chart12.Series.Add("Amount", "Waterfall!$B$2:$B$6", "Waterfall!$A$2:$A$6");
        chart12.Position.SetColumn(0).SetRow(8);
        chart12.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 13: Funnel chart (extended) ---
        var ws13 = wb.Worksheets.Add("Funnel");

        ws13.Cell("A1").Value = "Stage";
        ws13.Cell("B1").Value = "Count";
        ws13.Cell("A2").Value = "Visitors";
        ws13.Cell("A3").Value = "Leads";
        ws13.Cell("A4").Value = "Qualified";
        ws13.Cell("A5").Value = "Proposals";
        ws13.Cell("A6").Value = "Closed";
        ws13.Cell("B2").Value = 5000;
        ws13.Cell("B3").Value = 2500;
        ws13.Cell("B4").Value = 1200;
        ws13.Cell("B5").Value = 600;
        ws13.Cell("B6").Value = 200;

        ws13.Columns("A", "B").AdjustToContents();

        var chart13 = ws13.Charts.Add(XLChartType.Funnel);
        chart13.SetTitle("Sales Funnel");
        chart13.Series.Add("Count", "Funnel!$B$2:$B$6", "Funnel!$A$2:$A$6");
        chart13.Position.SetColumn(0).SetRow(8);
        chart13.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 14: Sunburst chart (extended, hierarchical data) ---
        var ws14 = wb.Worksheets.Add("Sunburst");

        // Sunburst requires hierarchical data: multiple category columns define the ring levels
        ws14.Cell("A1").Value = "Branch";
        ws14.Cell("B1").Value = "Category";
        ws14.Cell("C1").Value = "Item";
        ws14.Cell("D1").Value = "Value";

        ws14.Cell("A2").Value = "Food";  ws14.Cell("B2").Value = "Fruit";     ws14.Cell("C2").Value = "Apple";   ws14.Cell("D2").Value = 30;
        ws14.Cell("A3").Value = "Food";  ws14.Cell("B3").Value = "Fruit";     ws14.Cell("C3").Value = "Banana";  ws14.Cell("D3").Value = 25;
        ws14.Cell("A4").Value = "Food";  ws14.Cell("B4").Value = "Vegetable"; ws14.Cell("C4").Value = "Carrot";  ws14.Cell("D4").Value = 15;
        ws14.Cell("A5").Value = "Food";  ws14.Cell("B5").Value = "Vegetable"; ws14.Cell("C5").Value = "Peas";    ws14.Cell("D5").Value = 10;
        ws14.Cell("A6").Value = "Drink"; ws14.Cell("B6").Value = "Hot";       ws14.Cell("C6").Value = "Coffee";  ws14.Cell("D6").Value = 35;
        ws14.Cell("A7").Value = "Drink"; ws14.Cell("B7").Value = "Hot";       ws14.Cell("C7").Value = "Tea";     ws14.Cell("D7").Value = 20;
        ws14.Cell("A8").Value = "Drink"; ws14.Cell("B8").Value = "Cold";      ws14.Cell("C8").Value = "Juice";   ws14.Cell("D8").Value = 18;

        ws14.Columns("A", "D").AdjustToContents();

        // Category references span multiple columns (A:C) for hierarchy; dir="col" is set automatically
        var chart14 = ws14.Charts.Add(XLChartType.Sunburst);
        chart14.SetTitle("Food & Drink Breakdown");
        chart14.Series.Add("Value", "Sunburst!$D$2:$D$8", "Sunburst!$A$2:$C$8");
        chart14.Position.SetColumn(0).SetRow(10);
        chart14.SecondPosition.SetColumn(10).SetRow(28);

        // --- Sheet 15: Treemap chart (extended, hierarchical data) ---
        var ws15 = wb.Worksheets.Add("Treemap");

        // Treemap also requires hierarchical data
        ws15.Cell("A1").Value = "Region";
        ws15.Cell("B1").Value = "Country";
        ws15.Cell("C1").Value = "Revenue";

        ws15.Cell("A2").Value = "Americas"; ws15.Cell("B2").Value = "USA";       ws15.Cell("C2").Value = 400;
        ws15.Cell("A3").Value = "Americas"; ws15.Cell("B3").Value = "Canada";    ws15.Cell("C3").Value = 100;
        ws15.Cell("A4").Value = "Americas"; ws15.Cell("B4").Value = "Brazil";    ws15.Cell("C4").Value = 80;
        ws15.Cell("A5").Value = "Europe";   ws15.Cell("B5").Value = "UK";        ws15.Cell("C5").Value = 200;
        ws15.Cell("A6").Value = "Europe";   ws15.Cell("B6").Value = "Germany";   ws15.Cell("C6").Value = 150;
        ws15.Cell("A7").Value = "Asia";     ws15.Cell("B7").Value = "Japan";     ws15.Cell("C7").Value = 180;
        ws15.Cell("A8").Value = "Asia";     ws15.Cell("B8").Value = "Australia"; ws15.Cell("C8").Value = 100;

        ws15.Columns("A", "C").AdjustToContents();

        var chart15 = ws15.Charts.Add(XLChartType.Treemap);
        chart15.SetTitle("Revenue by Region");
        chart15.Series.Add("Revenue", "Treemap!$C$2:$C$8", "Treemap!$A$2:$B$8");
        chart15.Position.SetColumn(0).SetRow(10);
        chart15.SecondPosition.SetColumn(10).SetRow(28);

        // --- Sheet 16: Box & Whisker chart (extended) ---
        var ws16 = wb.Worksheets.Add("BoxWhisker");

        ws16.Cell("A1").Value = "Group";
        ws16.Cell("B1").Value = "Value";
        ws16.Cell("A2").Value = "A"; ws16.Cell("B2").Value = 12;
        ws16.Cell("A3").Value = "A"; ws16.Cell("B3").Value = 15;
        ws16.Cell("A4").Value = "A"; ws16.Cell("B4").Value = 18;
        ws16.Cell("A5").Value = "A"; ws16.Cell("B5").Value = 22;
        ws16.Cell("A6").Value = "B"; ws16.Cell("B6").Value = 8;
        ws16.Cell("A7").Value = "B"; ws16.Cell("B7").Value = 14;
        ws16.Cell("A8").Value = "B"; ws16.Cell("B8").Value = 20;
        ws16.Cell("A9").Value = "B"; ws16.Cell("B9").Value = 25;

        ws16.Columns("A", "B").AdjustToContents();

        var chart16 = ws16.Charts.Add(XLChartType.BoxWhisker);
        chart16.SetTitle("Distribution by Group");
        chart16.Series.Add("Value", "BoxWhisker!$B$2:$B$9", "BoxWhisker!$A$2:$A$9");
        chart16.Position.SetColumn(0).SetRow(11);
        chart16.SecondPosition.SetColumn(9).SetRow(26);

        // --- Sheet 17: Area chart ---
        var ws17 = wb.Worksheets.Add("Area");

        ws17.Cell("A1").Value = "Month";
        ws17.Cell("B1").Value = "Product A";
        ws17.Cell("C1").Value = "Product B";

        ws17.Cell("A2").Value = "Jan"; ws17.Cell("B2").Value = 30; ws17.Cell("C2").Value = 20;
        ws17.Cell("A3").Value = "Feb"; ws17.Cell("B3").Value = 35; ws17.Cell("C3").Value = 25;
        ws17.Cell("A4").Value = "Mar"; ws17.Cell("B4").Value = 28; ws17.Cell("C4").Value = 30;
        ws17.Cell("A5").Value = "Apr"; ws17.Cell("B5").Value = 40; ws17.Cell("C5").Value = 28;
        ws17.Cell("A6").Value = "May"; ws17.Cell("B6").Value = 45; ws17.Cell("C6").Value = 35;

        ws17.Columns("A", "C").AdjustToContents();

        var chart17 = ws17.Charts.Add(XLChartType.AreaStacked);
        chart17.SetTitle("Stacked Area");
        chart17.Series.Add("Product A", "Area!$B$2:$B$6", "Area!$A$2:$A$6");
        chart17.Series.Add("Product B", "Area!$C$2:$C$6", "Area!$A$2:$A$6");
        chart17.Position.SetColumn(0).SetRow(8);
        chart17.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 18: Doughnut chart ---
        var ws18 = wb.Worksheets.Add("Doughnut");

        ws18.Cell("A1").Value = "Source";
        ws18.Cell("B1").Value = "Traffic";

        ws18.Cell("A2").Value = "Organic"; ws18.Cell("B2").Value = 45;
        ws18.Cell("A3").Value = "Direct";  ws18.Cell("B3").Value = 25;
        ws18.Cell("A4").Value = "Social";  ws18.Cell("B4").Value = 20;
        ws18.Cell("A5").Value = "Referral"; ws18.Cell("B5").Value = 10;

        ws18.Columns("A", "B").AdjustToContents();

        var chart18 = ws18.Charts.Add(XLChartType.Doughnut);
        chart18.SetTitle("Traffic Sources");
        chart18.Series.Add("Traffic", "Doughnut!$B$2:$B$5", "Doughnut!$A$2:$A$5");
        chart18.Position.SetColumn(0).SetRow(7);
        chart18.SecondPosition.SetColumn(9).SetRow(22);

        // --- Sheet 19: Bubble chart ---
        var ws19 = wb.Worksheets.Add("Bubble");

        ws19.Cell("A1").Value = "X (Revenue)";
        ws19.Cell("B1").Value = "Y (Profit)";

        ws19.Cell("A2").Value = 100; ws19.Cell("B2").Value = 15;
        ws19.Cell("A3").Value = 200; ws19.Cell("B3").Value = 30;
        ws19.Cell("A4").Value = 150; ws19.Cell("B4").Value = 10;
        ws19.Cell("A5").Value = 300; ws19.Cell("B5").Value = 50;
        ws19.Cell("A6").Value = 250; ws19.Cell("B6").Value = 35;

        ws19.Columns("A", "B").AdjustToContents();

        var chart19 = ws19.Charts.Add(XLChartType.Bubble);
        chart19.SetTitle("Revenue vs Profit");
        chart19.Series.Add("Products", "Bubble!$B$2:$B$6", "Bubble!$A$2:$A$6");
        chart19.Position.SetColumn(0).SetRow(8);
        chart19.SecondPosition.SetColumn(9).SetRow(24);

        // --- Sheet 20: 3D shape charts (Cone, Cylinder, Pyramid) ---
        var ws20 = wb.Worksheets.Add("3D Shapes");

        ws20.Cell("A1").Value = "Item";
        ws20.Cell("B1").Value = "Value";
        ws20.Cell("A2").Value = "Alpha"; ws20.Cell("B2").Value = 45;
        ws20.Cell("A3").Value = "Beta";  ws20.Cell("B3").Value = 30;
        ws20.Cell("A4").Value = "Gamma"; ws20.Cell("B4").Value = 55;
        ws20.Cell("A5").Value = "Delta"; ws20.Cell("B5").Value = 20;

        ws20.Columns("A", "B").AdjustToContents();

        // Cone chart
        var chart20a = ws20.Charts.Add(XLChartType.ConeClustered);
        chart20a.SetTitle("Cone");
        chart20a.Series.Add("Value", "'3D Shapes'!$B$2:$B$5", "'3D Shapes'!$A$2:$A$5");
        chart20a.Position.SetColumn(0).SetRow(7);
        chart20a.SecondPosition.SetColumn(5).SetRow(20);

        // Cylinder chart
        var chart20b = ws20.Charts.Add(XLChartType.CylinderClustered);
        chart20b.SetTitle("Cylinder");
        chart20b.Series.Add("Value", "'3D Shapes'!$B$2:$B$5", "'3D Shapes'!$A$2:$A$5");
        chart20b.Position.SetColumn(6).SetRow(7);
        chart20b.SecondPosition.SetColumn(11).SetRow(20);

        // Pyramid chart
        var chart20c = ws20.Charts.Add(XLChartType.PyramidClustered);
        chart20c.SetTitle("Pyramid");
        chart20c.Series.Add("Value", "'3D Shapes'!$B$2:$B$5", "'3D Shapes'!$A$2:$A$5");
        chart20c.Position.SetColumn(0).SetRow(22);
        chart20c.SecondPosition.SetColumn(5).SetRow(35);

        wb.SaveAs(filePath);
    }
}
