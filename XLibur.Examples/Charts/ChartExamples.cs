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

        wb.SaveAs(filePath);
    }
}
