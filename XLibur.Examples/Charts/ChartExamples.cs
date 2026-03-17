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

        wb.SaveAs(filePath);
    }
}
