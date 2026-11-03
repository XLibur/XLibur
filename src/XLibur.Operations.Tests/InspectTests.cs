using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class InspectTests
{
    private static readonly byte[] Png = System.Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    private static string Rich(TempDirectory root) => Fixtures.Build(root, "rich.xlsx", wb =>
    {
        var data = wb.AddWorksheet("Data");
        data.Cell("A1").Value = "Item";
        data.Cell("B1").Value = "Units";
        data.Cell("A2").Value = "Apple";
        data.Cell("B2").Value = 3;
        data.Cell("A3").Value = "Pear";
        data.Cell("B3").Value = 5;
        data.Cell("B4").FormulaA1 = "SUM(B2:B3)";
        data.Range("A1:B3").CreateTable("Fruit");
        var chart = data.Charts.Add(XLChartType.ColumnClustered);
        chart.Title = "Units";
        chart.Series.Add("Units", "Data!$B$2:$B$3", "Data!$A$2:$A$3");
        data.AddPicture(new MemoryStream(Png), "Logo");

        var report = wb.AddWorksheet("Report");
        report.Cell("A1").Value = "By item";
        report.DefinedNames.Add("Heading", "Report!$A$1");
        wb.AddWorksheet("Hidden").Visibility = XLWorksheetVisibility.Hidden;
        wb.AddWorksheet("Secret").Visibility = XLWorksheetVisibility.VeryHidden;
        wb.DefinedNames.Add("Total", "Data!$B$4");
    });

    [Test]
    public async Task Inspect_lists_each_sheet_with_its_counts_objects_and_visibility_then_the_names()
    {
        using var root = new TempDirectory();
        var path = Rich(root);

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root), new InspectOperation(), new InspectRequest { Path = "rich.xlsx" });

        await Assert.That(outcome.Error).IsNull();
        await Assert.That(text).IsEqualTo(
            path + "\n" +
            "1 Data used=A1:B4 cells=7 formulas=1 tables=Fruit(A1:B3) charts=ColumnClustered:Units pictures=1\n" +
            "2 Report used=A1 cells=1 formulas=0\n" +
            "3 Hidden (hidden) used=empty cells=0 formulas=0\n" +
            "4 Secret (veryHidden) used=empty cells=0 formulas=0\n" +
            "name Total = Data!$B$4\n" +
            "name Report!Heading = Report!$A$1\n");
    }

    [Test]
    public async Task Inspect_as_JSON_carries_the_same_summary_as_data()
    {
        using var root = new TempDirectory();
        Rich(root);

        var result = (InspectResult)Fixtures.Run(Fixtures.Context(root), new InspectOperation(), new InspectRequest { Path = "rich.xlsx" },
            ResultFormat.Json).Outcome.Result!;

        var data = result.Sheets[0];
        await Assert.That(data.Tables).IsEquivalentTo([new TableSummary("Fruit", "A1:B3")]);
        await Assert.That(data.Charts).IsEquivalentTo([new ChartSummary("ColumnClustered", "Units", 1)]);
        await Assert.That(result.Sheets.Select(s => s.Visibility)).IsEquivalentTo(["visible", "visible", "hidden", "veryHidden"]);
        await Assert.That(result.DefinedNames).IsEquivalentTo([
            new NameSummary("Total", null, "Data!$B$4"),
            new NameSummary("Heading", "Report", "Report!$A$1"),
        ]);
    }

    [Test]
    public async Task Inspect_names_a_pivot_table_by_its_target_cell()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "pivot.xlsx", wb =>
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = "Item";
            data.Cell("B1").Value = "Units";
            data.Cell("A2").Value = "Apple";
            data.Cell("B2").Value = 3;
            var pivot = wb.AddWorksheet("Pivot").PivotTables.Add("ByItem", wb.Worksheet("Pivot").Cell("C3"), data.Range("A1:B2"));
            pivot.RowLabels.Add("Item");
            pivot.Values.Add("Units");
        });

        var text = Fixtures.Run(Fixtures.Context(root), new InspectOperation(), new InspectRequest { Path = "pivot.xlsx" }).Text;

        await Assert.That(text).Contains(" pivots=ByItem@C3\n");
    }
}
