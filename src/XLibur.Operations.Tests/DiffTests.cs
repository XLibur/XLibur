using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class DiffTests
{
    private static readonly DiffOperation Operation = new();

    [Test]
    public async Task A_workbook_is_identical_to_itself_and_to_its_XLibur_round_trip()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        using (var wb = new XLWorkbook(path))
            wb.SaveAs(root.Combine("roundtrip.xlsx"));

        var self = Fixtures.Run(Fixtures.Context(root), Operation, new DiffRequest { Left = "sales.xlsx", Right = "sales.xlsx" }).Outcome;
        var trip = Fixtures.Run(Fixtures.Context(root), Operation, new DiffRequest { Left = "sales.xlsx", Right = "roundtrip.xlsx" }).Outcome;

        await Assert.That(self.ExitCode).IsEqualTo(0);
        await Assert.That(trip.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task Value_formula_style_sheet_name_and_table_changes_are_all_reported()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        using (var wb = new XLWorkbook(path))
        {
            var ws = wb.Worksheet("Sales");
            ws.Cell("B2").Value = 4;
            ws.Cell("D5").FormulaA1 = "SUM(D2:D3)";
            ws.Cell("A1").Style.Font.Bold = true;
            wb.AddWorksheet("Extra");
            wb.DefinedNames.Add("Rate", "Sales!$C$2");
            ws.Table("SalesTable").Resize(ws.Range("A1:D5"));
            wb.SaveAs(root.Combine("changed.xlsx"), new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root), Operation, new DiffRequest { Left = "sales.xlsx", Right = "changed.xlsx" });

        var result = (DiffResult)outcome.Result!;
        await Assert.That(outcome.ExitCode).IsEqualTo(1);
        await Assert.That(result.SheetsAdded).IsEquivalentTo(["Extra"]);
        await Assert.That(result.CellChanges).Contains(new CellChange("Sales!B2", CellChangeKind.Value, "3", "4"));
        await Assert.That(result.CellChanges).Contains(new CellChange("Sales!D5", CellChangeKind.Formula, "=SUM(D2:D4)", "=SUM(D2:D3)"));
        await Assert.That(result.CellChanges.Any(c => c.Cell == "Sales!A1" && c.Kind == CellChangeKind.Style && c.Left.Contains("font"))).IsTrue();
        await Assert.That(result.DefinedNames).Contains(new ItemChange("Rate", ItemChangeKind.Added, null, "Sales!$C$2"));
        await Assert.That(result.Tables).Contains(new ItemChange("SalesTable", ItemChangeKind.Changed, "Sales!A1:D4", "Sales!A1:D5"));
        await Assert.That(text).Contains("value Sales!B2\t3 -> 4");
    }

    [Test]
    public async Task Ignored_sheets_and_ranges_and_styles_are_left_out()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        using (var wb = new XLWorkbook(path))
        {
            var ws = wb.Worksheet("Sales");
            ws.Cell("B2").Value = 4;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Yellow;
            wb.SaveAs(root.Combine("changed.xlsx"), new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        var context = Fixtures.Context(root);
        var ranges = Fixtures.Run(context, Operation, new DiffRequest
        {
            Left = "sales.xlsx", Right = "changed.xlsx", IgnoreRanges = ["Sales!A1:B5", "Sales!D1:D5"],
        }).Outcome;
        var sheets = Fixtures.Run(context, Operation, new DiffRequest
        {
            Left = "sales.xlsx", Right = "changed.xlsx", IgnoreSheets = ["sales"],
        }).Outcome;

        await Assert.That(ranges.ExitCode).IsEqualTo(0);
        await Assert.That(sheets.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task An_ignored_sheet_leaves_out_its_tables_and_sheet_scoped_names()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        using (var wb = new XLWorkbook(path))
        {
            var ws = wb.Worksheet("Sales");
            ws.Table("SalesTable").Resize(ws.Range("A1:D5"));
            ws.DefinedNames.Add("Local", "Sales!$A$1");
            wb.SaveAs(root.Combine("changed.xlsx"), new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        var context = Fixtures.Context(root);
        var all = Fixtures.Run(context, Operation, new DiffRequest { Left = "sales.xlsx", Right = "changed.xlsx" }).Outcome;
        var ignored = Fixtures.Run(context, Operation, new DiffRequest
        {
            Left = "sales.xlsx", Right = "changed.xlsx", IgnoreSheets = ["Sales"],
        }).Outcome;

        var result = (DiffResult)all.Result!;
        await Assert.That(result.Tables.Count).IsEqualTo(1);
        await Assert.That(result.DefinedNames.Select(n => n.Name)).Contains("Sales!Local");
        await Assert.That(ignored.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task Recalculated_mode_compares_what_the_engine_computes_not_what_the_file_cached()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "a.xlsx", wb => wb.AddWorksheet("S").Cell("A1").FormulaA1 = "1+1");
        Fixtures.Build(root, "b.xlsx", wb => wb.AddWorksheet("S").Cell("A1").FormulaA1 = "1+1");
        using (var wb = new XLWorkbook(root.Combine("b.xlsx")))
        {
            wb.Worksheet("S").Cell("A1").FormulaA1 = "1+1";
            wb.SaveAs(root.Combine("b.xlsx"), new SaveOptions { EvaluateFormulasBeforeSaving = false });
        }

        var context = Fixtures.Context(root);
        var cached = Fixtures.Run(context, Operation, new DiffRequest { Left = "a.xlsx", Right = "b.xlsx" }).Outcome;
        var recalculated = Fixtures.Run(context, Operation, new DiffRequest { Left = "a.xlsx", Right = "b.xlsx", Values = ValueSource.Recalculated }).Outcome;

        await Assert.That(cached.ExitCode).IsEqualTo(1);
        await Assert.That(recalculated.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task An_ignored_range_must_name_its_sheet()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var outcome = Fixtures.Run(Fixtures.Context(root), Operation,
            new DiffRequest { Left = "sales.xlsx", Right = "sales.xlsx", IgnoreRanges = ["A1:B2"] }).Outcome;

        await Assert.That(outcome.Error!.Code).IsEqualTo("bad_ignore_range");
        await Assert.That(outcome.ExitCode).IsEqualTo(2);
    }
}
