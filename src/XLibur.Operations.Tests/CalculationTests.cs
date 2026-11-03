using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class CalculationTests
{
    [Test]
    public async Task Recalc_sorts_failures_into_error_values_cycles_and_unsupported_functions()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "mixed.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").FormulaA1 = "1/0";
            ws.Cell("A2").FormulaA1 = "_xlfn.CEILING.PRECISE(2.5,1)";
            ws.Cell("A3").FormulaA1 = "A4+1";
            ws.Cell("A4").Value = 1;
            ws.Cell("A5").FormulaA1 = "SUM(1,2)";
        });
        using (var wb = new XLWorkbook(root.Combine("mixed.xlsx")))
        {
            wb.Worksheet("S").Cell("A4").FormulaA1 = "A3";
            wb.Save();
        }

        var result = (RecalcResult)Fixtures.Run(Fixtures.Context(root), new RecalcOperation(), new RecalcRequest { Path = "mixed.xlsx" }).Outcome.Result!;

        await Assert.That(result.Formulas).IsEqualTo(5);
        await Assert.That(result.ErrorValues.Cells.Select(c => c.Cell)).IsEquivalentTo(["S!A1"]);
        await Assert.That(result.Unsupported.Cells.Single().Detail).Contains("CEILING.PRECISE");
        await Assert.That(result.Circular.Cells.Select(c => c.Cell)).IsEquivalentTo(["S!A3", "S!A4"]);
        await Assert.That(result.Defects.Total).IsEqualTo(0);
    }

    [Test]
    public async Task Recalc_and_validate_with_max_issues_0_report_only_the_counts()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "errors.xlsx", wb => wb.AddWorksheet("S").Cell("A1").FormulaA1 = "1/0");
        var context = Fixtures.Context(root);

        var (recalc, recalcText) = Fixtures.Run(context, new RecalcOperation(), new RecalcRequest { Path = "errors.xlsx", MaxIssues = 0 });
        var (validate, validateText) = Fixtures.Run(context, new ValidateOperation(), new ValidateRequest { Path = "errors.xlsx", MaxIssues = 0 });
        var negative = Fixtures.Run(context, new RecalcOperation(), new RecalcRequest { Path = "errors.xlsx", MaxIssues = -1 }).Outcome;

        await Assert.That(recalc.Error).IsNull();
        await Assert.That(recalcText).Contains("1 more errorValue\n");
        await Assert.That(validate.Error).IsNull();
        await Assert.That(validateText).Contains("1 more errorValue\n");
        await Assert.That(negative.Error!.Code).IsEqualTo("bad_max_issues");
    }

    [Test]
    public async Task Check_formula_evaluates_against_the_workbook_without_writing()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        var before = await File.ReadAllBytesAsync(path);

        var result = (CheckFormulaResult)Fixtures.Run(Fixtures.Context(root, allowWrite: false), new CheckFormulaOperation(),
            new CheckFormulaRequest { Formula = "=SUMPRODUCT(B2:B4,C2:C4)+GrandTotal", Path = "sales.xlsx" }).Outcome.Result!;

        await Assert.That(result.Valid).IsTrue();
        await Assert.That(result.Value).IsEqualTo((XLCellValue)23d);
        await Assert.That(await File.ReadAllBytesAsync(path)).IsEquivalentTo(before);
    }

    [Test]
    public async Task Check_formula_reports_where_the_parser_stopped()
    {
        var (outcome, text) = Fixtures.Run(new OperationContext(Paths.PathPolicy.Unrestricted()), new CheckFormulaOperation(),
            new CheckFormulaRequest { Formula = "=SUM(1,2" });

        var result = (CheckFormulaResult)outcome.Result!;
        await Assert.That(result.Valid).IsFalse();
        await Assert.That(result.Position).IsEqualTo(7);
        await Assert.That(outcome.ExitCode).IsEqualTo(1);
        await Assert.That(text).StartsWith("refused: ");
    }

    [Test]
    public async Task Check_formula_without_a_workbook_evaluates_a_constant_formula()
    {
        var result = (CheckFormulaResult)Fixtures.Run(new OperationContext(Paths.PathPolicy.Unrestricted()), new CheckFormulaOperation(),
            new CheckFormulaRequest { Formula = "ROUND(PI(),2)" }).Outcome.Result!;

        await Assert.That(result.Value).IsEqualTo((XLCellValue)3.14);
    }

    [Test]
    public async Task Find_matches_text_regex_and_errors()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "find.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = "Invoice 2026-001";
            ws.Cell("A2").Value = "invoice draft";
            ws.Cell("B1").FormulaA1 = "1/0";
            ws.Cell("B2").FormulaA1 = "VLOOKUP(\"x\",A1:A2,1,FALSE)";
        });
        var context = Fixtures.Context(root);
        var find = new FindOperation();

        var text = (FindResult)Fixtures.Run(context, find, new FindRequest { Path = "find.xlsx", Text = "invoice" }).Outcome.Result!;
        var cased = (FindResult)Fixtures.Run(context, find, new FindRequest { Path = "find.xlsx", Text = "invoice", MatchCase = true }).Outcome.Result!;
        var regex = (FindResult)Fixtures.Run(context, find, new FindRequest { Path = "find.xlsx", Regex = @"\d{4}-\d{3}" }).Outcome.Result!;
        var errors = (FindResult)Fixtures.Run(context, find, new FindRequest { Path = "find.xlsx", Errors = true }).Outcome.Result!;
        var formulas = (FindResult)Fixtures.Run(context, find, new FindRequest { Path = "find.xlsx", Text = "VLOOKUP", InFormulas = true }).Outcome.Result!;

        await Assert.That(text.Total).IsEqualTo(2);
        await Assert.That(cased.Matches.Single().Cell).IsEqualTo("S!A2");
        await Assert.That(regex.Matches.Single().Cell).IsEqualTo("S!A1");
        await Assert.That(errors.Matches.Select(m => m.Value)).IsEquivalentTo(["#DIV/0!", "#N/A"]);
        await Assert.That(formulas.Matches.Single().Cell).IsEqualTo("S!B2");
    }

    [Test]
    public async Task Find_needs_exactly_one_criterion_and_a_valid_regex()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        var context = Fixtures.Context(root);

        var none = Fixtures.Run(context, new FindOperation(), new FindRequest { Path = "sales.xlsx" }).Outcome;
        var bad = Fixtures.Run(context, new FindOperation(), new FindRequest { Path = "sales.xlsx", Regex = "(" }).Outcome;

        await Assert.That(none.Error!.Code).IsEqualTo("bad_criteria");
        await Assert.That(bad.Error!.Code).IsEqualTo("bad_regex");
    }

    [Test]
    public async Task Validate_passes_a_clean_workbook_and_its_round_trip()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root, allowWrite: false), new ValidateOperation(), new ValidateRequest { Path = "sales.xlsx" });

        var result = (ValidateResult)outcome.Result!;
        await Assert.That(result.Ok).IsTrue();
        await Assert.That(result.RoundTrip!.Differences!.Identical).IsTrue();
        await Assert.That(outcome.ExitCode).IsEqualTo(0);
        await Assert.That(text).Contains("round trip: identical");
    }

    [Test]
    public async Task Validate_reports_a_file_that_does_not_load_as_a_failed_check()
    {
        using var root = new TempDirectory();
        root.File("broken.xlsx", "this is not a zip");

        var (outcome, _) = Fixtures.Run(Fixtures.Context(root), new ValidateOperation(), new ValidateRequest { Path = "broken.xlsx" });

        var result = (ValidateResult)outcome.Result!;
        await Assert.That(result.Loaded).IsFalse();
        await Assert.That(result.LoadError).Contains("not_a_workbook");
        await Assert.That(outcome.ExitCode).IsEqualTo(1);
    }
}
