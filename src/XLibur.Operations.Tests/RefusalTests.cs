using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class RefusalTests
{
    public static IEnumerable<Func<(IOperation Operation, object Request, string Code)>> Cases()
    {
        yield return () => (new FindOperation(), new FindRequest { Path = "sales.xlsx", Text = "a", MaxResults = 0 }, "bad_max_results");
        yield return () => (new ReadRangeOperation(), new ReadRangeRequest { Path = "sales.xlsx", MaxCells = 0 }, "bad_max_cells");
        yield return () => (new ReadRangeOperation(), new ReadRangeRequest { Path = "sales.xlsx", Range = "Sales!A1:D1", MaxCells = 2 }, "range_too_wide");
        yield return () => (new ReadRangeOperation(), new ReadRangeRequest { Path = "sales.xlsx", Range = "Sales!A1:" }, "bad_address");
        yield return () => (new EditOperation(), new EditRequest { Path = "sales.xlsx", Edits = [] }, "no_edits");
        yield return () => (new EditOperation(), new EditRequest { Path = "sales.xlsx", Edits = [new CellEdit { Cell = "A1", Value = 1, Formula = "2" }] }, "bad_edit");
        yield return () => (new EditOperation(), new EditRequest { Path = "sales.xlsx", Edits = [new CellEdit { Cell = "A1:B2", Value = 1 }] }, "not_a_cell");
        yield return () => (new EditOperation(), new EditRequest { Path = "sales.xlsx", OutputPath = "out.xlsm", Edits = [new CellEdit { Cell = "A1", Value = 1 }] }, "extension_mismatch");
        yield return () => (new CheckFormulaOperation(), new CheckFormulaRequest { Formula = "=" }, "formula_empty");
        yield return () => (new CheckFormulaOperation(), new CheckFormulaRequest { Formula = "1+1", Sheet = "Sales" }, "path_required");
        yield return () => (new RenderReportOperation(), new RenderReportRequest { Template = "sales.xlsx", Data = "data.json", Output = "report.csv" }, "not_a_workbook");
        yield return () => (new RenderReportOperation(), new RenderReportRequest { Template = "sales.xlsx", Data = "broken.json", Output = "report.xlsx" }, "bad_data");
        yield return () => (new ValidateOperation(), new ValidateRequest { Path = "table.csv" }, "not_a_workbook");
        yield return () => (new RecalcOperation(), new RecalcRequest { Path = "table.csv" }, "not_a_workbook");
        yield return () => (new ConvertOperation(), new ConvertRequest { Input = "table.csv", Output = "out.xlsm" }, "table_to_xlsm");
        yield return () => (new DiffOperation(), new DiffRequest { Left = "sales.xlsx", Right = "sales.xlsx", MaxChanges = -1 }, "bad_max_changes");
    }

    [Test]
    [MethodDataSource(nameof(Cases))]
    public async Task A_request_the_operation_cannot_carry_out_is_refused_with_its_code(IOperation operation, object request, string code)
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        root.File("table.csv", "a,b\r\n1,2\r\n");
        root.File("data.json", "{}");
        root.File("broken.json", "{");
        var before = Directory.GetFiles(root.Path).Order().ToList();

        var outcome = Fixtures.Run(Fixtures.Context(root), operation, request).Outcome;

        await Assert.That(outcome.Error?.Code).IsEqualTo(code);
        await Assert.That(outcome.ExitCode).IsEqualTo(2);
        await Assert.That(Directory.GetFiles(root.Path).Order().ToList()).IsEquivalentTo(before);
    }

    [Test]
    public async Task An_exception_that_is_not_an_operation_error_is_reported_as_a_defect()
    {
        using var root = new TempDirectory();

        var outcome = Fixtures.Run(Fixtures.Context(root), new Throwing(), new InspectRequest { Path = "x.xlsx" }).Outcome;

        await Assert.That(outcome.Error!.Kind).IsEqualTo(ErrorKind.Defect);
        await Assert.That(outcome.Error.Message).IsEqualTo("System.InvalidOperationException: broken");
    }

    [Test]
    public async Task A_result_longer_than_the_host_allows_is_replaced_by_an_error()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        var context = Fixtures.Context(root, limits: Limits.Default with { MaxResultChars = 20 });

        var (outcome, text) = Fixtures.Run(context, new InspectOperation(), new InspectRequest { Path = "sales.xlsx" });

        await Assert.That(outcome.Error!.Code).IsEqualTo("result_too_large");
        await Assert.That(text).StartsWith("error[result_too_large]");
    }

    private sealed class Throwing : Operation<InspectRequest, InspectResult>
    {
        public override string Name => "throwing";

        public override string Description => "Throws.";

        public override OperationEffect Effect => OperationEffect.ReadOnly;

        public override InspectResult Execute(InspectRequest request, OperationContext context) => throw new InvalidOperationException("broken");
    }
}
