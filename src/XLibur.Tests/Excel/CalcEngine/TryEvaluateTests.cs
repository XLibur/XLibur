using System;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using static XLibur.Tests.Excel.CalcEngine.EvaluationOutcomeTests;

namespace XLibur.Tests.Excel.CalcEngine;

public class TryEvaluateTests
{
    private const string SheetName = "Sheet1";
    private const string At = "A6";

    [Test]
    [Arguments(Kind.Cycle, XLFormulaIssueKind.CircularReference)]
    [Arguments(Kind.Unsupported, XLFormulaIssueKind.UnsupportedFeature)]
    [Arguments(Kind.Refused, XLFormulaIssueKind.RefusedFormula)]
    public async Task A_cell_reports_each_expected_failure_without_throwing(Kind failure, XLFormulaIssueKind expected)
    {
        using var wb = WorkbookWith(failure);

        var result = wb.Worksheet(SheetName).Cell(At).TryEvaluate();

        await AssertFailure(result, expected);
    }

    [Test]
    [Arguments(Kind.Cycle, XLFormulaIssueKind.CircularReference)]
    [Arguments(Kind.Unsupported, XLFormulaIssueKind.UnsupportedFeature)]
    [Arguments(Kind.Refused, XLFormulaIssueKind.RefusedFormula)]
    public async Task A_worksheet_expression_reports_each_expected_failure_without_throwing(
        Kind failure, XLFormulaIssueKind expected)
    {
        using var wb = WorkbookWith(failure);

        var result = wb.Worksheet(SheetName).TryEvaluate(Expression(failure, qualified: false));

        await AssertFailure(result, expected);
    }

    [Test]
    [Arguments(Kind.Cycle, XLFormulaIssueKind.CircularReference)]
    [Arguments(Kind.Unsupported, XLFormulaIssueKind.UnsupportedFeature)]
    [Arguments(Kind.Refused, XLFormulaIssueKind.RefusedFormula)]
    public async Task A_workbook_expression_reports_each_expected_failure_without_throwing(
        Kind failure, XLFormulaIssueKind expected)
    {
        using var wb = WorkbookWith(failure);

        var result = wb.TryEvaluate(Expression(failure, qualified: true));

        await AssertFailure(result, expected);
    }

    [Test]
    public async Task A_refused_formula_carries_the_exception_the_throwing_path_throws()
    {
        using var wb = WorkbookWith(Kind.Refused);
        var cell = wb.Worksheet(SheetName).Cell(At);

        var thrown = await Assert.That(() => _ = cell.Value).Throws<ExpressionParseException>();
        var issue = cell.TryEvaluate().Issue;

        await Assert.That(issue!.Message).IsEqualTo(thrown!.Message);
        await Assert.That(issue.Inner).IsTypeOf<ExpressionParseException>();
        await Assert.That(issue.Position).IsEqualTo(2);
    }

    [Test]
    public async Task A_circular_reference_carries_the_exception_the_throwing_path_throws()
    {
        using var wb = WorkbookWith(Kind.Cycle);
        var ws = wb.Worksheet(SheetName);

        var issue = ws.Cell(At).TryEvaluate().Issue;

        await Assert.That(issue!.Inner).IsTypeOf<XLCircularReferenceException>();
        await Assert.That(issue.Message).IsEqualTo("Formula in a cell '$Sheet1'!$A6 is part of a cycle.");
    }

    [Test]
    [Arguments("1+2", 3.0)]
    [Arguments("SUM(1,2,3)", 6.0)]
    public async Task A_successful_evaluation_is_the_value_the_throwing_path_gives(string formula, double expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = formula;

        var viaCell = ws.Cell("A1").TryEvaluate();
        var viaSheet = ws.TryEvaluate(formula);
        var viaBook = wb.TryEvaluate(formula);

        foreach (var result in new[] { viaCell, viaSheet, viaBook })
        {
            await Assert.That(result.IsSuccess).IsTrue();
            await Assert.That(result.Issue).IsNull();
            await Assert.That(result.Value).IsEqualTo(expected);
        }

        await Assert.That(viaCell.Value).IsEqualTo(ws.Cell("A1").Value);
        await Assert.That(viaSheet.Value).IsEqualTo(ws.Evaluate(formula));
        await Assert.That(viaBook.Value).IsEqualTo(wb.Evaluate(formula));
    }

    [Test]
    [Arguments("1/0", XLError.DivisionByZero)]
    [Arguments("NA()", XLError.NoValueAvailable)]
    public async Task An_Excel_error_value_is_a_successful_result(string formula, XLError error)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = formula;

        foreach (var result in new[] { ws.Cell("A1").TryEvaluate(), ws.TryEvaluate(formula), wb.TryEvaluate(formula) })
        {
            await Assert.That(result.IsSuccess).IsTrue();
            await Assert.That(result.Value).IsEqualTo(error);
        }
    }

    [Test]
    public async Task A_cell_without_a_formula_is_its_value()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet(SheetName).Cell("A1");
        cell.Value = "text";

        await Assert.That(cell.TryEvaluate().Value).IsEqualTo("text");
    }

    [Test]
    public async Task A_defect_still_propagates_from_each_overload()
    {
        using var wb = WorkbookWithDefectFunction();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell(At).FormulaA1 = "XLIBURDEFECT()";

        await Assert.That(() => ws.Cell(At).TryEvaluate()).Throws<NullReferenceException>();
        await Assert.That(() => ws.TryEvaluate("XLIBURDEFECT()")).Throws<NullReferenceException>();
        await Assert.That(() => wb.TryEvaluate("XLIBURDEFECT()")).Throws<NullReferenceException>();
    }

    [Test]
    public async Task A_missing_context_is_a_mistake_in_the_call_and_still_throws()
    {
        using var wb = WorkbookWith(Kind.NoContext);

        await Assert.That(() => wb.TryEvaluate("ROW()")).Throws<XLNoWorksheetContextException>();
        await Assert.That(() => wb.Worksheet(SheetName).TryEvaluate("ROW()")).Throws<XLNoWorksheetContextException>();
    }

    /// <summary>
    /// ADR 0002 for the public API: a failed call changes nothing, so the cell is still dirty with no
    /// cached value, and a second call finds the same problem.
    /// </summary>
    [Test]
    [Arguments(Kind.Cycle)]
    [Arguments(Kind.Unsupported)]
    [Arguments(Kind.Refused)]
    public async Task A_failed_call_leaves_the_cell_dirty_and_a_second_call_gives_the_same_issue(Kind failure)
    {
        using var wb = WorkbookWith(failure);
        var cell = wb.Worksheet(SheetName).Cell(At);

        var first = cell.TryEvaluate();

        await Assert.That(cell.NeedsRecalculation).IsTrue();
        await Assert.That(cell.CachedValue.IsBlank).IsTrue();

        var second = cell.TryEvaluate();

        await Assert.That(second.IsSuccess).IsFalse();
        await Assert.That(second.Issue!.Kind).IsEqualTo(first.Issue!.Kind);
        await Assert.That(second.Issue.Message).IsEqualTo(first.Issue.Message);
        await Assert.That(cell.NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task Value_of_a_failed_result_throws_instead_of_returning_a_blank()
    {
        using var wb = WorkbookWith(Kind.Refused);
        var result = wb.Worksheet(SheetName).Cell(At).TryEvaluate();

        await Assert.That(() => _ = result.Value).Throws<InvalidOperationException>();
    }

    /// <summary>
    /// ADR 0001: an expected failure is one definition, shared. Save leaves a cell dirty for exactly
    /// the kinds that <c>TryEvaluate</c> reports, and each of them has a public name, so a kind added
    /// to one is added to the other.
    /// </summary>
    [Test]
    public async Task Save_and_TryEvaluate_tolerate_the_same_failures()
    {
        var disagreements = Enum.GetValues<EvaluationFailureKind>().Where(kind =>
        {
            var expected = EvaluationFailure.IsExpected(kind);
            return (EvaluationPolicy.For(EvaluationEntryPoint.Save, kind) == EvaluationOutcome.LeaveDirty) != expected
                   || (EvaluationPolicy.For(EvaluationEntryPoint.TryEvaluate, kind) == EvaluationOutcome.Report) != expected
                   || (EvaluationPolicy.For(EvaluationEntryPoint.Recalculation, kind) == EvaluationOutcome.LeaveDirty) != expected
                   || EvaluationFailure.TryGetIssueKind(kind, out _) != expected;
        });

        await Assert.That(disagreements).IsEmpty();
        await Assert.That(Enum.GetValues<EvaluationFailureKind>().Count(EvaluationFailure.IsExpected)).IsEqualTo(3);
    }

    [Test]
    public async Task Only_the_exceptions_of_expected_failures_become_issues()
    {
        Exception[] expected =
        [
            new XLCircularReferenceException("cycle"),
            new UnsupportedFeatureException("unsupported"),
            new ExpressionParseException("refused"),
        ];
        Exception[] notExpected =
        [
            new NullReferenceException(),
            new NotImplementedException(),
            new XLNoWorksheetContextException(),
        ];

        foreach (var exception in expected)
            await Assert.That(EvaluationFailure.TryCreateIssue(exception, out _)).IsTrue();

        foreach (var exception in notExpected)
            await Assert.That(EvaluationFailure.TryCreateIssue(exception, out _)).IsFalse();
    }

    /// <summary>
    /// The success path adds nothing to a read of <see cref="IXLCell.Value"/>: no closure, no tuple on
    /// the heap, and the result is a struct.
    /// </summary>
    [Test]
    public async Task The_success_path_does_not_allocate()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet(SheetName).Cell("A1");
        cell.Value = 42;

        for (var i = 0; i < 1_000; i++)
            _ = cell.TryEvaluate();

        var before = GC.GetAllocatedBytesForCurrentThread();
        for (var i = 0; i < 1_000; i++)
            _ = cell.TryEvaluate();

        await Assert.That(GC.GetAllocatedBytesForCurrentThread() - before).IsEqualTo(0);
    }

    private static async Task AssertFailure(XLEvaluationResult result, XLFormulaIssueKind expected)
    {
        await Assert.That(result.IsSuccess).IsFalse();
        await Assert.That(result.Issue).IsNotNull();
        await Assert.That(result.Issue!.Kind).IsEqualTo(expected);
        await Assert.That(result.Issue.Message).IsNotEmpty();
        await Assert.That(result.Issue.Inner).IsNotNull();

        // Only a refusal by the parser has a place in the text.
        if (expected == XLFormulaIssueKind.RefusedFormula)
            await Assert.That(result.Issue.Position).IsNotNull();
        else
            await Assert.That(result.Issue.Position).IsNull();
    }

    private static string Expression(Kind failure, bool qualified) => failure switch
    {
        Kind.Cycle => qualified ? $"{SheetName}!{At}" : At,
        Kind.Unsupported => "Sdemo123|tik!'id1?req?AAPL'",
        Kind.Refused => "1+",
        _ => throw new ArgumentOutOfRangeException(nameof(failure)),
    };
}
