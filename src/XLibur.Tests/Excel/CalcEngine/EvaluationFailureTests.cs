using System;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The read APIs that tolerate a formula they cannot evaluate - <c>TryGetValue</c>,
/// <c>GetFormattedString</c> and <c>Search</c> - tolerate only the failures a formula can
/// legitimately produce. Anything else is a defect in the library and reaches the caller.
/// </summary>
public class EvaluationFailureTests
{
    [Test]
    public async Task Each_failure_a_formula_can_legitimately_produce_has_its_own_kind()
    {
        (Exception Failure, EvaluationFailureKind Kind)[] failures =
        [
            (new XLCircularReferenceException("cycle"), EvaluationFailureKind.Cycle),
            (new UnsupportedFeatureException("unsupported"), EvaluationFailureKind.Unsupported),
            (new NotSupportedException(), EvaluationFailureKind.Unsupported),
            (new ExpressionParseException("unreadable"), EvaluationFailureKind.Refused),
            (new MissingContextException(), EvaluationFailureKind.NoContext),
            (new XLNoWorksheetContextException(), EvaluationFailureKind.NoContext),
            (new GettingDataException(new SheetPoint(1, new Point(1, 1))), EvaluationFailureKind.Pending),
        ];

        var misclassified = failures
            .Where(f => EvaluationFailure.Classify(f.Failure) != f.Kind)
            .Select(f => $"{f.Failure.GetType().Name} as {EvaluationFailure.Classify(f.Failure)}");

        await Assert.That(misclassified).IsEmpty();
    }

    [Test]
    public async Task Defects_are_classified_as_defects()
    {
        Exception[] defects =
        [
            new NullReferenceException(),
            new IndexOutOfRangeException(),
            new InvalidCastException(),
            new ArgumentException(),
            // What an internal invariant throws - a calculation chain entry with no formula, say.
            new InvalidOperationException(),
            // Spec 56 (Q15): the calc engine raises an unsupported feature as
            // UnsupportedFeatureException, so a plain NotImplementedException, thrown anywhere else
            // in the library, is a defect. Before spec 56 it was tolerated as "unimplemented".
            new NotImplementedException(),
        ];

        var misclassified = defects
            .Where(d => EvaluationFailure.Classify(d) != EvaluationFailureKind.Defect)
            .Select(d => d.GetType().Name);

        await Assert.That(misclassified).IsEmpty();
    }

    [Test]
    public async Task TryGetValue_on_an_unimplemented_feature_returns_false()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet().Cell("A1");
        // Dynamic data exchange parses, but the calc engine cannot evaluate it.
        cell.FormulaA1 = "Sdemo123|tik!'id1?req?AAPL'";

        await Assert.That(cell.TryGetValue(out double _)).IsFalse();
    }

    [Test]
    public async Task GetFormattedString_on_a_circular_formula_falls_back_to_the_cached_value()
    {
        using var wb = new XLWorkbook();
        var cell = (XLCell)wb.AddWorksheet().Cell("A1");
        cell.FormulaA1 = "A1+1";
        cell.SetOnlyValue(42);

        await Assert.That(cell.GetFormattedString()).IsEqualTo("42");
    }

    [Test]
    public async Task Search_skips_cells_whose_formula_cannot_be_evaluated()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "A1+1";
        ws.Cell("A2").FormulaA1 = "SUM(B1:C2 C1:D2)";
        ws.Cell("A3").Value = "needle";

        var found = string.Join(",", ws.Search("needle").Select(c => c.Address.ToString()));

        await Assert.That(found).IsEqualTo("A3");
    }
}
