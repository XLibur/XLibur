using System;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The read APIs that tolerate a formula they cannot evaluate - <c>TryGetValue</c>,
/// <c>GetFormattedString</c> and <c>Search</c> - tolerate only the failures a formula can
/// legitimately produce. Anything else is a defect in the library and reaches the caller.
/// </summary>
public class EvaluationFailureTests
{
    [Test]
    public async Task Failures_a_formula_can_legitimately_produce_are_expected()
    {
        Exception[] failures =
        [
            new NotImplementedException(),
            new NotSupportedException(),
            new ExpressionParseException("unreadable"),
            new MissingContextException(),
            new XLNoWorksheetContextException(),
            new CircularReferenceException("cycle"),
        ];

        var rejected = failures.Where(f => !EvaluationFailure.IsExpected(f)).Select(f => f.GetType().Name);

        await Assert.That(rejected).IsEmpty();
    }

    [Test]
    public async Task Defects_are_not_expected()
    {
        Exception[] defects =
        [
            new NullReferenceException(),
            new IndexOutOfRangeException(),
            new InvalidCastException(),
            new ArgumentException(),
            // What an internal invariant throws - a calculation chain entry with no formula, say.
            new InvalidOperationException(),
        ];

        var accepted = defects.Where(EvaluationFailure.IsExpected).Select(d => d.GetType().Name);

        await Assert.That(accepted).IsEmpty();
    }

    [Test]
    public async Task TryGetValue_on_an_unimplemented_operator_returns_false()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet().Cell("A1");
        cell.FormulaA1 = "SUM(B1:C2 C1:D2)";

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
