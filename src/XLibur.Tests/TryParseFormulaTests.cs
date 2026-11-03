using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Parser;
using XLibur.Tests.Excel.CalcEngine;

namespace XLibur.Tests;

public class TryParseFormulaTests
{
    [Test]
    [Arguments("SUM(A1:B2)")]
    [Arguments("=SUM(A1:B2)")]
    [Arguments("Sheet9!A1+MissingName")]
    [Arguments("#DIV/0!")]
    public async Task An_accepted_text_returns_true_and_no_issue(string formula)
    {
        await Assert.That(XLHelper.TryParseFormula(formula, out var issue)).IsTrue();
        await Assert.That(issue).IsNull();
    }

    [Test]
    public async Task Text_the_parser_refuses_is_a_refused_formula_with_the_parsers_exception()
    {
        await Assert.That(XLHelper.TryParseFormula("SUM(", out var issue)).IsFalse();

        await Assert.That(issue!.Kind).IsEqualTo(XLFormulaIssueKind.RefusedFormula);
        await Assert.That(issue.Message).IsNotEmpty();
        await Assert.That(issue.Inner).IsTypeOf<ParsingException>();
        await Assert.That(issue.Position).IsNull();
    }

    [Test]
    [Arguments("ABS(1,2)")]
    [Arguments("ABS()")]
    [Arguments("=ABS(1,2)")]
    public async Task Text_a_factory_refuses_is_reported_the_same_way(string formula)
    {
        await Assert.That(XLHelper.TryParseFormula(formula, out var issue)).IsFalse();

        await Assert.That(issue!.Kind).IsEqualTo(XLFormulaIssueKind.RefusedFormula);
        await Assert.That(issue.Message).Contains("parameters for function 'ABS'");
        await Assert.That(issue.Inner).IsTypeOf<ExpressionParseException>();
    }

    [Test]
    public async Task A_leading_equals_does_not_change_the_result()
    {
        await Assert.That(XLHelper.TryParseFormula("SUM(", out var plain)).IsFalse();
        await Assert.That(XLHelper.TryParseFormula("=SUM(", out var withEquals)).IsFalse();

        await Assert.That(withEquals!.Kind).IsEqualTo(plain!.Kind);
    }

    [Test]
    public async Task Null_is_a_programmer_error_not_an_issue()
    {
        await Assert.That(() => XLHelper.TryParseFormula(null!, out _)).Throws<ArgumentNullException>();
    }

    /// <summary>
    /// Every form of the formula-text corpus is accepted exactly when evaluating it does not throw
    /// <see cref="ExpressionParseException"/>, and no form lets an exception out.
    /// </summary>
    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task The_corpus_is_accepted_exactly_when_evaluation_does_not_refuse_it(
        FormulaTextCorpusTests.CorpusRow row)
    {
        var refusedByEvaluation = row.Cells["evaluate"] == "THROWS ExpressionParseException";

        await Assert.That(XLHelper.TryParseFormula(row.Text, out var issue)).IsEqualTo(!refusedByEvaluation);
        await Assert.That(issue is null).IsEqualTo(!refusedByEvaluation);
    }

    /// <summary>
    /// Reading <c>FormulaR1C1</c> converts the text, which the parser alone decides. A text it refuses
    /// is refused by <c>TryParseFormula</c> too.
    /// </summary>
    [Test]
    [MethodDataSource(typeof(FormulaTextCorpusTests), nameof(FormulaTextCorpusTests.Rows))]
    public async Task The_corpus_agrees_with_whether_FormulaR1C1_throws(FormulaTextCorpusTests.CorpusRow row)
    {
        if (string.IsNullOrWhiteSpace(row.Text))
            return;

        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("Sheet1").Cell("C3");
        cell.FormulaA1 = row.Text;

        var refused = false;
        try
        {
            _ = cell.FormulaR1C1;
        }
        catch (ExpressionParseException)
        {
            refused = true;
        }

        await Assert.That(XLHelper.TryParseFormula(row.Text, out _)).IsEqualTo(!refused);
    }
}
