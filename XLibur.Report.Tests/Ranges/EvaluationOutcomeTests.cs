using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// What generating a report does when a template cell's formula fails, for every kind of failure —
/// the Report row of spec 56's policy table.
/// </summary>
/// <remarks>
/// A formula cell is read only inside a bound range, where the expander looks through every cell
/// for tags and expressions. Outside one, a formula cell is left alone and never evaluated, so no
/// failure can reach generation from there.
/// </remarks>
public class EvaluationOutcomeTests
{
    public enum Kind
    {
        Cycle,
        Unsupported,
        Refused,
        NoContext,
        Pending,
        Defect,
    }

    private const string DefectFunction = "XLIBURDEFECT";

    /// <summary>Dynamic data exchange: the parser reads it, the calc engine does not evaluate it.</summary>
    private const string UnsupportedFormula = "Sdemo123|tik!'id1?req?AAPL'";

    /// <summary>Text the parser cannot read, which <c>FormulaA1</c> stores all the same.</summary>
    private const string RefusedFormula = "1+";

    private const string UnsupportedMessage =
        "The formula in this cell uses, or depends on, a feature XLibur does not evaluate, so its value cannot be read.";

    private const string RefusedMessage =
        "The formula in this cell is, or depends on, a formula XLibur cannot parse, so its value cannot be read.";

    [Test]
    [Arguments(Kind.Cycle, "generates, template error: The formula in this cell is part of, or depends on, a circular reference, so its value cannot be read.")]
    [Arguments(Kind.Unsupported, "generates, template error: " + UnsupportedMessage)]
    [Arguments(Kind.Refused, "generates, template error: " + RefusedMessage)]
    [Arguments(Kind.NoContext, "generates")]
    [Arguments(Kind.Pending, "generates")]
    [Arguments(Kind.Defect, "throws NullReferenceException")]
    public async Task A_formula_in_a_bound_range(Kind kind, string expected)
    {
        await Assert.That(Observe(kind, inBoundRange: true)).IsEqualTo(expected);
    }

    [Test]
    [Arguments(Kind.Cycle)]
    [Arguments(Kind.Unsupported)]
    [Arguments(Kind.Refused)]
    [Arguments(Kind.NoContext)]
    [Arguments(Kind.Pending)]
    [Arguments(Kind.Defect)]
    public async Task A_formula_outside_a_bound_range_is_never_evaluated(Kind kind)
    {
        await Assert.That(Observe(kind, inBoundRange: false)).IsEqualTo("generates");
    }

    /// <summary>
    /// Spec 56 (Q39): a circular reference is a template error at its cell, reported once, and the
    /// cell keeps its formula. The rest of the range is generated.
    /// </summary>
    [Test]
    public async Task A_circular_reference_in_a_bound_range_is_a_template_error_at_its_cell()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("B1").FormulaA1 = "B1+1";
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", new List<SaleItem>
        {
            new() { Product = "Widget", Quantity = 1, UnitPrice = 1m, SoldOn = new DateTime(2026, 1, 1) },
        });
        var result = template.Generate();

        await Assert.That(result.ParsingErrors.Count).IsEqualTo(1);
        await Assert.That(result.ParsingErrors[0].Location).IsEqualTo("Report!B1");
        await Assert.That(result.ParsingErrors[0].Exception is XLibur.Excel.CalcEngine.Exceptions.XLCircularReferenceException).IsTrue();
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("B1").FormulaA1).IsEqualTo("B1+1");
    }

    /// <summary>
    /// #488: an unsupported feature and a refused formula are expected failures, like a cycle
    /// (spec 56, Q22), so they get the cycle's treatment: a template error at the cell, reported
    /// once, and the cell keeps its formula. Both used to throw out of <c>Generate()</c>.
    /// </summary>
    [Test]
    [Arguments(UnsupportedFormula, UnsupportedMessage, typeof(NotImplementedException))]
    [Arguments(RefusedFormula, RefusedMessage, typeof(ExpressionParseException))]
    public async Task An_unevaluable_formula_in_a_bound_range_is_a_template_error_at_its_cell(
        string formula,
        string message,
        Type exceptionType)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("B1").FormulaA1 = formula;
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", new List<SaleItem>
        {
            new() { Product = "Widget", Quantity = 1, UnitPrice = 1m, SoldOn = new DateTime(2026, 1, 1) },
        });
        var result = template.Generate();

        await Assert.That(result.ParsingErrors.Count).IsEqualTo(1);
        await Assert.That(result.ParsingErrors[0].Location).IsEqualTo("Report!B1");
        await Assert.That(result.ParsingErrors[0].Message).IsEqualTo(message);
        await Assert.That(exceptionType.IsInstanceOfType(result.ParsingErrors[0].Exception)).IsTrue();
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("B1").FormulaA1).IsEqualTo(formula);
    }

    /// <summary>
    /// A cell whose own formula is valid, but which reads a cell that cannot be evaluated, fails
    /// with the precedent's kind. It is the cell in the bound range that is reported.
    /// </summary>
    [Test]
    [Arguments(UnsupportedFormula, UnsupportedMessage)]
    [Arguments(RefusedFormula, RefusedMessage)]
    public async Task A_formula_depending_on_an_unevaluable_formula_is_a_template_error_at_its_cell(
        string precedent,
        string message)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("B1").FormulaA1 = "Data!A1+1";
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        // On its own sheet, so removing the empty options row does not move it.
        workbook.AddWorksheet("Data").Cell("A1").FormulaA1 = precedent;

        await Assert.That(Generate(workbook)).IsEqualTo("generates, template error: " + message);
        await Assert.That(sheet.Cell("B1").FormulaA1).IsEqualTo("Data!A1+1");
    }

    /// <summary>
    /// A plain <see cref="NotImplementedException"/> is a defect, not an unsupported feature (spec 56,
    /// Q15), so it still reaches the caller (#459). Report cannot tell the two apart by type, since
    /// the unsupported-feature type is internal to XLibur; it asks the policy table instead.
    /// </summary>
    [Test]
    public async Task A_plain_NotImplementedException_in_a_bound_range_is_a_defect_and_still_throws()
    {
        using var workbook = WorkbookWithDefectFunction(
            static () => new NotImplementedException("A function failed the way a bug in XLibur would."));
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("B1").FormulaA1 = DefectFunction + "()";
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        await Assert.That(Generate(workbook)).IsEqualTo("throws NotImplementedException");
    }

    /// <summary>
    /// Review finding 2, executed. B1 is valid, but reading it falls back to a full recalculation,
    /// which meets the cycle at Z100 and throws (the behaviour Q38 records, filed as a follow-up).
    /// Report took that for B1 being in a cycle: it recorded "The formula in this cell is part of a
    /// circular reference" at Report!B1 and read B1 as blank. A cell outside the cycle must not be
    /// blamed for it. Fixed with #492: the read no longer throws for a cycle B1 does not depend on,
    /// and B1 reads 3.
    /// </summary>
    [Test]
    public async Task A_cycle_elsewhere_is_not_blamed_on_the_cell_that_was_read()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("B1").FormulaA1 = "G20+1";
        sheet.Cell("G20").FormulaA1 = "2";
        sheet.Cell("Z100").FormulaA1 = "Z100+1";
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", new List<SaleItem>
        {
            new() { Product = "Widget", Quantity = 1, UnitPrice = 1m, SoldOn = new DateTime(2026, 1, 1) },
        });

        // Since #492 the read of B1 no longer meets the cycle at Z100, so generation must not throw
        // for it: a throw here is the regression this test guards against.
        var result = template.Generate();

        await Assert.That(result.ParsingErrors.Select(e => e.Location)).DoesNotContain("Report!B1");
    }

    [Test]
    public async Task A_template_expression_without_a_worksheet_is_a_template_error()
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("Report").Cell("A1").Value = "{{ ROW() }}";

        await Assert.That(Generate(workbook)).IsEqualTo(
            "generates, template error: <input>(1,1) : error : This function needs a worksheet to work against, "
            + "which a template expression does not have. Use it in a cell formula instead.");
    }

    private static string Observe(Kind kind, bool inBoundRange)
    {
        using var workbook = kind == Kind.Defect
            ? WorkbookWithDefectFunction(static () => new NullReferenceException("A function failed the way a bug in XLibur would."))
            : new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        var at = inBoundRange ? "B1" : "E10";
        var cell = sheet.Cell(at);
        switch (kind)
        {
            case Kind.Cycle:
                cell.FormulaA1 = at + "+1";
                break;
            case Kind.Unsupported:
                cell.FormulaA1 = UnsupportedFormula;
                break;
            case Kind.Refused:
                cell.FormulaA1 = RefusedFormula;
                break;
            case Kind.NoContext:
                workbook.DefinedNames.Add("MyRow", "ROW()");
                cell.FormulaA1 = "MyRow";
                break;
            case Kind.Pending:
                sheet.Cell("G20").FormulaA1 = "2";
                cell.FormulaA1 = "G20+1";
                break;
            case Kind.Defect:
                cell.FormulaA1 = DefectFunction + "()";
                break;
        }

        return Generate(workbook);
    }

    private static string Generate(XLWorkbook workbook)
    {
        try
        {
            using var template = new XLTemplate(workbook);
            template.AddVariable("Items", new List<SaleItem>
            {
                new() { Product = "Widget", Quantity = 1, UnitPrice = 1m, SoldOn = new DateTime(2026, 1, 1) },
            });

            var result = template.Generate();
            return result.ParsingErrors.Count == 0
                ? "generates"
                : "generates, template error: " + string.Join(" | ", result.ParsingErrors.Select(e => e.Message));
        }
        catch (Exception e)
        {
            // The most derived type a caller outside XLibur can name in a catch.
            var type = e.GetType();
            while (!type.IsVisible)
                type = type.BaseType!;

            return "throws " + type.Name;
        }
    }

    /// <summary>
    /// A workbook whose calc engine knows one function, which fails the way a defect in XLibur would:
    /// it throws what <paramref name="defect"/> makes.
    /// </summary>
    private static XLWorkbook WorkbookWithDefectFunction(Func<Exception> defect)
    {
        var functions = new FunctionRegistry();
        functions.RegisterFunction(
            DefectFunction,
            0,
            0,
            (_, _) => throw defect(),
            FunctionFlags.Scalar);

        var workbook = new XLWorkbook();
        workbook.CalcEngine = new XLCalcEngine(CultureInfo.CurrentCulture, functions);
        return workbook;
    }
}
