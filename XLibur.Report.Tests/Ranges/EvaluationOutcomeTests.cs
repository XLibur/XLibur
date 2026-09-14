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

    [Test]
    [Arguments(Kind.Cycle, "generates, template error: The formula in this cell is part of a circular reference, so its value cannot be read.")]
    [Arguments(Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Kind.Refused, "throws ExpressionParseException")]
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
        using var workbook = kind == Kind.Defect ? WorkbookWithDefectFunction() : new XLWorkbook();
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
                // Dynamic data exchange: the parser reads it, the calc engine does not evaluate it.
                cell.FormulaA1 = "Sdemo123|tik!'id1?req?AAPL'";
                break;
            case Kind.Refused:
                cell.FormulaA1 = "1+";
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
    /// A workbook whose calc engine knows one function, which fails the way a defect in XLibur would.
    /// </summary>
    private static XLWorkbook WorkbookWithDefectFunction()
    {
        var functions = new FunctionRegistry();
        functions.RegisterFunction(
            DefectFunction,
            0,
            0,
            static (_, _) => throw new NullReferenceException("A function failed the way a bug in XLibur would."),
            FunctionFlags.Scalar);

        var workbook = new XLWorkbook();
        workbook.CalcEngine = new XLCalcEngine(CultureInfo.CurrentCulture, functions);
        return workbook;
    }
}
