using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// What a caller sees when a formula fails, for every public entry point into the calc engine and
/// every kind of failure (spec 56).
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Matrix"/> is the policy table, observed from outside. Each row names an entry point, a
/// kind of failure, and what the caller gets. A row changes only when spec 56 decided that cell; the
/// rest record what the code did when the matrix was first measured.
/// </para>
/// <para>
/// An exception is shown by the most derived type a caller outside XLibur can name in a
/// <c>catch</c>, so an internal type shows as its nearest public base. <c>n/a</c> marks a cell no
/// input can reach: <c>EvaluateExpr</c> has no workbook, so no cells; <c>TryInvoke</c> has no
/// cells and parses nothing, and its arguments are scalars; and a loaded workbook's engine cannot
/// be given the test's defect function before the load recalculates.
/// </para>
/// </remarks>
public class EvaluationOutcomeTests
{
    public enum Entry
    {
        Value,
        TryGetValue,
        GetFormattedString,
        Search,
        WorksheetEvaluate,
        WorkbookEvaluate,
        EvaluateExpr,
        TryInvoke,
        RecalculateAllFormulas,
        RecalculateOnLoad,
        Save,
    }

    public enum Kind
    {
        Cycle,
        Unsupported,
        Refused,
        NoContext,
        Pending,
        Defect,
    }

    private const string SheetName = "Sheet1";
    private const string At = "A6";
    private const string CycleFormula = "A6+1";

    // Dynamic data exchange: the parser reads it, the calc engine does not evaluate it.
    private const string UnsupportedFormula = "Sdemo123|tik!'id1?req?AAPL'";
    private const string RefusedFormula = "1+";
    private const string DefectFunction = "XLIBURDEFECT";

    [Test]
    [Arguments(Entry.Value, Kind.Cycle, "throws XLCircularReferenceException")]
    [Arguments(Entry.Value, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.Value, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.Value, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.Value, Kind.Pending, "3")]
    [Arguments(Entry.Value, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.TryGetValue, Kind.Cycle, "false")]
    [Arguments(Entry.TryGetValue, Kind.Unsupported, "false")]
    [Arguments(Entry.TryGetValue, Kind.Refused, "false")]
    [Arguments(Entry.TryGetValue, Kind.NoContext, "false")]
    [Arguments(Entry.TryGetValue, Kind.Pending, "true: 3")]
    [Arguments(Entry.TryGetValue, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.GetFormattedString, Kind.Cycle, "42")]
    [Arguments(Entry.GetFormattedString, Kind.Unsupported, "42")]
    [Arguments(Entry.GetFormattedString, Kind.Refused, "42")]
    [Arguments(Entry.GetFormattedString, Kind.NoContext, "42")]
    [Arguments(Entry.GetFormattedString, Kind.Pending, "3")]
    [Arguments(Entry.GetFormattedString, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.Search, Kind.Cycle, "found A1")]
    [Arguments(Entry.Search, Kind.Unsupported, "found A1")]
    [Arguments(Entry.Search, Kind.Refused, "found A1")]
    [Arguments(Entry.Search, Kind.NoContext, "found A1")]
    [Arguments(Entry.Search, Kind.Pending, "found A1")]
    [Arguments(Entry.Search, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.WorksheetEvaluate, Kind.Cycle, "throws XLCircularReferenceException")]
    [Arguments(Entry.WorksheetEvaluate, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.WorksheetEvaluate, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.WorksheetEvaluate, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.WorksheetEvaluate, Kind.Pending, "3")]
    [Arguments(Entry.WorksheetEvaluate, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Cycle, "throws Exception")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Pending, "throws Exception")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.EvaluateExpr, Kind.Cycle, "n/a")]
    [Arguments(Entry.EvaluateExpr, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.EvaluateExpr, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.EvaluateExpr, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.EvaluateExpr, Kind.Pending, "n/a")]
    [Arguments(Entry.EvaluateExpr, Kind.Defect, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.Cycle, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.Unsupported, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.Refused, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.TryInvoke, Kind.Pending, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.Defect, "n/a")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Cycle, "throws XLCircularReferenceException")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Pending, "completes, A6 = 3")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Cycle, "throws XLCircularReferenceException")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.RecalculateOnLoad, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Pending, "opens, A6 = 3")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Defect, "n/a")]
    [Arguments(Entry.Save, Kind.Cycle, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.Unsupported, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.Refused, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.NoContext, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.Pending, "saves, A6 <v>3</v>")]
    [Arguments(Entry.Save, Kind.Defect, "saves, A6 has no <v>")]
    public async Task Matrix(Entry entry, Kind kind, string expected)
    {
        await Assert.That(Observe(entry, kind)).IsEqualTo(expected);
    }

    /// <summary>
    /// D57. <c>wb.Evaluate</c> on a dirty cell used to throw an internal exception with the message
    /// "Exception of type … was thrown", while <c>ws.Evaluate</c> on the same cell answered.
    /// </summary>
    [Test]
    public async Task D57_workbook_evaluate_calculates_a_dirty_cell_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "1+1";

        await Assert.That(wb.Evaluate("Sheet1!A1")).IsEqualTo(2);
    }

    [Test]
    public async Task D57_workbook_evaluate_calculates_a_dirty_precedent_of_an_expression()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "1+1";

        await Assert.That(wb.Evaluate("Sheet1!A1+0")).IsEqualTo(2);
    }

    [Test]
    public async Task D57_workbook_evaluate_sees_an_input_edited_after_the_cell_was_read()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").FormulaA1 = "A1*2";
        _ = ws.Cell("B1").Value;
        ws.Cell("A1").Value = 5;

        await Assert.That(wb.Evaluate("Sheet1!B1")).IsEqualTo(10);
    }

    /// <summary>
    /// The same leak as D57, by a second road: a defined name is evaluated in a context of its own,
    /// which calculated nothing first even when the caller's context did.
    /// </summary>
    [Test]
    public async Task Worksheet_evaluate_of_a_defined_name_calculates_a_dirty_cell_first()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("B1").FormulaA1 = "1+1";
        wb.DefinedNames.Add("Dirty", "Sheet1!$B$1+0");

        await Assert.That(ws.Evaluate("Dirty")).IsEqualTo(2);
    }

    /// <summary>
    /// D58. A cycle was reported as an internal type, which a caller could catch only as
    /// <see cref="InvalidOperationException"/> — the type that also covers misuse and defects.
    /// </summary>
    [Test]
    public async Task D58_a_circular_reference_reaches_the_caller_as_a_public_type()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet(SheetName).Cell("A1");
        cell.FormulaA1 = "A1+1";

        var ex = await Assert.That(() => _ = cell.Value).Throws<XLCircularReferenceException>();
        await Assert.That(ex!.GetType().IsVisible).IsTrue();

        // Still the type a cycle was reported as before, so a caller's existing catch still runs.
        await Assert.That(ex is InvalidOperationException).IsTrue();
        await Assert.That(ex.Message).IsEqualTo("Formula in a cell '$Sheet1'!$A1 is part of a cycle.");
    }

    /// <summary>
    /// D59. Excel returns 2; XLibur does not apply a scalar function over an array argument yet
    /// (spec 30). Until it does, this is an unsupported feature, not a defect.
    /// </summary>
    [Test]
    public async Task D59_an_array_passed_to_a_scalar_parameter_is_an_unsupported_feature()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet(SheetName).Cell("A1");
        cell.FormulaA1 = "LEN({\"ab\",\"c\"})";

        var ex = await Assert.That(() => _ = cell.Value).Throws<NotImplementedException>();
        await Assert.That(ex!.Message).IsEqualTo("Array formulas not implemented.");
        await Assert.That(EvaluationFailure.Classify(ex)).IsEqualTo(EvaluationFailureKind.Unsupported);
    }

    /// <summary>
    /// D60. The cell calling the name was not passed to it, so <c>ROW()</c> reported that it had no
    /// cell to be relative to — and advised using it in a cell formula, which it already was.
    /// </summary>
    [Test]
    public async Task D60_a_defined_name_calling_ROW_answers_with_its_calling_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        wb.DefinedNames.Add("MyRow", "ROW()");
        ws.Cell("A6").FormulaA1 = "MyRow";

        await Assert.That(ws.Cell("A6").Value).IsEqualTo(6);
    }

    /// <summary>
    /// D61. Save swallowed every evaluation failure, so a bug in a function was written to the file
    /// exactly as a formula XLibur cannot evaluate. ADR 0001: a defect reaches the caller.
    /// </summary>
    [Test]
    public async Task D61_a_defect_during_save_reaches_the_caller()
    {
        using var wb = WorkbookWith(Kind.Defect);
        using var stream = new MemoryStream();

        await Assert.That(() => wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true }))
            .Throws<NullReferenceException>();
    }

    /// <summary>
    /// ADR 0001: an expected failure leaves the cell with no cached value, and Excel recalculates it
    /// when the file is opened.
    /// </summary>
    [Test]
    [Arguments(Kind.Cycle)]
    [Arguments(Kind.Unsupported)]
    [Arguments(Kind.Refused)]
    public async Task Save_writes_no_cached_value_for_an_expected_failure(Kind kind)
    {
        using var wb = WorkbookWith(kind);
        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(CachedValueInFile(stream, At)).IsEqualTo($"{At} has no <v>");
    }

    internal static string Observe(Entry entry, Kind kind)
    {
        switch (entry)
        {
            case Entry.Value:
                return InCell(kind, cell => Show(cell.Value));
            case Entry.TryGetValue:
                return InCell(kind, cell => cell.TryGetValue(out double value) ? $"true: {Show(value)}" : "false");
            case Entry.GetFormattedString:
                return InCell(kind, cell =>
                {
                    // A cached value to fall back to.
                    ((XLCell)cell).SetOnlyValue(42);
                    return cell.GetFormattedString(CultureInfo.InvariantCulture);
                });
            case Entry.Search:
                return InCell(kind, cell =>
                {
                    cell.Worksheet.Cell("A1").Value = "needle";
                    var found = cell.Worksheet.Search("needle").Select(c => c.Address.ToString());
                    return "found " + string.Join(",", found);
                });
            case Entry.WorksheetEvaluate:
                return InCell(kind, cell => Show(cell.Worksheet.Evaluate(Expression(kind, qualified: false))));
            case Entry.WorkbookEvaluate:
                return InCell(kind, cell => Show(cell.Worksheet.Workbook.Evaluate(Expression(kind, qualified: true))));
            case Entry.EvaluateExpr:
                return kind is Kind.Unsupported or Kind.Refused or Kind.NoContext
                    ? Observe(() => Show(XLWorkbook.EvaluateExpr(Expression(kind, qualified: false))))
                    : "n/a";
            case Entry.TryInvoke:
                return kind is Kind.NoContext
                    ? Observe(() => new XLFunctionLibrary().TryInvoke("ROW", ReadOnlySpan<XLCellValue>.Empty, out var result)
                        ? Show(result)
                        : "no such function")
                    : "n/a";
            case Entry.RecalculateAllFormulas:
                return InCell(kind, cell =>
                {
                    cell.Worksheet.Workbook.RecalculateAllFormulas();
                    return "completes, " + Describe(cell);
                });
            case Entry.RecalculateOnLoad:
                if (kind == Kind.Defect)
                    return "n/a";

                using (var wb = WorkbookWith(kind))
                {
                    using var stream = new MemoryStream();
                    wb.SaveAs(stream);
                    stream.Position = 0;
                    return Observe(() =>
                    {
                        using var loaded = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
                        return "opens, " + Describe(loaded.Worksheet(SheetName).Cell(At));
                    });
                }
            case Entry.Save:
                return InCell(kind, cell =>
                {
                    using var stream = new MemoryStream();
                    cell.Worksheet.Workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });
                    return "saves, " + CachedValueInFile(stream, At);
                });
            default:
                throw new ArgumentOutOfRangeException(nameof(entry));
        }
    }

    /// <summary>
    /// The text an <c>Evaluate</c> entry point is given for <paramref name="kind"/>: the failing cell
    /// itself where the kind needs a cell, the failing expression otherwise.
    /// </summary>
    private static string Expression(Kind kind, bool qualified) => kind switch
    {
        Kind.Cycle or Kind.Pending => qualified ? $"{SheetName}!{At}" : At,
        Kind.Unsupported => UnsupportedFormula,
        Kind.Refused => RefusedFormula,
        Kind.NoContext => "ROW()",
        Kind.Defect => DefectFunction + "()",
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    /// <summary>
    /// A workbook whose <c>Sheet1!A6</c> holds a formula that fails with <paramref name="kind"/>.
    /// </summary>
    internal static XLWorkbook WorkbookWith(Kind kind)
    {
        var wb = kind == Kind.Defect ? WorkbookWithDefectFunction() : new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        var cell = ws.Cell(At);
        switch (kind)
        {
            case Kind.Cycle:
                cell.FormulaA1 = CycleFormula;
                break;
            case Kind.Unsupported:
                cell.FormulaA1 = UnsupportedFormula;
                break;
            case Kind.Refused:
                cell.FormulaA1 = RefusedFormula;
                break;
            case Kind.NoContext:
                // D60: a defined name that needs its calling cell. The only way a formula in a
                // cell can lack a context.
                wb.DefinedNames.Add("MyRow", "ROW()");
                cell.FormulaA1 = "MyRow";
                break;
            case Kind.Pending:
                // Both dirty: reading A6 needs B1 calculated first.
                ws.Cell("B1").FormulaA1 = "2";
                cell.FormulaA1 = "B1+1";
                break;
            case Kind.Defect:
                cell.FormulaA1 = DefectFunction + "()";
                break;
        }

        return wb;
    }

    /// <summary>
    /// A workbook whose calc engine knows one function, which fails the way a defect in XLibur would.
    /// </summary>
    /// <remarks>
    /// The engine is built over a function table of the test's own, so the real table — shared by
    /// every engine in the process — is never touched.
    /// </remarks>
    internal static XLWorkbook WorkbookWithDefectFunction()
    {
        var functions = new FunctionRegistry();
        functions.RegisterFunction(
            DefectFunction,
            0,
            0,
            static (_, _) => throw new NullReferenceException("A function failed the way a bug in XLibur would."),
            FunctionFlags.Scalar);

        var wb = new XLWorkbook();
        wb.CalcEngine = new XLCalcEngine(CultureInfo.CurrentCulture, functions);
        return wb;
    }

    private static string InCell(Kind kind, Func<IXLCell, string> action)
    {
        using var wb = WorkbookWith(kind);
        var cell = wb.Worksheet(SheetName).Cell(At);
        return Observe(() => action(cell));
    }

    private static string Observe(Func<string> action)
    {
        try
        {
            return action();
        }
        catch (Exception e)
        {
            return "throws " + CatchableName(e);
        }
    }

    /// <summary>
    /// The most derived type of <paramref name="exception"/> that a caller outside XLibur can name in
    /// a <c>catch</c>. An internal type shows as its nearest public base.
    /// </summary>
    private static string CatchableName(Exception exception)
    {
        var type = exception.GetType();
        while (!type.IsVisible)
            type = type.BaseType!;

        return type.Name;
    }

    private static string Describe(IXLCell cell) =>
        cell.NeedsRecalculation ? $"{At} dirty" : $"{At} = {Show(cell.CachedValue)}";

    private static string Show(XLCellValue value) => value.ToString(CultureInfo.InvariantCulture);

    private static string Show(double value) => value.ToString(CultureInfo.InvariantCulture);

    private static string CachedValueInFile(MemoryStream stream, string address)
    {
        using var zip = new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read);
        var entry = zip.Entries.First(e => e.FullName.EndsWith("sheet1.xml", StringComparison.OrdinalIgnoreCase));
        using var reader = new StreamReader(entry.Open());
        var sheet = System.Xml.Linq.XDocument.Parse(reader.ReadToEnd());
        var cell = sheet.Descendants().First(e => e.Name.LocalName == "c" && (string?)e.Attribute("r") == address);
        var value = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
        return value is null ? $"{address} has no <v>" : $"{address} <v>{value.Value}</v>";
    }
}
