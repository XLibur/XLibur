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
/// kind of failure, and what the caller gets. A row changes only when a decision changes that cell:
/// spec 56's, or a follow-up to it such as #489 and #490, which gave recalculation the same outcome
/// for a refused formula and an unsupported feature as for a cycle. The rest record what the code did
/// when the matrix was first measured.
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

        /// <summary>
        /// Not a failure of the cell itself: its read falls back to a full recalculation, which
        /// meets a cycle the cell does not depend on (#492).
        /// </summary>
        CycleElsewhere,

        /// <summary>
        /// As <see cref="CycleElsewhere"/>, but the full recalculation meets an unsupported feature the
        /// cell does not depend on.
        /// </summary>
        UnsupportedElsewhere,

        /// <summary>
        /// As <see cref="CycleElsewhere"/>, but the full recalculation meets a refused formula the cell
        /// does not depend on.
        /// </summary>
        RefusedElsewhere,
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
    [Arguments(Entry.Value, Kind.NoContext, "6")]
    [Arguments(Entry.Value, Kind.Pending, "3")]
    [Arguments(Entry.Value, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.TryGetValue, Kind.Cycle, "false")]
    [Arguments(Entry.TryGetValue, Kind.Unsupported, "false")]
    [Arguments(Entry.TryGetValue, Kind.Refused, "false")]
    [Arguments(Entry.TryGetValue, Kind.NoContext, "true: 6")]
    [Arguments(Entry.TryGetValue, Kind.Pending, "true: 3")]
    [Arguments(Entry.TryGetValue, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.GetFormattedString, Kind.Cycle, "42")]
    [Arguments(Entry.GetFormattedString, Kind.Unsupported, "42")]
    [Arguments(Entry.GetFormattedString, Kind.Refused, "42")]
    [Arguments(Entry.GetFormattedString, Kind.NoContext, "6")]
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
    [Arguments(Entry.WorkbookEvaluate, Kind.Cycle, "throws XLCircularReferenceException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Unsupported, "throws NotImplementedException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Refused, "throws ExpressionParseException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.NoContext, "throws XLNoWorksheetContextException")]
    [Arguments(Entry.WorkbookEvaluate, Kind.Pending, "3")]
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
    [Arguments(Entry.RecalculateAllFormulas, Kind.Cycle, "completes, A6 dirty")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Unsupported, "completes, A6 dirty")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Refused, "completes, A6 dirty")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.NoContext, "completes, A6 = 6")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Pending, "completes, A6 = 3")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Cycle, "opens, A6 dirty")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Unsupported, "opens, A6 dirty")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Refused, "opens, A6 dirty")]
    [Arguments(Entry.RecalculateOnLoad, Kind.NoContext, "opens, A6 = 6")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Pending, "opens, A6 = 3")]
    [Arguments(Entry.RecalculateOnLoad, Kind.Defect, "n/a")]
    [Arguments(Entry.Save, Kind.Cycle, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.Unsupported, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.Refused, "saves, A6 has no <v>")]
    [Arguments(Entry.Save, Kind.NoContext, "saves, A6 <v>6</v>")]
    [Arguments(Entry.Save, Kind.Pending, "saves, A6 <v>3</v>")]
    [Arguments(Entry.Save, Kind.Defect, "throws NullReferenceException")]
    [Arguments(Entry.Value, Kind.CycleElsewhere, "3")]
    [Arguments(Entry.TryGetValue, Kind.CycleElsewhere, "true: 3")]
    [Arguments(Entry.GetFormattedString, Kind.CycleElsewhere, "3")]
    [Arguments(Entry.Search, Kind.CycleElsewhere, "found A1")]
    [Arguments(Entry.WorksheetEvaluate, Kind.CycleElsewhere, "3")]
    [Arguments(Entry.WorkbookEvaluate, Kind.CycleElsewhere, "3")]
    [Arguments(Entry.EvaluateExpr, Kind.CycleElsewhere, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.CycleElsewhere, "n/a")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.CycleElsewhere, "completes, A6 = 3")]
    [Arguments(Entry.RecalculateOnLoad, Kind.CycleElsewhere, "opens, A6 = 3")]
    [Arguments(Entry.Save, Kind.CycleElsewhere, "saves, A6 <v>3</v>")]
    [Arguments(Entry.Value, Kind.UnsupportedElsewhere, "3")]
    [Arguments(Entry.TryGetValue, Kind.UnsupportedElsewhere, "true: 3")]
    [Arguments(Entry.GetFormattedString, Kind.UnsupportedElsewhere, "3")]
    [Arguments(Entry.Search, Kind.UnsupportedElsewhere, "found A1")]
    [Arguments(Entry.WorksheetEvaluate, Kind.UnsupportedElsewhere, "3")]
    [Arguments(Entry.WorkbookEvaluate, Kind.UnsupportedElsewhere, "3")]
    [Arguments(Entry.EvaluateExpr, Kind.UnsupportedElsewhere, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.UnsupportedElsewhere, "n/a")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.UnsupportedElsewhere, "completes, A6 = 3")]
    [Arguments(Entry.RecalculateOnLoad, Kind.UnsupportedElsewhere, "opens, A6 = 3")]
    [Arguments(Entry.Save, Kind.UnsupportedElsewhere, "saves, A6 <v>3</v>")]
    [Arguments(Entry.Value, Kind.RefusedElsewhere, "3")]
    [Arguments(Entry.TryGetValue, Kind.RefusedElsewhere, "true: 3")]
    [Arguments(Entry.GetFormattedString, Kind.RefusedElsewhere, "3")]
    [Arguments(Entry.Search, Kind.RefusedElsewhere, "found A1")]
    [Arguments(Entry.WorksheetEvaluate, Kind.RefusedElsewhere, "3")]
    [Arguments(Entry.WorkbookEvaluate, Kind.RefusedElsewhere, "3")]
    [Arguments(Entry.EvaluateExpr, Kind.RefusedElsewhere, "n/a")]
    [Arguments(Entry.TryInvoke, Kind.RefusedElsewhere, "n/a")]
    [Arguments(Entry.RecalculateAllFormulas, Kind.RefusedElsewhere, "completes, A6 = 3")]
    [Arguments(Entry.RecalculateOnLoad, Kind.RefusedElsewhere, "opens, A6 = 3")]
    [Arguments(Entry.Save, Kind.RefusedElsewhere, "saves, A6 <v>3</v>")]
    public async Task Matrix(Entry entry, Kind kind, string expected)
    {
        await Assert.That(Observe(entry, kind)).IsEqualTo(expected);
    }

    /// <summary>
    /// The policy table itself, cell by cell. It also covers the cells <see cref="Matrix"/> cannot
    /// reach from outside, such as save meeting a missing context now that a cell always has one.
    /// </summary>
    [Test]
    public async Task The_policy_table()
    {
        var kinds = Enum.GetValues<EvaluationFailureKind>();
        var rows = Enum.GetValues<EvaluationEntryPoint>()
            .Select(entry => $"{entry}: " + string.Join(", ", kinds.Select(kind => EvaluationPolicy.For(entry, kind))));

        await Assert.That(string.Join("\n", rows)).IsEqualTo(
            //                 Cycle,      Unsupported, Refused,    NoContext, Pending, Defect
            "CellValue: Throw, Throw, Throw, Throw, Throw, Throw\n"
            + "TolerantRead: NoValue, NoValue, NoValue, NoValue, Throw, Throw\n"
            + "Evaluate: Throw, Throw, Throw, Throw, Throw, Throw\n"
            + "FunctionLibrary: Throw, Throw, Throw, Throw, Throw, Throw\n"
            + "Recalculation: LeaveDirty, LeaveDirty, LeaveDirty, Throw, Throw, Throw\n"
            + "Save: LeaveDirty, LeaveDirty, LeaveDirty, Throw, Throw, Throw");
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

    [Test]
    public async Task D60_a_defined_name_calling_COLUMN_answers_with_its_calling_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        wb.DefinedNames.Add("MyColumn", "COLUMN()");
        ws.Cell("C3").FormulaA1 = "MyColumn";

        await Assert.That(ws.Cell("C3").Value).IsEqualTo(3);
    }

    /// <summary>
    /// A name read through <c>ws.Evaluate</c> gets the formula address the caller gave, and still
    /// reports a missing one publicly when none was given.
    /// </summary>
    [Test]
    public async Task A_defined_name_in_worksheet_evaluate_gets_the_formula_address()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        wb.DefinedNames.Add("MyRow", "ROW()");

        await Assert.That(ws.Evaluate("MyRow", "B7")).IsEqualTo(7);
        await Assert.That(() => ws.Evaluate("MyRow")).Throws<XLNoWorksheetContextException>();
    }

    /// <summary>
    /// #491. A name needs a worksheet only when its formula does. <c>IXLWorkbook.Evaluate</c> has no
    /// sheet, and a workbook-scoped name whose references all name their sheet answers there, as its
    /// own text does. It threw <see cref="XLNoWorksheetContextException"/>, because the name's lookup
    /// and the context built for its formula both asked for the sheet before anything needed it.
    /// </summary>
    [Test]
    [Arguments("Total")]
    [Arguments("Total+0")]
    public async Task Workbook_evaluate_reads_a_name_that_needs_no_sheet(string expression)
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet(SheetName).Cell("A1").Value = 5;
        wb.DefinedNames.Add("Total", "Sheet1!$A$1*2");

        await Assert.That(wb.Evaluate(expression)).IsEqualTo(10);
    }

    [Test]
    public async Task Workbook_evaluate_reads_a_name_that_refers_to_another_name()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet(SheetName).Cell("A1").Value = 5;
        wb.DefinedNames.Add("Total", "Sheet1!$A$1*2");
        wb.DefinedNames.Add("Twice", "Total*2");

        await Assert.That(wb.Evaluate("Twice")).IsEqualTo(20);
    }

    /// <summary>
    /// A sheet-scoped name referred to with its sheet needs no sheet from the caller: the prefix
    /// says which sheet's names to look in.
    /// </summary>
    [Test]
    public async Task Workbook_evaluate_reads_a_sheet_scoped_name_referred_to_with_its_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 5;
        ws.DefinedNames.Add("Local", "Sheet1!$A$1*3");

        await Assert.That(wb.Evaluate("Sheet1!Local")).IsEqualTo(15);
    }

    /// <summary>
    /// A sheet-scoped name shadows the workbook's only on its own sheet. <c>IXLWorkbook.Evaluate</c>
    /// is on no sheet, so nothing shadows there and an unqualified name is the workbook's. On the
    /// sheet, <c>IXLWorksheet.Evaluate</c> and a cell read the sheet's own, as before.
    /// </summary>
    [Test]
    public async Task Workbook_evaluate_reads_the_workbook_scoped_name_where_a_sheet_has_its_own()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 5;
        ws.Cell("B1").Value = 7;
        wb.DefinedNames.Add("Total", "Sheet1!$A$1*2");
        ws.DefinedNames.Add("Total", "Sheet1!$B$1");
        ws.Cell("C1").FormulaA1 = "Total";

        await Assert.That(wb.Evaluate("Total")).IsEqualTo(10);
        await Assert.That(wb.Evaluate("Sheet1!Total")).IsEqualTo(7);
        await Assert.That(ws.Evaluate("Total")).IsEqualTo(7);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(7);
    }

    /// <summary>
    /// #491 does not give up the check: a name whose formula needs the sheet or the cell still
    /// reports that from <c>IXLWorkbook.Evaluate</c>. <c>!A1</c> is the relative reference a defined
    /// name can hold: a sheet-less <c>A1</c> is refused when the name is defined, and <c>!A1</c> is
    /// read on the sheet of the cell using the name (#446).
    /// </summary>
    [Test]
    [Arguments("!A1")]
    [Arguments("ROW()")]
    [Arguments("COLUMN()")]
    public async Task Workbook_evaluate_still_throws_for_a_name_that_needs_a_cell(string refersTo)
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet(SheetName);
        wb.DefinedNames.Add("Here", refersTo);

        await Assert.That(() => wb.Evaluate("Here")).Throws<XLNoWorksheetContextException>();
        await Assert.That(() => wb.Evaluate("Here+0")).Throws<XLNoWorksheetContextException>();
    }

    /// <summary>
    /// An unqualified name that only a sheet defines needs the sheet to be found, directly or through
    /// another name.
    /// </summary>
    [Test]
    public async Task Workbook_evaluate_still_throws_for_an_unqualified_sheet_scoped_name()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 5;
        ws.DefinedNames.Add("Local", "Sheet1!$A$1*3");
        wb.DefinedNames.Add("Outer", "!Local");

        await Assert.That(() => wb.Evaluate("Local")).Throws<XLNoWorksheetContextException>();
        await Assert.That(() => wb.Evaluate("Outer")).Throws<XLNoWorksheetContextException>();
    }

    /// <summary>
    /// The same names, where there is a sheet: <c>IXLWorksheet.Evaluate</c> and a cell read them as
    /// they did before #491.
    /// </summary>
    [Test]
    public async Task Worksheet_evaluate_and_a_cell_read_names_as_before()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 5;
        ws.Cell("B1").Value = 7;
        wb.DefinedNames.Add("Total", "Sheet1!$A$1*2");
        wb.DefinedNames.Add("Here", "!B1");
        ws.DefinedNames.Add("Local", "Sheet1!$A$1*3");
        wb.DefinedNames.Add("Outer", "!Local");
        ws.Cell("C1").FormulaA1 = "Total";
        ws.Cell("C2").FormulaA1 = "Here";
        ws.Cell("C3").FormulaA1 = "Local";
        ws.Cell("C4").FormulaA1 = "Outer";

        await Assert.That(ws.Evaluate("Total")).IsEqualTo(10);
        await Assert.That(ws.Evaluate("Here")).IsEqualTo(7);
        await Assert.That(ws.Evaluate("Local")).IsEqualTo(15);
        await Assert.That(ws.Evaluate("Outer")).IsEqualTo(15);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(10);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(7);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo(15);
        await Assert.That(ws.Cell("C4").Value).IsEqualTo(15);
    }

    /// <summary>
    /// A defined name that depends on its own value, directly (<c>Loop</c>) or through another name
    /// (<c>Ping</c> and <c>Pong</c>), is a circular reference. Name evaluation had no guard, so it
    /// recursed until the stack overflowed, which ends the process; #491 made that reachable from
    /// <c>IXLWorkbook.Evaluate</c> too. Each entry point now gets spec 56's outcome for a cycle:
    /// <see cref="XLCircularReferenceException"/> from both <c>Evaluate</c>s and from a cell read,
    /// and no value from <c>TryGetValue</c>.
    /// </summary>
    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task A_circular_defined_name_is_a_circular_reference(string name)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        wb.DefinedNames.Add("Loop", "Loop+1");
        wb.DefinedNames.Add("Ping", "Pong");
        wb.DefinedNames.Add("Pong", "Ping");
        ws.Cell("B1").FormulaA1 = name;

        await Assert.That(() => wb.Evaluate(name)).Throws<XLCircularReferenceException>();
        await Assert.That(() => ws.Evaluate(name)).Throws<XLCircularReferenceException>();
        await Assert.That(() => ws.Cell("B1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(ws.Cell("B1").TryGetValue<double>(out _)).IsFalse();
    }

    /// <summary>
    /// A circular name met by the full recalculation a read falls back to. On its own, #491 raised
    /// the cycle part way through that pass, which stopped it, as a cycle did before #492, so a read
    /// of a cell that does not use the name threw. The fallback pass leaves the name's cell dirty, as
    /// it does any expected failure, and only a read of that cell reports the cycle.
    /// </summary>
    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task A_read_does_not_throw_for_a_circular_name_it_does_not_use(string name)
    {
        using var wb = WorkbookWithCircularName(name);
        var ws = wb.Worksheet(SheetName);

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(6);

        // The read fell back to a full pass, which walks every formula, B1 included.
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<XLCircularReferenceException>();
    }

    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task TryGetValue_answers_for_a_cell_that_does_not_use_a_circular_name(string name)
    {
        using var wb = WorkbookWithCircularName(name);
        var ws = wb.Worksheet(SheetName);

        await Assert.That(ws.Cell("C1").TryGetValue(out double value)).IsTrue();
        await Assert.That(value).IsEqualTo(6);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").TryGetValue(out double _)).IsFalse();
    }

    /// <summary>
    /// Recalculation leaves a circular name's cell dirty and calculates the rest, as it does for any
    /// cycle (Q23). Before D79 it overflowed the stack building the dependency tree.
    /// </summary>
    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task RecalculateAllFormulas_leaves_a_circular_name_dirty_and_calculates_the_rest(string name)
    {
        using var wb = WorkbookWithCircularName(name);
        var ws = wb.Worksheet(SheetName);

        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("C1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("C1").CachedValue).IsEqualTo(6);
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<XLCircularReferenceException>();
    }

    /// <summary>
    /// ADR 0001: a save that evaluates formulas writes a circular name's cell with no cached value.
    /// </summary>
    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task Save_writes_no_cached_value_for_a_circular_name(string name)
    {
        using var wb = WorkbookWithCircularName(name);
        using var stream = new MemoryStream();

        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(CachedValueInFile(stream, "B1")).IsEqualTo("B1 has no <v>");
        await Assert.That(CachedValueInFile(stream, "C1")).IsEqualTo("C1 <v>6</v>");
    }

    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task A_workbook_with_a_circular_name_opens_with_recalculate_on_load(string name)
    {
        using var stream = new MemoryStream();
        using (var wb = WorkbookWithCircularName(name))
            wb.SaveAs(stream);

        stream.Position = 0;
        using var loaded = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
        var sheet = loaded.Worksheet(SheetName);

        await Assert.That(sheet.Cell("C1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("C1").CachedValue).IsEqualTo(6);
        await Assert.That(sheet.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = sheet.Cell("B1").Value).Throws<XLCircularReferenceException>();
    }

    /// <summary>
    /// B1 uses a circular name: <c>Loop</c> refers to itself, and <c>Ping</c> and <c>Pong</c> refer
    /// to each other. C1 does not use it, but needs A1, a formula no one has read, so a read of C1
    /// falls back to a full recalculation, which meets B1.
    /// </summary>
    private static XLWorkbook WorkbookWithCircularName(string name)
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        if (name == "Loop")
        {
            wb.DefinedNames.Add("Loop", "Loop+1");
        }
        else
        {
            wb.DefinedNames.Add("Ping", "Pong");
            wb.DefinedNames.Add("Pong", "Ping");
        }

        ws.Cell("A1").FormulaA1 = "2+3";
        ws.Cell("B1").FormulaA1 = name;
        ws.Cell("C1").FormulaA1 = "A1+1";
        return wb;
    }

    /// <summary>
    /// The guard stops a name met again inside its own evaluation, not one used twice side by side.
    /// </summary>
    [Test]
    [Arguments("Total+Total")]
    [Arguments("Both")]
    public async Task A_name_used_twice_is_not_a_circular_reference(string expression)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 5;
        wb.DefinedNames.Add("Total", "Sheet1!$A$1*2");
        wb.DefinedNames.Add("Both", "Total+Total");
        ws.Cell("B1").FormulaA1 = expression;

        await Assert.That(wb.Evaluate(expression)).IsEqualTo(20);
        await Assert.That(ws.Evaluate(expression)).IsEqualTo(20);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(20);
    }

    /// <summary>
    /// A name that reads a cell whose formula uses the same name, for its own cell, is not a cycle.
    /// <c>Above</c> is the cell above the one using it, plus one. Evaluated for A3 it reads A2, and
    /// A2's formula is <c>Above</c> again, calculated for A2 in a context of its own. A guard shared
    /// by every evaluation would see <c>Above</c> twice and call it a cycle.
    /// </summary>
    [Test]
    public async Task A_name_met_again_through_a_cell_it_reads_is_not_a_circular_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 1;
        wb.DefinedNames.Add("Above", "INDEX(Sheet1!$A$1:$A$10,ROW()-1)+1");
        ws.Cell("A2").FormulaA1 = "Above";

        await Assert.That(ws.Evaluate("Above", "A3")).IsEqualTo(3);
    }

    /// <summary>
    /// Review finding (medium). A sheet-only recalculation reads another sheet's cells as they
    /// stand, so a formula on Sheet1 that reads a dirty Sheet2 cell takes its current value. A
    /// defined name's context did not carry the sheet filter: the name asked for the dirty Sheet2
    /// cell, the chain moved it to the front, the filter skipped it, and the pass came back to
    /// Sheet1!A1 for ever. The name must read the way the same formula typed into the cell reads.
    /// </summary>
    [Test]
    public async Task A_sheet_only_recalculation_reads_a_name_the_way_the_cell_would()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet2.Cell("A1").FormulaA1 = "1+1";
        wb.DefinedNames.Add("X", "Sheet2!$A$1*2");
        sheet1.Cell("A1").FormulaA1 = "X";
        sheet1.Cell("B1").FormulaA1 = "Sheet2!$A$1*2";

        // Bounded, so that a hang fails this test instead of stalling the suite.
        await Task.Run(() => sheet1.RecalculateAllFormulas()).WaitAsync(TimeSpan.FromSeconds(10));

        await Assert.That(sheet1.Cell("A1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet1.Cell("A1").CachedValue).IsEqualTo(sheet1.Cell("B1").CachedValue);
        await Assert.That(sheet2.Cell("A1").NeedsRecalculation).IsTrue();
    }

    /// <summary>
    /// Where spec 53 meets D60. A name now has its calling cell, but the name's context keeps
    /// <c>IntersectOperands</c> off, as before, so an operator at the top of the name's formula
    /// keeps its range operand whole. <c>Plus10</c> read in row 2 is the top-left of
    /// <c>{11;12;13}</c>, where the same text typed into a cell in row 2 intersects and gives 12.
    /// </summary>
    /// <remarks>
    /// Verified in Excel by the owner on 2026-09-14, with spec 53's <c>Range.Formula</c> trick:
    /// <c>Range("B2").Formula = "=Plus10"</c> shows 11, and Excel adds <c>@</c>. Excel evaluates a
    /// name's formula as an array, without intersecting its range operands, and the cell's implicit
    /// intersection then takes the first element.
    /// </remarks>
    [Test]
    public async Task A_defined_name_keeps_array_semantics_for_its_operators()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        wb.DefinedNames.Add("Plus10", "Sheet1!$A$1:$A$3+10");
        ws.Cell("B2").FormulaA1 = "Plus10";
        ws.Cell("C2").FormulaA1 = "Sheet1!$A$1:$A$3+10";
        ws.Cell("D2").FormulaA1 = "SUM(Plus10)";

        await Assert.That(ws.Cell("B2").Value).IsEqualTo(11);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(12);
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(36);
    }

    /// <summary>
    /// Q23. Recalculation leaves the cells of a cycle dirty, with every formula that depends on
    /// them, and calculates the rest. Reading a cell of the cycle, or one that depends on it, still
    /// throws.
    /// </summary>
    [Test]
    public async Task RecalculateAllFormulas_skips_a_cycle_and_calculates_the_rest()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "5*2";
        ws.Cell("A2").FormulaA1 = "A3+1";
        ws.Cell("A3").FormulaA1 = "A2+1";
        ws.Cell("A4").FormulaA1 = "A3*2";
        ws.Cell("A5").FormulaA1 = "A1+1";
        ws.Cell("A6").FormulaA1 = "A6";

        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A1").CachedValue).IsEqualTo(10);
        await Assert.That(ws.Cell("A5").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A5").CachedValue).IsEqualTo(11);
        foreach (var address in new[] { "A2", "A3", "A4", "A6" })
            await Assert.That(ws.Cell(address).NeedsRecalculation).IsTrue().Because($"{address} is in or after a cycle");

        await Assert.That(() => _ = ws.Cell("A2").Value).Throws<XLCircularReferenceException>();
        await Assert.That(() => _ = ws.Cell("A4").Value).Throws<XLCircularReferenceException>();
        await Assert.That(() => _ = ws.Cell("A6").Value).Throws<XLCircularReferenceException>();
    }

    [Test]
    public async Task Worksheet_RecalculateAllFormulas_skips_a_cycle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";
        ws.Cell("B1").FormulaA1 = "2*3";

        ws.RecalculateAllFormulas();

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("B1").CachedValue).IsEqualTo(6);
    }

    /// <summary>
    /// Acceptance: a workbook Excel opens can be opened with recalculate-on-load. The constructor
    /// used to throw on the first cycle.
    /// </summary>
    [Test]
    public async Task A_workbook_with_a_cycle_opens_with_recalculate_on_load()
    {
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet(SheetName);
            ws.Cell("A1").Value = 4;
            ws.Cell("B1").FormulaA1 = "C1+1";
            ws.Cell("C1").FormulaA1 = "B1+1";
            ws.Cell("D1").FormulaA1 = "A1*2";
            wb.SaveAs(stream);
        }

        stream.Position = 0;
        using var loaded = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
        var sheet = loaded.Worksheet(SheetName);

        await Assert.That(sheet.Cell("D1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("D1").CachedValue).IsEqualTo(8);
        await Assert.That(sheet.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = sheet.Cell("B1").Value).Throws<XLCircularReferenceException>();
    }

    /// <summary>
    /// Review finding 1. A pass that ends with the chain's cycle flag set must not hand it to the
    /// next pass, whose first cell would then be taken for part of a cycle: A1 was left dirty by
    /// the second recalculation, and once the cycle was gone a read that fell back to full
    /// recalculation threw a circular reference naming A1.
    /// </summary>
    [Test]
    public async Task A_cycle_found_by_one_recalculation_is_not_carried_into_the_next()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "1+1";
        ws.Cell("B1").FormulaA1 = "B1+1";

        wb.RecalculateAllFormulas();
        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A1").CachedValue).IsEqualTo(2);

        ws.Cell("B1").Value = 5;
        ws.Cell("C1").FormulaA1 = "D1+1";
        ws.Cell("D1").FormulaA1 = "2";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);
    }

    /// <summary>
    /// Review finding 2. A pass that throws stops with the chain positioned on the cell that
    /// threw. Unless the chain is reset, the next pass carries on from there. Here that cell has
    /// since been overwritten with a value, so it is no longer in the chain at all, and the next
    /// read failed with the chain's own invariant: "Book point [1]A2 is not in the chain."
    /// </summary>
    /// <remarks>
    /// Found with a cycle at A2. Since #492 a read's pass leaves a cycle dirty instead of throwing,
    /// and it now leaves an unsupported feature and a refused formula dirty too. A defect is the
    /// failure that still stops the pass wherever it is, so the pass is made to throw with one.
    /// </remarks>
    [Test]
    public async Task A_pass_after_one_that_threw_starts_from_the_beginning_of_the_chain()
    {
        using var wb = WorkbookWithDefectFunction();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "B1*2";
        ws.Cell("B1").FormulaA1 = "A3+0";
        ws.Cell("A2").FormulaA1 = DefectFunction + "()";
        ws.Cell("A3").Value = 5;
        ws.Cell("A4").FormulaA1 = "A5+1";
        ws.Cell("A5").FormulaA1 = "1";

        await Assert.That(() => _ = ws.Cell("A4").Value).Throws<NullReferenceException>();

        ws.Cell("A3").Value = 7;
        ws.Cell("A2").Value = 0;

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(14);
        await Assert.That(ws.Cell("A4").Value).IsEqualTo(2);
    }

    /// <summary>
    /// Review finding 1, by a second road. A cell read that meets a cycle throws with the chain's
    /// cycle flag still set. Once the cycle was fixed, the next read took the first cell it reached
    /// for part of a cycle and threw again.
    /// </summary>
    [Test]
    public async Task A_cycle_fixed_after_a_read_threw_is_not_reported_again()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";
        ws.Cell("B1").FormulaA1 = "C1*2";
        ws.Cell("C1").FormulaA1 = "5";

        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();

        ws.Cell("A1").FormulaA1 = "1";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(10);
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
    }

    /// <summary>
    /// #492. A read that falls back to recalculating the whole workbook used to throw on the first
    /// cycle that pass met, wherever it was, so B1 threw for A1's cycle and was left without a
    /// value. The pass now leaves the cycle dirty, as recalculation does, and B1 reads 3. A1 is
    /// still dirty, and reading it still throws.
    /// </summary>
    [Test]
    public async Task Issue492_a_read_does_not_throw_for_a_cycle_the_cell_does_not_depend_on()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";
        ws.Cell("B1").FormulaA1 = "C1+1";
        ws.Cell("C1").FormulaA1 = "2";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("C1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("C1").CachedValue).IsEqualTo(2);
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
    }

    /// <summary>
    /// #492, the other side. A cell that depends on a cycle cannot be calculated either, so reading
    /// it still throws, naming the cycle.
    /// </summary>
    [Test]
    public async Task A_read_behind_a_cycle_still_throws_naming_the_cycle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";
        ws.Cell("D1").FormulaA1 = "A1*2";

        var ex = await Assert.That(() => _ = ws.Cell("D1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(ex!.Message).IsEqualTo("Formula in a cell '$Sheet1'!$A1 is part of a cycle.");
    }

    /// <summary>
    /// #492. With two cycles in the workbook, a read in the second one is named by its own cycle,
    /// not by the first cycle the pass meets.
    /// </summary>
    [Test]
    public async Task A_read_in_a_cycle_is_named_by_its_own_cycle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";
        ws.Cell("C5").FormulaA1 = "D5+1";
        ws.Cell("D5").FormulaA1 = "C5+1";

        var ex = await Assert.That(() => _ = ws.Cell("C5").Value).Throws<XLCircularReferenceException>();
        await Assert.That(ex!.Message).IsEqualTo("Formula in a cell '$Sheet1'!$C5 is part of a cycle.");

        var first = await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(first!.Message).IsEqualTo("Formula in a cell '$Sheet1'!$A1 is part of a cycle.");
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

    /// <summary>
    /// Review finding (low). Save calculated each dirty formula with a single-cell attempt, whose
    /// fallback already ran a full recalculation, and then ran a second one. Measured on main at
    /// 1b476e8f with the same workbook: 6,002 passes. Once one full pass has run during a save,
    /// every formula is either calculated or left dirty by it, and nothing edits the workbook
    /// mid-save, so the rest of the save needs no further pass.
    /// </summary>
    [Test]
    public async Task Save_runs_one_calculation_pass_for_a_cycle_with_many_dependents()
    {
        const int dependents = 3000;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";
        for (var row = 1; row <= dependents; row++)
            ws.Cell(row, 2).FormulaA1 = $"A1+{row}";

        // Calculable, but only by a pass, because D1 is dirty too.
        ws.Cell("C1").FormulaA1 = "D1*2";
        ws.Cell("D1").FormulaA1 = "5";

        var before = wb.CalcEngine.PassCount;
        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(wb.CalcEngine.PassCount - before).IsEqualTo(1);
        await Assert.That(CachedValueInFile(stream, "C1")).IsEqualTo("C1 <v>10</v>");
        await Assert.That(CachedValueInFile(stream, "A1")).IsEqualTo("A1 has no <v>");
        await Assert.That(CachedValueInFile(stream, "B3000")).IsEqualTo("B3000 has no <v>");
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
                return ObserveEvaluateExpr(kind);
            case Entry.TryInvoke:
                return ObserveTryInvoke(kind);
            case Entry.RecalculateAllFormulas:
                return InCell(kind, cell =>
                {
                    cell.Worksheet.Workbook.RecalculateAllFormulas();
                    return "completes, " + Describe(cell);
                });
            case Entry.RecalculateOnLoad:
                return ObserveRecalculateOnLoad(kind);
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

    private static string ObserveEvaluateExpr(Kind kind)
    {
        return kind is Kind.Unsupported or Kind.Refused or Kind.NoContext
            ? Observe(() => Show(XLWorkbook.EvaluateExpr(Expression(kind, qualified: false))))
            : "n/a";
    }

    private static string ObserveTryInvoke(Kind kind)
    {
        return kind is Kind.NoContext
            ? Observe(() => new XLFunctionLibrary().TryInvoke("ROW", ReadOnlySpan<XLCellValue>.Empty, out var result)
                ? Show(result)
                : "no such function")
            : "n/a";
    }

    private static string ObserveRecalculateOnLoad(Kind kind)
    {
        if (kind == Kind.Defect)
            return "n/a";

        using var wb = WorkbookWith(kind);
        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        stream.Position = 0;
        return Observe(() =>
        {
            using var loaded = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
            return "opens, " + Describe(loaded.Worksheet(SheetName).Cell(At));
        });
    }

    /// <summary>
    /// The text an <c>Evaluate</c> entry point is given for <paramref name="kind"/>: the failing cell
    /// itself where the kind needs a cell, the failing expression otherwise.
    /// </summary>
    private static string Expression(Kind kind, bool qualified) => kind switch
    {
        Kind.Cycle or Kind.Pending or Kind.CycleElsewhere or Kind.UnsupportedElsewhere or Kind.RefusedElsewhere
            => qualified ? $"{SheetName}!{At}" : At,
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
                // D60: a defined name that needs its calling cell. Before spec 56 this was the
                // only way a formula in a cell could lack a context; now the name has the cell.
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
            case Kind.CycleElsewhere:
                // A6 needs B1 calculated first, so its read falls back to a full recalculation,
                // which meets the cycle at Z100. A6 does not depend on Z100.
                ws.Cell("B1").FormulaA1 = "2";
                ws.Cell("Z100").FormulaA1 = "Z100+1";
                cell.FormulaA1 = "B1+1";
                break;
            case Kind.UnsupportedElsewhere:
            case Kind.RefusedElsewhere:
                // As CycleElsewhere, but Z100 holds a formula the pass cannot calculate.
                ws.Cell("B1").FormulaA1 = "2";
                ws.Cell("Z100").FormulaA1 = kind == Kind.UnsupportedElsewhere ? UnsupportedFormula : RefusedFormula;
                cell.FormulaA1 = "B1+1";
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

    internal static string CachedValueInFile(MemoryStream stream, string address)
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
