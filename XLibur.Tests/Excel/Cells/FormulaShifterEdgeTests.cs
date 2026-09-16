using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Cells;

/// <summary>
/// The parts of <see cref="XLCellFormulaShifter"/> the equivalence corpus cannot reach.
/// <para>
/// The corpus drives the two single-block entry points with 2,072 plain formulas, so three things
/// stay dark: the guard clauses that return before any parse, the structured-reference colon
/// protection (no corpus formula contains a structured reference), and the whole scattered-deletion
/// overload, whose fallback decomposes a deletion map into runs rather than calling the regex shifter
/// once.
/// </para>
/// </summary>
public class FormulaShifterEdgeTests
{
    /// <summary>
    /// A scattered whole-row deletion of a formula the parser cannot read. There is no batch regex
    /// shifter, so the map is decomposed into contiguous runs and the single-block fallback is applied
    /// to each, furthest down the sheet first — the one place in the shifter that answers a parse
    /// failure with something other than a straight delegation.
    /// </summary>
    /// <remarks>
    /// Rows 3, 4 and 8 go, so <c>B12</c> loses three rows above it and lands on <c>B9</c>. Getting
    /// there run by run is what the bottom-up order exists for: taking row 8 out first leaves rows 3-4
    /// still meaning rows 3-4.
    /// </remarks>
    [Test]
    public async Task A_scattered_deletion_of_an_unparseable_formula_is_applied_run_by_run()
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var map = XLRowDeletionMap.Create([3, 4, 8])!;

        var shifted = XLCellFormulaShifter.ShiftFormulaRows(
            "'[file.xlsx]Sheet'!A1+B12", sheet, "Sheet1", map);

        await Assert.That(shifted).IsEqualTo("'[file.xlsx]Sheet'!A1+B9");
    }

    /// <summary>
    /// The same overload on the parser path, for a formula the deletion cannot reach. It returns the
    /// original string rather than a re-rendered copy, which is what keeps an unaffected formula to the
    /// cost of one parse.
    /// </summary>
    [Test]
    public async Task A_scattered_deletion_leaves_a_formula_above_it_untouched()
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var map = XLRowDeletionMap.Create([30, 31])!;

        var shifted = XLCellFormulaShifter.ShiftFormulaRows("A1+B2", sheet, "Sheet1", map);

        await Assert.That(shifted).IsEqualTo("A1+B2");
    }

    /// <summary>
    /// A colon inside a single-bracket structured reference is a literal part of the column name, not a
    /// range operator. It is swapped for a same-width placeholder before parsing and restored after, so
    /// a formula carrying one still shifts its ordinary references and comes back with its column name
    /// intact. No corpus row exercises this.
    /// </summary>
    [Test]
    public async Task A_structured_reference_keeps_its_colon_through_a_row_shift()
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var range = (XLRange)sheet.Range(1, 1, 2, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftFormulaRows(
            "SUM(Table1[Some Header: Other])+B2", sheet, range, 3);

        await Assert.That(shifted).IsEqualTo("SUM(Table1[Some Header: Other])+B5");
    }

    /// <summary>
    /// The same protection on the scattered-deletion overload, which restores the placeholder in its
    /// own return rather than sharing the single-block one.
    /// </summary>
    [Test]
    public async Task A_structured_reference_keeps_its_colon_through_a_scattered_deletion()
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var map = XLRowDeletionMap.Create([3, 4, 8])!;

        var shifted = XLCellFormulaShifter.ShiftFormulaRows(
            "SUM(Table1[Some Header: Other])+B12", sheet, "Sheet1", map);

        await Assert.That(shifted).IsEqualTo("SUM(Table1[Some Header: Other])+B9");
    }

    /// <summary>The column axis of the same protection.</summary>
    [Test]
    public async Task A_structured_reference_keeps_its_colon_through_a_column_shift()
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var range = (XLRange)sheet.Range(1, 1, XLHelper.MaxRowNumber, 2);

        var shifted = XLCellFormulaShifter.ShiftFormulaColumns(
            "SUM(Table1[Some Header: Other])+B2", sheet, range, 3);

        await Assert.That(shifted).IsEqualTo("SUM(Table1[Some Header: Other])+E2");
    }

    /// <summary>
    /// An empty or blank formula yields an empty string rather than being handed to the parser. Both
    /// entry points guard this, and a caller storing the result relies on it being empty rather than
    /// null or whitespace.
    /// </summary>
    [Test]
    [Arguments("")]
    [Arguments("   ")]
    public async Task A_blank_formula_shifts_to_empty(string formula)
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var range = (XLRange)sheet.Range(1, 1, 2, XLHelper.MaxColumnNumber);
        var map = XLRowDeletionMap.Create([1])!;

        await Assert.That(XLCellFormulaShifter.ShiftFormulaRows(formula, sheet, range, 3))
            .IsEqualTo(string.Empty);
        await Assert.That(XLCellFormulaShifter.ShiftFormulaColumns(formula, sheet, range, 3))
            .IsEqualTo(string.Empty);
        await Assert.That(XLCellFormulaShifter.ShiftFormulaRows(formula, sheet, "Sheet1", map))
            .IsEqualTo(string.Empty);
    }

    /// <summary>
    /// The production route to the regex shifter: a formula holding an external workbook reference,
    /// which the parser refuses, alongside a reference to a sheet whose name holds an apostrophe. The
    /// regex read <c>'Bob''s'!A5</c> as <c>'s'!A5</c> — a reference to a sheet named <c>s</c> — so the
    /// name never matched the shifted sheet and the reference stayed where it was (#570).
    /// </summary>
    [Test]
    public async Task A_row_shift_moves_a_reference_to_a_sheet_whose_name_holds_an_apostrophe()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var apostropheSheet = (XLWorksheet)wb.AddWorksheet("Bob's");
        var inserted = (XLRange)apostropheSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftFormulaRows(
            "'[file.xlsx]Sheet'!A1+'Bob''s'!A5", formulaSheet, inserted, 3);

        await Assert.That(shifted).IsEqualTo("'[file.xlsx]Sheet'!A1+'Bob''s'!A8");
    }

    /// <summary>
    /// Every written form of a sheet name the regex shifter has to read, each one shifted through the
    /// fallback directly. A name needing no quotes and a quoted name without an apostrophe both worked
    /// before; the two apostrophe cases did not.
    /// </summary>
    /// <remarks>
    /// <c>'''s'</c> and <c>'s'''</c> are not covered because no sheet can be named <c>'s</c> or
    /// <c>s'</c> — a sheet name may not begin or end with an apostrophe
    /// (<see cref="XLHelper.ValidateSheetName"/>). <c>Ann''s</c> stands in for the
    /// consecutive-apostrophe case, which is a legal name, and it is the row that catches the second
    /// half of the defect: reading the name correctly is not enough if the sheet is then looked up by a
    /// name that has had its apostrophes undoubled a second time.
    /// <para>
    /// The fallback is called directly because each of these formulas parses perfectly well, so the
    /// production path would never reach the regex with one.
    /// </para>
    /// </remarks>
    [Test]
    [Arguments("Bob's", "'Bob''s'")]
    [Arguments("Ann''s", "'Ann''''s'")]
    [Arguments("Data", "Data")]
    [Arguments("My Sheet", "'My Sheet'")]
    public async Task The_regex_shifter_reads_every_written_form_of_a_sheet_name(
        string sheetName, string writtenName)
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Formulas");
        var referencedSheet = (XLWorksheet)wb.AddWorksheet(sheetName);
        var inserted = (XLRange)referencedSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftUnparseable(
            writtenName + "!A5", formulaSheet, inserted, 3, XLCellFormulaShifter.ShiftAxis.Row);

        await Assert.That(shifted).IsEqualTo(writtenName + "!A8");

        // The parser path is the one that normally handles these, and the two must agree about which
        // sheet a written name refers to.
        var byParser = XLCellFormulaShifter.ShiftFormulaRows(
            writtenName + "!A5", formulaSheet, inserted, 3);

        await Assert.That(byParser).IsEqualTo(shifted);
    }

    /// <summary>The column axis of the same sheet-name reading.</summary>
    [Test]
    public async Task A_column_shift_moves_a_reference_to_a_sheet_whose_name_holds_an_apostrophe()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var apostropheSheet = (XLWorksheet)wb.AddWorksheet("Bob's");
        var inserted = (XLRange)apostropheSheet.Range(1, 1, XLHelper.MaxRowNumber, 2);

        var shifted = XLCellFormulaShifter.ShiftUnparseable(
            "'Bob''s'!E5", formulaSheet, inserted, 3, XLCellFormulaShifter.ShiftAxis.Column);

        await Assert.That(shifted).IsEqualTo("'Bob''s'!H5");
    }

    /// <summary>
    /// A zero-row or zero-column shift is a no-op, returned before the formula is parsed at all. The
    /// corpus has no zero-shift row because a shift of nothing is not an equivalence case.
    /// </summary>
    [Test]
    public async Task A_zero_shift_returns_the_formula_unparsed()
    {
        using var wb = new XLWorkbook();
        var sheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var range = (XLRange)sheet.Range(1, 1, 2, XLHelper.MaxColumnNumber);

        // Deliberately unparseable: reaching the parser at all would throw rather than return this.
        const string formula = "this is not a formula(";

        await Assert.That(XLCellFormulaShifter.ShiftFormulaRows(formula, sheet, range, 0))
            .IsEqualTo(formula);
        await Assert.That(XLCellFormulaShifter.ShiftFormulaColumns(formula, sheet, range, 0))
            .IsEqualTo(formula);
    }
}
