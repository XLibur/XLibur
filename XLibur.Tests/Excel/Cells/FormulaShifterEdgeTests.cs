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
    /// A sheet name written with one apostrophe instead of two names no sheet, so the reference is left
    /// exactly as written. The regex shifter used to strip a closing apostrophe that was never there,
    /// reading <c>'Data!A1</c> as a reference to a sheet named <c>Dat</c>, and a workbook that really
    /// had a sheet by that name got the malformed reference shifted along with it (#576).
    /// </summary>
    /// <remarks>
    /// The truncated name has to hit a real sheet for anything to happen, so the sheet being shifted
    /// here is named for what the old code read rather than for what the reference says: without a
    /// sheet named <c>Dat</c> the reference matched nothing either way. The trailing-apostrophe row is
    /// the external-workbook mechanism seen from the inside — the reference regex's optional
    /// apostrophes are what swallow the closing apostrophe of <c>'[file.xlsx]Data'!A1</c>, so the name
    /// reaching this helper is <c>Data'</c>, which no sheet can be called.
    /// </remarks>
    [Test]
    [Arguments("'Data!A1", "Dat")]
    [Arguments("Data'!A1", "Data")]
    public async Task The_regex_shifter_leaves_a_sheet_name_written_with_one_apostrophe_alone(
        string formula, string shiftedSheetName)
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Formulas");
        var shiftedSheet = (XLWorksheet)wb.AddWorksheet(shiftedSheetName);
        var inserted = (XLRange)shiftedSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftUnparseable(
            formula, formulaSheet, inserted, 3, XLCellFormulaShifter.ShiftAxis.Row);

        await Assert.That(shifted).IsEqualTo(formula);
    }

    /// <summary>
    /// The column axis of the same malformed name. It reads the sheet name through the same helper and
    /// has its own matching test, so both call sites have to agree that no sheet was named.
    /// </summary>
    [Test]
    public async Task A_column_shift_leaves_a_sheet_name_written_with_one_apostrophe_alone()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Formulas");
        var shiftedSheet = (XLWorksheet)wb.AddWorksheet("Dat");
        var inserted = (XLRange)shiftedSheet.Range(1, 1, XLHelper.MaxRowNumber, 2);

        var shifted = XLCellFormulaShifter.ShiftUnparseable(
            "'Data!E5", formulaSheet, inserted, 3, XLCellFormulaShifter.ShiftAxis.Column);

        await Assert.That(shifted).IsEqualTo("'Data!E5");
    }

    /// <summary>
    /// The same malformed name reached the way a workbook reaches it: beside an external workbook
    /// reference, which is what makes the parser refuse the formula and hand it to the regex shifter.
    /// The well-formed reference in the same formula still moves, so the malformed one is being left
    /// alone rather than the whole formula being skipped.
    /// </summary>
    [Test]
    public async Task A_row_shift_leaves_a_malformed_sheet_name_alone_beside_an_external_reference()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var shiftedSheet = (XLWorksheet)wb.AddWorksheet("Dat");
        var inserted = (XLRange)shiftedSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftFormulaRows(
            "'[file.xlsx]Sheet'!A1+'Data!A1+Dat!A5", formulaSheet, inserted, 3);

        await Assert.That(shifted).IsEqualTo("'[file.xlsx]Sheet'!A1+'Data!A1+Dat!A8");
    }

    /// <summary>
    /// <c>!</c> is a legal sheet-name character (<see cref="XLHelper.TryValidateSheetName"/> excludes
    /// <c>: \ / ? * [ ]</c> and not this one), so <c>Q1!Sales</c> is a name a workbook can really have,
    /// and Excel writes a reference to it as <c>'Q1!Sales'!A5</c>. The regex shifter used to split that
    /// match at its first <c>!</c> — the one inside the name — reading the name as <c>'Q1</c>, which is
    /// unterminated and reported as no sheet at all, so the reference was left where it was (#584).
    /// Reached through the production entry point, alongside the external workbook reference that is
    /// what makes the parser refuse the formula and hand it to this regex fallback in the first place.
    /// </summary>
    [Test]
    public async Task A_row_shift_moves_a_reference_to_a_sheet_whose_name_holds_an_exclamation_mark()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var exclamationSheet = (XLWorksheet)wb.AddWorksheet("Q1!Sales");
        var inserted = (XLRange)exclamationSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftFormulaRows(
            "'[file.xlsx]Sheet'!A1+'Q1!Sales'!A5", formulaSheet, inserted, 3);

        await Assert.That(shifted).IsEqualTo("'[file.xlsx]Sheet'!A1+'Q1!Sales'!A8");
    }

    /// <summary>
    /// The column axis of the same defect. The column path's range extraction split at the first
    /// <c>!</c> too, so it needed its own failure to prove the shared fix covers it — fixing only the
    /// sheet-name extraction would still hand a malformed range address to <c>Range(string)</c> for a
    /// reference like this one.
    /// </summary>
    [Test]
    public async Task A_column_shift_moves_a_reference_to_a_sheet_whose_name_holds_an_exclamation_mark()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var exclamationSheet = (XLWorksheet)wb.AddWorksheet("Q1!Sales");
        var inserted = (XLRange)exclamationSheet.Range(1, 1, XLHelper.MaxRowNumber, 2);

        var shifted = XLCellFormulaShifter.ShiftFormulaColumns(
            "'[file.xlsx]Sheet'!A1+'Q1!Sales'!E5", formulaSheet, inserted, 3);

        await Assert.That(shifted).IsEqualTo("'[file.xlsx]Sheet'!A1+'Q1!Sales'!H5");
    }

    /// <summary>
    /// The related case the issue flagged as worth checking: a workbook holding both a sheet named
    /// <c>Q1!Sales</c> and a sheet actually named <c>Q</c>. Before #579 the truncated read of
    /// <c>'Q1!Sales'!A5</c> could have matched <c>Q</c> and handed <c>Sales'!A5</c> to
    /// <c>Range(string)</c>; #579 closed that specific path by rejecting the unterminated name, but the
    /// fixed extraction has to read the whole <c>Q1!Sales</c> name correctly with a same-prefixed real
    /// sheet in the workbook too, not just in isolation.
    /// </summary>
    [Test]
    public async Task A_reference_to_a_sheet_whose_name_holds_an_exclamation_mark_moves_alongside_a_shorter_same_prefixed_sheet()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        var exclamationSheet = (XLWorksheet)wb.AddWorksheet("Q1!Sales");
        wb.AddWorksheet("Q");
        var inserted = (XLRange)exclamationSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftUnparseable(
            "'Q1!Sales'!A5", formulaSheet, inserted, 3, XLCellFormulaShifter.ShiftAxis.Row);

        await Assert.That(shifted).IsEqualTo("'Q1!Sales'!A8");
    }

    /// <summary>
    /// The other side of the same pairing: shifting the shorter sheet <c>Q</c> must not touch a
    /// reference naming the longer <c>Q1!Sales</c>, since the extracted name is the whole quoted name
    /// and never just its prefix.
    /// </summary>
    [Test]
    public async Task A_shift_of_a_shorter_same_prefixed_sheet_leaves_a_reference_to_the_exclamation_marked_sheet_alone()
    {
        using var wb = new XLWorkbook();
        var formulaSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Q1!Sales");
        var qSheet = (XLWorksheet)wb.AddWorksheet("Q");
        var inserted = (XLRange)qSheet.Range(1, 1, 3, XLHelper.MaxColumnNumber);

        var shifted = XLCellFormulaShifter.ShiftUnparseable(
            "'Q1!Sales'!A5", formulaSheet, inserted, 3, XLCellFormulaShifter.ShiftAxis.Row);

        await Assert.That(shifted).IsEqualTo("'Q1!Sales'!A5");
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
