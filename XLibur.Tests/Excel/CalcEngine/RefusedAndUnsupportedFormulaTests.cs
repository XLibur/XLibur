using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Drawings;
using XLibur.Tests.Excel.Cells;
using XLibur.Tests.Excel.IO;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A refused formula and an unsupported feature are expected failures (spec 56, Q22). Like a
/// circular reference, each fails only the cell that holds it and the formulas that depend on it.
/// It blocks no write elsewhere in the workbook (#489), and recalculation leaves it dirty and
/// calculates the rest (#489, #490).
/// </summary>
public class RefusedAndUnsupportedFormulaTests
{
    public enum Failure
    {
        Refused,
        Unsupported,
    }

    /// <summary>
    /// The form the formula bar shows for an external reference. <see cref="IXLCell.FormulaA1"/>
    /// accepts it, and the parser refuses it.
    /// </summary>
    private const string Refused = "'[Book2.xlsx]Sheet1'!A1";

    /// <summary>Dynamic data exchange: the parser reads it, the calc engine does not evaluate it.</summary>
    private const string Unsupported = "Sdemo123|tik!'id1?req?AAPL'";

    /// <summary>
    /// #489. A successful evaluation tells the calc engine to build its dependency tree on the next
    /// write. The build parsed every formula in the workbook and threw on the refused one, so every
    /// later write anywhere threw <see cref="ExpressionParseException"/>, after it had been applied.
    /// </summary>
    [Test]
    public async Task Issue489_a_write_after_a_read_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);

        ws.Cell("C1").Value = 5;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(5);
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// #489. The tree the write builds still does its job: the write reaches the formulas that
    /// depend on it.
    /// </summary>
    [Test]
    public async Task Issue489_a_write_still_marks_its_dependents_dirty()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("C1").Value = 1;
        ws.Cell("D1").FormulaA1 = "C1*2";
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(2);

        ws.Cell("C1").Value = 5;

        await Assert.That(ws.Cell("D1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(10);
    }

    /// <summary>
    /// #489. The tree covers the whole workbook, so a refused formula on one sheet blocked writes on
    /// every other.
    /// </summary>
    [Test]
    public async Task Issue489_a_refused_formula_on_another_sheet_does_not_block_a_write()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet2.Cell("A1").FormulaA1 = Refused;
        sheet1.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(sheet1.Cell("B1").Value).IsEqualTo(2);

        sheet1.Cell("C1").Value = 5;

        await Assert.That(sheet1.Cell("C1").Value).IsEqualTo(5);
        await Assert.That(() => _ = sheet2.Cell("A1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// #489. The obvious recovery threw too: setting a value on the refused cell marks the cell dirty
    /// before it clears the formula.
    /// </summary>
    [Test]
    public async Task Issue489_setting_a_value_on_the_refused_cell_recovers_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);

        ws.Cell("A1").Value = 0;

        await Assert.That(ws.Cell("A1").HasFormula).IsFalse();
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(0);
        ws.Cell("C1").Value = 5;
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(5);
    }

    /// <summary>
    /// #489. The writes were half-applied: the formula set on C1 was stored and the value set on C2
    /// was cached, yet each call threw. Both now complete, and the value reaches its dependent.
    /// </summary>
    [Test]
    public async Task Issue489_a_formula_write_and_a_value_write_complete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("C2").Value = 1;
        ws.Cell("D2").FormulaA1 = "C2+1";
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(2);

        ws.Cell("C1").FormulaA1 = "2*3";
        ws.Cell("C2").Value = 7;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("D2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(8);
    }

    /// <summary>
    /// #489, by a second road. Once a recalculation has built the tree, each formula set afterwards
    /// is parsed for its precedents, so a refused formula threw from its own setter, after it had been
    /// stored.
    /// </summary>
    [Test]
    public async Task Issue489_a_refused_formula_can_be_set_after_a_recalculation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("B1").FormulaA1 = "1+1";
        wb.RecalculateAllFormulas();

        ws.Cell("A1").FormulaA1 = Refused;

        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo(Refused);
        ws.Cell("C1").Value = 5;
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(5);
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// Review finding 1: #489 through a defined name. A load keeps a name whose text the parser
    /// refuses, so that one bad name cannot stop the workbook from opening. The dependency tree read
    /// the text of each name a formula uses, and threw on that one, so every write after a read threw,
    /// and so did a recalculation.
    /// </summary>
    [Test]
    public async Task Issue489_a_refused_defined_name_does_not_block_a_write()
    {
        using var package = BookWithRefusedDefinedName();
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Sheet1");
        await Assert.That(wb.DefinedNames.Single(name => name.Name == "Ext").RefersTo).IsEqualTo(Refused);
        ws.Cell("C1").FormulaA1 = "1+1";
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2);

        ws.Cell("D1").Value = 5;

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(5);
        wb.RecalculateAllFormulas();
        await Assert.That(ws.Cell("C1").CachedValue).IsEqualTo(2);
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// #490, as the issue gives it: a formula XLibur does not evaluate stopped the recalculation of
    /// the whole workbook with <see cref="NotImplementedException"/>.
    /// </summary>
    [Test]
    public async Task Issue490_RecalculateAllFormulas_does_not_throw_on_a_DDE_formula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Unsupported;

        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<NotImplementedException>();
    }

    /// <summary>
    /// #489 and #490. Recalculation leaves the cell dirty with no value, together with every formula
    /// that depends on it, and calculates the rest, as it does for a cycle (Q23). Reading the cell, or
    /// a formula that depends on it, still fails as it did.
    /// </summary>
    [Test]
    [Arguments(Failure.Refused)]
    [Arguments(Failure.Unsupported)]
    public async Task RecalculateAllFormulas_leaves_the_cell_dirty_and_calculates_the_rest(Failure failure)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = FormulaFor(failure);
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = "2*3";

        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("C1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("C1").CachedValue).IsEqualTo(6);
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A1").CachedValue.IsBlank).IsTrue();
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await AssertReadFails(ws.Cell("A1"), failure);
        await AssertReadFails(ws.Cell("B1"), failure);
    }

    [Test]
    [Arguments(Failure.Refused)]
    [Arguments(Failure.Unsupported)]
    public async Task Worksheet_RecalculateAllFormulas_leaves_the_cell_dirty_and_calculates_the_rest(Failure failure)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = FormulaFor(failure);
        ws.Cell("B1").FormulaA1 = "2*3";

        ws.RecalculateAllFormulas();

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("B1").CachedValue).IsEqualTo(6);
    }

    /// <summary>
    /// A workbook holding both kinds and ordinary formulas recalculates. The ordinary formulas, on the
    /// same sheet and on another, get their values. The two cells, and the formulas that depend on
    /// them, are left dirty, and each cell still fails as it did when read. A write afterwards still
    /// reaches its dependents.
    /// </summary>
    [Test]
    public async Task A_refused_formula_and_an_unsupported_feature_leave_their_neighbours_to_calculate()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var other = wb.AddWorksheet("Sheet2");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("A2").FormulaA1 = Unsupported;
        ws.Cell("B1").FormulaA1 = "2*3";
        ws.Cell("B2").FormulaA1 = "B1+1";
        ws.Cell("C1").Value = 5;
        ws.Cell("C2").FormulaA1 = "SUM(C1,B2)";
        ws.Cell("D1").FormulaA1 = "A1+1";
        ws.Cell("D2").FormulaA1 = "A2+1";
        other.Cell("A1").FormulaA1 = "Sheet1!B2*10";

        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("B1").CachedValue).IsEqualTo(6);
        await Assert.That(ws.Cell("B2").CachedValue).IsEqualTo(7);
        await Assert.That(ws.Cell("C2").CachedValue).IsEqualTo(12);
        await Assert.That(other.Cell("A1").CachedValue).IsEqualTo(70);
        foreach (var address in new[] { "B1", "B2", "C2" })
            await Assert.That(ws.Cell(address).NeedsRecalculation).IsFalse().Because($"{address} is ordinary");
        await Assert.That(other.Cell("A1").NeedsRecalculation).IsFalse();

        foreach (var address in new[] { "A1", "A2", "D1", "D2" })
            await Assert.That(ws.Cell(address).NeedsRecalculation).IsTrue().Because($"{address} is, or depends on, a cell that cannot be calculated");
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<ExpressionParseException>();
        await Assert.That(() => _ = ws.Cell("A2").Value).Throws<NotImplementedException>();

        ws.Cell("C1").Value = 1;
        await Assert.That(ws.Cell("C2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(8);
    }

    /// <summary>
    /// #489 and #490. A workbook Excel opens now opens with recalculate-on-load. The cell is saved with
    /// no cached value (ADR 0001), and after the load it is dirty and fails when read, as before.
    /// </summary>
    [Test]
    [Arguments(Failure.Refused)]
    [Arguments(Failure.Unsupported)]
    public async Task A_saved_workbook_opens_with_recalculate_on_load(Failure failure)
    {
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").FormulaA1 = FormulaFor(failure);
            ws.Cell("B1").FormulaA1 = "A1+1";
            ws.Cell("C1").Value = 4;
            ws.Cell("D1").FormulaA1 = "C1*2";
            wb.RecalculateAllFormulas();
            wb.SaveAs(stream);
        }

        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "A1")).IsEqualTo("A1 has no <v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "B1")).IsEqualTo("B1 has no <v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "D1")).IsEqualTo("D1 <v>8</v>");

        stream.Position = 0;
        using var loaded = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
        var sheet = loaded.Worksheet("Sheet1");

        await Assert.That(sheet.Cell("A1").FormulaA1).IsEqualTo(FormulaFor(failure));
        await Assert.That(sheet.Cell("D1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("D1").CachedValue).IsEqualTo(8);
        await Assert.That(sheet.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(sheet.Cell("B1").NeedsRecalculation).IsTrue();
        await AssertReadFails(sheet.Cell("A1"), failure);
    }

    /// <summary>
    /// ADR 0001, where a save falls back to a full recalculation. With a refused formula anywhere in
    /// the workbook, that recalculation threw while it built the dependency tree, so a formula whose
    /// precedents were still dirty when it was written got no cached value. Now only the cell, and the
    /// formula that depends on it, have none.
    /// </summary>
    [Test]
    [Arguments(Failure.Refused)]
    [Arguments(Failure.Unsupported)]
    public async Task Save_writes_no_cached_value_for_the_cell_and_writes_the_rest(Failure failure)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = FormulaFor(failure);
        ws.Cell("B1").FormulaA1 = "A1+1";

        // Calculable, but only by a full recalculation: C1 is written while D1 is still dirty,
        // because D1 needs E1, which is dirty too.
        ws.Cell("C1").FormulaA1 = "D1*2";
        ws.Cell("D1").FormulaA1 = "E1+1";
        ws.Cell("E1").FormulaA1 = "4";

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "A1")).IsEqualTo("A1 has no <v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "B1")).IsEqualTo("B1 has no <v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "C1")).IsEqualTo("C1 <v>10</v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "D1")).IsEqualTo("D1 <v>5</v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "E1")).IsEqualTo("E1 <v>4</v>");
    }

    /// <summary>
    /// Review finding 2. A read whose cell needs a dirty precedent falls back to a full
    /// recalculation. That pass met the cell recalculate-on-load had left dirty, could not calculate
    /// it, and the read threw its failure, although the cell read does not depend on it. The pass now
    /// leaves the cell dirty, as recalculation does, as #492 did for a cycle.
    /// </summary>
    [Test]
    [Arguments(Failure.Refused)]
    [Arguments(Failure.Unsupported)]
    public async Task A_read_does_not_throw_for_a_failure_the_cell_does_not_depend_on(Failure failure)
    {
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").FormulaA1 = FormulaFor(failure);
            ws.Cell("C2").FormulaA1 = "C1*2";
            ws.Cell("E1").FormulaA1 = "C2+1";
            wb.SaveAs(stream);
        }

        stream.Position = 0;
        using var loaded = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
        var sheet = loaded.Worksheet("Sheet1");
        await Assert.That(sheet.Cell("A1").NeedsRecalculation).IsTrue();

        sheet.Cell("C1").Value = 3;

        await Assert.That(sheet.Cell("E1").Value).IsEqualTo(7);
        await Assert.That(sheet.Cell("A1").NeedsRecalculation).IsTrue();
        await AssertReadFails(sheet.Cell("A1"), failure);
    }

    /// <summary>
    /// Review finding 2, the other side. A cell that depends on the failing cell, directly or through
    /// another formula, cannot be calculated either, so reading it still fails as it did.
    /// </summary>
    [Test]
    [Arguments(Failure.Refused)]
    [Arguments(Failure.Unsupported)]
    public async Task A_read_behind_the_failure_still_fails(Failure failure)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = FormulaFor(failure);
        ws.Cell("B1").FormulaA1 = "A1+C1";
        ws.Cell("C1").FormulaA1 = "2";
        ws.Cell("D1").FormulaA1 = "B1*2";

        await AssertReadFails(ws.Cell("D1"), failure);
        await AssertReadFails(ws.Cell("B1"), failure);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2);
    }

    /// <summary>
    /// Review finding 2. With both kinds in the workbook, a read fails with the failure on its own
    /// precedents, not with the first one the pass meets.
    /// </summary>
    [Test]
    public async Task A_read_fails_with_the_failure_it_depends_on()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("A2").FormulaA1 = Unsupported;
        ws.Cell("D1").FormulaA1 = "A1+1";
        ws.Cell("D2").FormulaA1 = "A2+1";

        await Assert.That(() => _ = ws.Cell("D2").Value).Throws<NotImplementedException>();
        await Assert.That(() => _ = ws.Cell("D1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>What makes a formula's precedents unknown.</summary>
    public enum UnknownPrecedents
    {
        RefusedFormula,
        RefusedName,
    }

    /// <summary>
    /// Review finding. A load gives a formula with a cached value that value and leaves it clean, and
    /// Excel's value is what a caller reads until something changes. A formula whose precedents are
    /// unknown is taken to depend on every cell, so before any edit it still reads its cached value.
    /// </summary>
    [Test]
    [Arguments(UnknownPrecedents.RefusedFormula)]
    [Arguments(UnknownPrecedents.RefusedName)]
    public async Task A_loaded_formula_with_unknown_precedents_reads_its_cached_value_before_any_edit(UnknownPrecedents unknown)
    {
        using var wb = LoadedWithUnknownPrecedents(unknown);
        var ws = wb.Worksheet("Sheet1");

        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(10);
    }

    /// <summary>
    /// Review finding. The tree gave such a formula no precedents, and a load left it clean, so after
    /// an edit it may read, nothing marked it dirty: it kept returning its old cached value. Any edit
    /// now marks it dirty, with what depends on it, and reading it fails as a refused formula does
    /// (spec 56). An edit to a cell it does not read does the same: its precedents are unknown, so it
    /// is taken to read every cell.
    /// </summary>
    /// <remarks>
    /// <c>Ext+A2</c> has one known precedent, A2, so an edit to A2 marked it dirty before this fix
    /// too. D9 is the edit that needs the rule for a refused name.
    /// </remarks>
    [Test]
    [Arguments(UnknownPrecedents.RefusedFormula, "A2")]
    [Arguments(UnknownPrecedents.RefusedFormula, "D9")]
    [Arguments(UnknownPrecedents.RefusedName, "A2")]
    [Arguments(UnknownPrecedents.RefusedName, "D9")]
    public async Task After_an_edit_a_loaded_formula_with_unknown_precedents_fails_when_read(
        UnknownPrecedents unknown, string edited)
    {
        using var wb = LoadedWithUnknownPrecedents(unknown);
        var ws = wb.Worksheet("Sheet1");
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2);

        ws.Cell(edited).Value = 5;

        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("E1").NeedsRecalculation).IsTrue();
        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// Review finding. A save wrote the old cached value as current, even when it evaluated formulas.
    /// After an edit the formula is dirty, so it is written with no cached value (ADR 0001).
    /// </summary>
    [Test]
    [Arguments(UnknownPrecedents.RefusedFormula, "A2")]
    [Arguments(UnknownPrecedents.RefusedFormula, "D9")]
    [Arguments(UnknownPrecedents.RefusedName, "A2")]
    [Arguments(UnknownPrecedents.RefusedName, "D9")]
    public async Task A_save_after_an_edit_writes_no_cached_value_for_a_loaded_formula_with_unknown_precedents(
        UnknownPrecedents unknown, string edited)
    {
        using var wb = LoadedWithUnknownPrecedents(unknown);
        var ws = wb.Worksheet("Sheet1");
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2);
        ws.Cell(edited).Value = 5;

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "B1")).IsEqualTo("B1 has no <v>");
    }

    /// <summary>
    /// Review finding. A save writes each in-cell image's rich-value index into the image cell's
    /// value, through the setter an edit uses, so the save counted as an edit: every formula whose
    /// precedents are unknown was marked dirty. B1 and E1 were written with no cached value, with or
    /// without evaluating formulas, and read as a failure afterwards. The index is the save's own
    /// bookkeeping, not an edit, so it marks nothing dirty.
    /// </summary>
    [Test]
    [Arguments(UnknownPrecedents.RefusedFormula, false)]
    [Arguments(UnknownPrecedents.RefusedFormula, true)]
    [Arguments(UnknownPrecedents.RefusedName, false)]
    [Arguments(UnknownPrecedents.RefusedName, true)]
    public async Task A_save_with_an_in_cell_image_keeps_the_cached_value_of_a_loaded_formula_with_unknown_precedents(
        UnknownPrecedents unknown, bool evaluateFormulas)
    {
        using var wb = LoadedWithUnknownPrecedents(unknown, withInCellImage: true);
        var ws = wb.Worksheet("Sheet1");
        await Assert.That(ws.Cell("G1").HasCellImage).IsTrue();
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2);

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = evaluateFormulas });

        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "B1")).IsEqualTo("B1 <v>10</v>");
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(stream, "E1")).IsEqualTo("E1 <v>20</v>");
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(10);
    }

    /// <summary>
    /// D80, pinned as it is today and left for its own fix. After a load the calc engine has no
    /// dependency tree, and it builds one on an edit only once it has calculated a formula. So on a
    /// freshly loaded workbook an edit marks nothing dirty: a formula whose precedents are unknown
    /// keeps reading its cached value, exactly as the ordinary F1, which reads the edited A2, does.
    /// When D80 is fixed, both are marked dirty by the edit and this test changes.
    /// </summary>
    [Test]
    [Arguments(UnknownPrecedents.RefusedFormula)]
    [Arguments(UnknownPrecedents.RefusedName)]
    public async Task D80_on_a_fresh_load_an_edit_leaves_it_clean_as_it_leaves_an_ordinary_formula(
        UnknownPrecedents unknown)
    {
        using var wb = LoadedWithUnknownPrecedents(unknown);
        var ws = wb.Worksheet("Sheet1");

        ws.Cell("A2").Value = 5;

        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(10);
        await Assert.That(ws.Cell("F1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("F1").Value).IsEqualTo(3);
    }

    /// <summary>
    /// A workbook as a load leaves it. B1 reads A2 through a formula whose precedents are unknown: its
    /// text is refused, or it uses the name <c>Ext</c>, whose text is refused. B1 has Excel's cached
    /// value 10, E1, which reads B1, has 20, and F1, an ordinary formula that reads A2, has 3, so all
    /// three are clean. C1 has no cached value. Reading it is a calculation, after which the calc
    /// engine builds its dependency tree on the next edit. With <paramref name="withInCellImage"/>,
    /// G1, which nothing reads, holds an in-cell image.
    /// </summary>
    private static XLWorkbook LoadedWithUnknownPrecedents(UnknownPrecedents unknown, bool withInCellImage = false)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var values = ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice;
            ws.Cell("A2").Value = 1;
            var b1 = (XLCell)ws.Cell("B1");
            b1.FormulaA1 = unknown == UnknownPrecedents.RefusedFormula ? Refused + "+A2" : "Ext+A2";
            values.SetCellValue(b1.SheetPoint, 10);
            b1.Formula!.MarkClean();
            var e1 = (XLCell)ws.Cell("E1");
            e1.FormulaA1 = "B1*2";
            values.SetCellValue(e1.SheetPoint, 20);
            e1.Formula!.MarkClean();
            var f1 = (XLCell)ws.Cell("F1");
            f1.FormulaA1 = "A2*3";
            values.SetCellValue(f1.SheetPoint, 3);
            f1.Formula!.MarkClean();
            ws.Cell("C1").FormulaA1 = "1+1";
            if (withInCellImage)
            {
                using var image = new MemoryStream(XLCellImageTests.CreateTestPng());
                ws.Cell("G1").SetCellImage(image, XLPictureFormat.Png);
            }

            wb.SaveAs(package);
        }

        if (unknown == UnknownPrecedents.RefusedName)
            package = package.RewriteWorkbook(WithRefusedDefinedName);

        package.Position = 0;
        return new XLWorkbook(package);
    }

    private static string FormulaFor(Failure failure) => failure == Failure.Refused ? Refused : Unsupported;

    /// <summary>
    /// A workbook whose name <c>Ext</c> holds <see cref="Refused"/>, and whose B1 is <c>Ext+1</c>. Only
    /// a load can give a name such text: <see cref="IXLDefinedName.RefersTo"/> rejects it.
    /// </summary>
    private static MemoryStream BookWithRefusedDefinedName()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1").Cell("B1").FormulaA1 = "Ext+1";
            wb.SaveAs(package);
        }

        return package.RewriteWorkbook(WithRefusedDefinedName);
    }

    /// <summary>Splices a name <c>Ext</c> holding <see cref="Refused"/> into a workbook part.</summary>
    private static string WithRefusedDefinedName(string workbookXml)
    {
        var refersTo = Refused.Replace("'", "&apos;", StringComparison.Ordinal);
        var rewritten = workbookXml.Replace("<x:definedNames />",
            $"<x:definedNames><x:definedName name=\"Ext\">{refersTo}</x:definedName></x:definedNames>");
        if (ReferenceEquals(rewritten, workbookXml))
            throw new InvalidOperationException("The defined name was not spliced into the workbook part.");

        return rewritten;
    }

    /// <summary>What reading the cell throws for <paramref name="failure"/>, unchanged by #489 and #490.</summary>
    private static async Task AssertReadFails(IXLCell cell, Failure failure)
    {
        if (failure == Failure.Refused)
            await Assert.That(() => _ = cell.Value).Throws<ExpressionParseException>();
        else
            await Assert.That(() => _ = cell.Value).Throws<NotImplementedException>();
    }
}
