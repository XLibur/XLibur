using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
// Phase B1 — in-memory spilling of dynamic-array formulas.
// A dynamic-array formula in a single anchor cell auto-fills its computed footprint into the
// neighbouring cells, which stay formula-less. A blocked footprint (existing content) or one that
// runs past the sheet edge collapses to a #SPILL! error on the anchor. Re-evaluating to a smaller
// footprint clears the cells the previous result no longer covers.
public class SpillEvaluationTests
{
    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return wb.AddWorksheet("Sheet1");
    }

    [Test]
    public async Task Spill_ColumnVector_FillsFootprint()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");

            // Reading the anchor evaluates the formula and spills into A2:A3.
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Spill_TwoDimensional_FillsGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(2, 3)");

            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(4);
            await Assert.That(ws.Cell("B2").Value).IsEqualTo(5);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(6);
        }
    }

    [Test]
    public async Task Spill_AnchorHoldsFormula_SpilledCellsDoNot()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1); // trigger the spill

            await Assert.That(ws.Cell("A1").HasFormula).IsTrue().Because("Anchor must keep the formula");
            await Assert.That(ws.Cell("A2").HasFormula).IsFalse().Because("Spilled cell must be formula-less");
            await Assert.That(ws.Cell("A3").HasFormula).IsFalse().Because("Spilled cell must be formula-less");
        }
    }

    [Test]
    public async Task Spill_BlockedByExistingValue_ProducesSpillErrorAndPreservesBlocker()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");

            // The footprint A1:A3 collides with A2, so only the anchor is written.
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.SpillRange);
            await Assert.That(ws.Cell("A2").Value).IsEqualTo("block").Because("Blocking value must be untouched");
            await Assert.That(ws.Cell("A3").IsEmpty()).IsTrue().Because("No value is written to blocked-spill cells");
        }
    }

    [Test]
    public async Task Spill_BlockedByFormula_ProducesSpillError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A3").FormulaA1 = "1+1";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");

            await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.SpillRange);
            await Assert.That(ws.Cell("A2").IsEmpty()).IsTrue().Because("No value is written to blocked-spill cells");
        }
    }

    [Test]
    public async Task Spill_ShrinkingResult_ClearsStaleCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 2;
            ws.Cell("B3").Value = 3;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(B1:B3)");

            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1); // spills A1:A3 = {1;2;3}
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);

            // Collapse a source value so only two distinct values remain: the same formula
            // instance now spills A1:A2 only and must clear the stale A3.
            ws.Cell("B3").Value = 2;
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("A3").IsEmpty()).IsTrue().Because("Stale cell of the previous footprint must be cleared");
        }
    }

    [Test]
    public async Task Spill_DependentOfSpilledCell_RecalculatesWhenSourceChanges()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)"); // spills A1:A3 = {1;2;3}
            ws.Cell("C1").FormulaA1 = "A3*10";                  // depends on the spilled A3

            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(30);

            // Change a source cell so the spilled A3 becomes 5.
            ws.Cell("D3").Value = 5;
            await Assert.That(ws.Cell("C1").NeedsRecalculation).IsTrue().Because("A dependent of a spilled (non-anchor) cell must be invalidated when the array's source changes");

            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(50);
        }
    }

    [Test]
    public async Task Spill_DependentBeforeAnchor_RecalculatesAfterInitialSpill()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            // Anchor at C1 spills C1:C3; the dependent at A1 sits positionally BEFORE the anchor
            // and reads the spilled (non-anchor) C3 directly.
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            ws.Cell("A1").FormulaA1 = "C3*10";

            // Establish the spill so the spill-owner lookup knows C3 belongs to C1.
            wb.CalcEngine.Recalculate(wb, null);

            // A later source change must recompute the dependent with the fresh spilled value:
            // reading C3 now forces the dirty anchor C1 to evaluate first.
            ws.Cell("D3").Value = 5;
            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(50);
        }
    }

    [Test]
    [Skip("Remaining limitation: on the VERY FIRST evaluation the spill footprint is unknown " +
            "until the anchor runs, so a dependent positioned before a not-yet-spilled anchor still " +
            "reads a blank cell. Full ordering here needs a calc-chain pre-pass that sizes arrays " +
            "before evaluation. Post-first-spill ordering is covered by " +
            nameof(Spill_DependentBeforeAnchor_RecalculatesAfterInitialSpill) + ".")]
    public async Task Spill_DependentBeforeAnchor_FirstEvaluationOrdering()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            ws.Cell("A1").FormulaA1 = "C3*10";

            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(30);
        }
    }

    [Test]
    public async Task SpillOperator_ReferencesWholeFootprint()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)"); // spills A1:A3 = {1;2;3}
            ws.Cell("C1").FormulaA1 = "SUM(A1#)";

            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(6);
        }
    }

    [Test]
    public async Task SpillOperator_NonAnchorCell_ReturnsRefError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // A1 holds no dynamic array, so A1# is #REF!.
            await Assert.That(ws.Evaluate("A1#")).IsEqualTo(XLError.CellReference);
        }
    }

    [Test]
    public async Task SpillOperator_MultiCellOperand_ReturnsRefError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // The spill operator requires a single-cell anchor; a multi-cell operand is #REF!.
            wb.DefinedNames.Add("Rng", "Sheet1!A1:B2");
            await Assert.That(ws.Evaluate("Rng#")).IsEqualTo(XLError.CellReference);
        }
    }

    [Test]
    public async Task SpillOperator_TracksFootprintWhenItShrinks()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)"); // spills A1:A3
            ws.Cell("C1").FormulaA1 = "SUM(A1#)";

            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(6); // 1+2+3

            // Collapse to two distinct values: A1# now covers A1:A2 only.
            ws.Cell("D3").Value = 1;
            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(3); // 1+2
        }
    }

    [Test]
    public async Task SpillOperator_EvaluatesAnchorFirst_EvenWhenDependentComesBefore()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // The dependent A1=SUM(C1#) sits positionally BEFORE its anchor C1. Because the
            // spill operator's range includes the anchor cell (which holds the dirty formula),
            // reading it forces the anchor to evaluate first — so this orders correctly, unlike
            // a plain read of a non-anchor spilled cell.
            ws.Cell("C1").SetDynamicFormulaA1("SEQUENCE(3)"); // spills C1:C3
            ws.Cell("A1").FormulaA1 = "SUM(C1#)";

            wb.CalcEngine.Recalculate(wb, null);
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(6);
        }
    }

    [Test]
    public async Task Spill_SurvivesRowInsertAndReSpills()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)"); // spills A1:A3
            wb.RecalculateAllFormulas();
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);

            // A structural insert relocates the anchor A1 -> A2. It stays a dynamic array and
            // re-spills over the shifted footprint A2:A4 after recalculation.
            ws.Row(1).InsertRowsAbove(1);
            wb.RecalculateAllFormulas();

            await Assert.That(ws.Cell("A2").HasFormula).IsTrue().Because("Anchor must stay dynamic after the shift");
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A4").Value).IsEqualTo(3);
            await Assert.That(ws.Cell("A3").HasFormula).IsFalse().Because("Spilled cell stays formula-less after the shift");
        }
    }

    [Test]
    public async Task Spill_SurvivesRowDeleteAndReSpills()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A2").SetDynamicFormulaA1("SEQUENCE(3)"); // spills A2:A4
            wb.RecalculateAllFormulas();
            await Assert.That(ws.Cell("A4").Value).IsEqualTo(3);

            // Deleting the empty row above relocates the anchor A2 -> A1 (negative shift); it
            // re-spills over A1:A3.
            ws.Row(1).Delete();
            wb.RecalculateAllFormulas();

            await Assert.That(ws.Cell("A1").HasFormula).IsTrue().Because("Anchor must stay dynamic after the delete");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);
            await Assert.That(ws.Cell("A2").HasFormula).IsFalse().Because("Spilled cell stays formula-less after the delete");
        }
    }

    [Test]
    public async Task Spill_PastSheetEdge_ProducesSpillError()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            // Anchor on the last row: a 2-row result would need a row beyond the sheet.
            var anchor = ws.Cell(XLHelper.MaxRowNumber, 1);
            anchor.SetDynamicFormulaA1("SEQUENCE(2)");

            await Assert.That(anchor.Value).IsEqualTo(XLError.SpillRange);
        }
    }

    [Test]
    public async Task Spill_HorizontalVector_FillsRow()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(1, 3)"); // spills A1:C1

            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Spill_GrowingResult_FillsNewCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 2;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)"); // {1;2} -> A1:A2

            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1); // trigger the spill
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("A3").IsEmpty()).IsTrue().Because("Only two distinct values initially");

            // A third distinct value grows the footprint into the previously-empty A3.
            ws.Cell("D3").Value = 3;
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Spill_ErrorIsReportedByErrorFunctions()
    {
        // A real #SPILL! cell reports through ERROR.TYPE (9) and ISERROR — exercising the
        // XLError.SpillRange enum member end to end (the literal can't be parsed).
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.SpillRange);

            ws.Cell("C1").FormulaA1 = "ERROR.TYPE(A1)";
            ws.Cell("C2").FormulaA1 = "ISERROR(A1)";
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(9);
            await Assert.That(ws.Cell("C2").Value).IsEqualTo(ExpectedCellValue.From(true));
        }
    }

    [Test]
    public async Task Spill_RecoversAfterBlockerClearedAndAnchorReevaluates()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.SpillRange);

            // Clear the blocker and change a source so the anchor re-evaluates: the spill recovers.
            ws.Cell("A2").Value = Blank.Value;
            ws.Cell("D3").Value = 4;
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(4);
        }
    }

    [Test]
    public async Task Spill_SurvivesColumnInsertAndReSpills()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(1, 3)"); // spills A1:C1
            wb.RecalculateAllFormulas();
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);

            // A column insert relocates the anchor A1 -> B1; it re-spills over B1:D1.
            ws.Column(1).InsertColumnsBefore(1);
            wb.RecalculateAllFormulas();

            await Assert.That(ws.Cell("B1").HasFormula).IsTrue().Because("Anchor must stay dynamic after the shift");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Spill_SurvivesColumnDeleteAndReSpills()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("B1").SetDynamicFormulaA1("SEQUENCE(1, 3)"); // spills B1:D1
            wb.RecalculateAllFormulas();
            await Assert.That(ws.Cell("D1").Value).IsEqualTo(3);

            // Deleting the empty column to the left relocates the anchor B1 -> A1 (negative
            // shift); it re-spills over A1:C1.
            ws.Column(1).Delete();
            wb.RecalculateAllFormulas();

            await Assert.That(ws.Cell("A1").HasFormula).IsTrue().Because("Anchor must stay dynamic after the delete");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);
        }
    }

    [Test]
    public async Task Spill_DependentBeforeAnchor_OrdersCorrectlyOnInteractiveRead()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("D1").Value = 1;
            ws.Cell("D2").Value = 2;
            ws.Cell("D3").Value = 3;
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            ws.Cell("A1").FormulaA1 = "C3*10";

            // Establish the spill by reading the anchor, then the dependent.
            await Assert.That(ws.Cell("C1").Value).IsEqualTo(1);
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(30);

            // A plain .Value read of the dependent after a source change must order the dirty
            // anchor first (via the fallback to a full, dependency-ordered recalculation).
            ws.Cell("D3").Value = 5;
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(50);
        }
    }

    /// <summary>
    /// Excel reads a shared-string or inline-string cell inside a spill footprint as content
    /// occupying the range, and renders every spilled cell below the anchor as <c>#VALUE!</c>. The
    /// spilled cells therefore have to be saved the way Excel saves them — as cached formula results
    /// (<c>t="str"</c> with a <c>&lt;v&gt;</c>), even though they carry no formula of their own.
    /// </summary>
    [Test]
    public async Task Spill_TextFootprint_SavedAsFormulaResultsNotSharedStrings()
    {
        using var ms = new MemoryStream();
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "alpha";
            ws.Cell("A2").Value = "beta";
            ws.Cell("A3").Value = "alpha";
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(A1:A3)");

            await Assert.That(ws.Cell("C1").Value).IsEqualTo("alpha");
            await Assert.That(ws.Cell("C2").Value).IsEqualTo("beta");

            wb.SaveAs(ms, validate: false);
        }

        var sheetXml = ReadSheetXml(ms);

        // The anchor keeps its formula; the spilled cell holds only the cached result, but both are
        // typed as formula results.
        await Assert.That(CellXml(sheetXml, "C1")).Contains(@"t=""str""").And.Contains(@"cm=""1""");
        await Assert.That(CellXml(sheetXml, "C2")).IsEqualTo(@"<x:c r=""C2"" s=""0"" t=""str""><x:v>beta</x:v></x:c>");
    }

    /// <summary>
    /// A dynamic array that has never been evaluated has no footprint at all until the save itself
    /// spills it. If that happens part-way through the write pass, the spilled cells land behind the
    /// enumerator and never reach the file, leaving the anchor claiming a <c>ref</c> the file has no
    /// cells for; and the footprint the pass started with does not cover them either, so any that do
    /// get written go out as constants.
    /// </summary>
    [Test]
    public async Task Spill_TextFootprint_EvaluatedDuringSave_IsWrittenAsFormulaResults()
    {
        using var ms = new MemoryStream();
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "alpha";
            ws.Cell("A2").Value = "beta";
            ws.Cell("A3").Value = "alpha";

            // Deliberately never read: the spill has to happen during the save.
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(A1:A3)");

            wb.SaveAs(ms, new SaveOptions { EvaluateFormulasBeforeSaving = true, ValidatePackage = false });
        }

        var sheetXml = ReadSheetXml(ms);

        await Assert.That(CellXml(sheetXml, "C1")).Contains(@"ref=""C1:C2""");
        await Assert.That(CellXml(sheetXml, "C2"))
            .IsEqualTo(@"<x:c r=""C2"" s=""0"" t=""str""><x:v>beta</x:v></x:c>")
            .Because("the cell the anchor's ref promises must exist, and be typed as a formula result");
    }

    /// <summary>
    /// The same footprint typing has to survive a load/save round trip: a spilled cell arrives with
    /// <c>t="str"</c> and no <c>&lt;f&gt;</c>, so nothing but the anchor's footprint marks it as part
    /// of the array. Reading it as an ordinary text constant is what turned a spill into #VALUE!.
    /// </summary>
    [Test]
    public async Task Spill_TextFootprint_SurvivesRoundTrip()
    {
        using var first = new MemoryStream();
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "alpha";
            ws.Cell("A2").Value = "beta";
            ws.Cell("A3").Value = "alpha";
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(A1:A3)");
            await Assert.That(ws.Cell("C1").Value).IsEqualTo("alpha");
            await Assert.That(ws.Cell("C2").Value).IsEqualTo("beta");

            wb.SaveAs(first, validate: false);
        }

        first.Position = 0;
        using var second = new MemoryStream();
        using (var reloaded = new XLWorkbook(first))
        {
            var sheet = reloaded.Worksheets.First();
            // The _xlfn. prefix is how a future function is stored in the file, and XLibur keeps it
            // on the loaded text; the point here is that the anchor is still a dynamic array.
            await Assert.That(sheet.Cell("C1").FormulaA1).EndsWith("UNIQUE(A1:A3)");
            await Assert.That(sheet.Cell("C2").Value).IsEqualTo("beta");

            reloaded.SaveAs(second, validate: false);
        }

        var sheetXml = ReadSheetXml(second);
        await Assert.That(CellXml(sheetXml, "C2")).IsEqualTo(@"<x:c r=""C2"" s=""0"" t=""str""><x:v>beta</x:v></x:c>");
    }

    private static string ReadSheetXml(MemoryStream savedWorkbook)
    {
        using var zip = new ZipArchive(new MemoryStream(savedWorkbook.ToArray()), ZipArchiveMode.Read);
        var sheetEntry = zip.Entries.First(e => e.FullName.Contains("sheet1.xml", StringComparison.OrdinalIgnoreCase));
        using var reader = new StreamReader(sheetEntry.Open());
        return reader.ReadToEnd();
    }

    private static string CellXml(string sheetXml, string cellRef)
    {
        var match = Regex.Match(sheetXml, $@"<x:c r=""{cellRef}"".*?(?:/>|</x:c>)", RegexOptions.Singleline);
        return match.Success ? match.Value : $"<missing cell {cellRef}>";
    }
}
