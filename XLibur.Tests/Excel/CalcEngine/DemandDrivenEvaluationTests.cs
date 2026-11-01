using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Spec 04 (#841). Reading a dirty formula calculates the precedents it reads, and only those. It
/// does not build the dependency tree or the calculation chain, and does not fall back to a pass
/// over the whole workbook.
/// </summary>
public class DemandDrivenEvaluationTests
{
    private const string SheetName = "Sheet1";

    [Test]
    public async Task Reading_a_formula_over_plain_values_builds_no_tree_and_no_chain()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 20;
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = "A1+2";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(21);

        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
        await Assert.That(wb.CalcEngine.HasCalculationChain).IsFalse();

        // The formula nobody read is left alone.
        await Assert.That(ws.Cell("C1").NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task Reading_a_formula_calculates_its_dirty_precedents_without_a_pass()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").FormulaA1 = "A1+1";
        ws.Cell("A3").FormulaA1 = "A2+1";
        ws.Cell("A4").FormulaA1 = "A3+1";
        ws.Cell("B1").FormulaA1 = "A1*100";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(4);

        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
        await Assert.That(ws.Cell("A2").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A3").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task A_range_function_calculates_every_dirty_formula_in_the_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        for (var row = 1; row <= 5; row++)
            ws.Cell(row, 1).FormulaA1 = $"{row}*10";

        ws.Cell("B1").FormulaA1 = "SUM(A:A)";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(150);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_precedent_on_another_sheet_is_calculated_on_demand()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");
        second.Cell("A1").Value = 5;
        second.Cell("A2").FormulaA1 = "A1*2";
        first.Cell("A1").FormulaA1 = "Second!A2+1";

        await Assert.That(first.Cell("A1").Value).IsEqualTo(11);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(second.Cell("A2").NeedsRecalculation).IsFalse();
    }

    [Test]
    public async Task A_defined_name_over_a_dirty_cell_is_calculated_on_demand()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "7*6";
        wb.DefinedNames.Add("Answer", $"{SheetName}!$A$1");
        ws.Cell("B1").FormulaA1 = "Answer+0";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(42);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_cycle_throws_for_its_cells_and_not_for_a_cell_that_does_not_use_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "B1+1";
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = "A1+5";
        ws.Cell("D1").FormulaA1 = "2+2";

        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(() => _ = ws.Cell("C1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(4);

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_formula_that_reads_itself_is_a_cycle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "A1+1";

        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task A_cycle_through_a_range_is_found()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "SUM(B1:B3)";
        ws.Cell("B2").FormulaA1 = "A1";

        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
    }

    [Test]
    public async Task A_cycle_is_found_again_when_read_again()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "B1";
        ws.Cell("B1").FormulaA1 = "A1";

        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();

        // Breaking the cycle makes it calculable.
        ws.Cell("B1").Value = 9;
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(9);
    }

    [Test]
    public async Task An_edit_after_a_read_dirties_the_formulas_that_read_the_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").FormulaA1 = "A1+1";
        ws.Cell("A3").FormulaA1 = "A2+1";

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);

        ws.Cell("A1").Value = 10;

        await Assert.That(ws.Cell("A3").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(12);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(11);
    }

    [Test]
    public async Task A_read_does_not_clean_a_formula_that_a_later_edit_must_reach()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").FormulaA1 = "A1*2";
        ws.Cell("A3").FormulaA1 = "A2*2";

        _ = ws.Cell("A2").Value;
        ws.Cell("A1").Value = 5;

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(20);
    }

    [Test]
    public async Task A_formula_that_cannot_be_parsed_stops_only_the_formulas_that_read_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "1+";
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = "40+2";

        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<ExpressionParseException>();
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(42);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_volatile_precedent_is_calculated_once_for_the_formulas_that_read_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "RAND()";
        ws.Cell("B1").FormulaA1 = "A1";
        ws.Cell("C1").FormulaA1 = "A1";

        var b = ws.Cell("B1").Value;
        var c = ws.Cell("C1").Value;

        await Assert.That(b).IsEqualTo(c);
        await Assert.That(b).IsEqualTo(ws.Cell("A1").Value);
    }

    [Test]
    public async Task A_cell_in_the_middle_of_a_dirty_spill_range_calculates_the_owner()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("C1").Value = 1;
        ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)*C1");
        ws.Cell("B1").FormulaA1 = "A3+0";

        // The footprint of an owner is known once the owner has been calculated.
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(3);

        // The edit dirties the owner and the cell that reads its spilled A3.
        ws.Cell("C1").Value = 10;
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(30);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(20);
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsFalse();
    }

    [Test]
    public async Task A_spilled_cell_is_current_after_a_structural_edit_dirtied_every_formula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("C1").Value = 1;
        ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)*C1");
        ws.Cell("B1").FormulaA1 = "A3+0";
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);

        // A row insert drops the engine's footprints and marks every formula dirty, with no tree.
        ws.Row(100).InsertRowsBelow(1);
        ws.Cell("C1").Value = 10;

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(30);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
    }

    [Test]
    public async Task A_range_over_a_dirty_spill_owner_sees_the_spilled_cells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("C1").Value = 1;
        ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(4)*C1");
        ws.Cell("B1").FormulaA1 = "SUM(A2:A4)";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(9);

        ws.Cell("C1").Value = 2;
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(18);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task Reading_a_hundred_formulas_of_a_hundred_thousand_calculates_only_their_closure()
    {
        const int count = 100_000;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        for (var row = 1; row <= count; row++)
        {
            ws.Cell(row, 1).Value = row;
            ws.Cell(row, 2).FormulaA1 = $"A{row}+1";
            ws.Cell(row, 3).FormulaA1 = $"B{row}*2";
        }

        var step = count / 100;
        for (var i = 0; i < 100; i++)
        {
            var row = 1 + i * step;
            await Assert.That(ws.Cell(row, 3).Value).IsEqualTo((row + 1) * 2);
        }

        // Each read calculates the formula and the one precedent behind it.
        await Assert.That(wb.CalcEngine.DemandEvaluationCount).IsEqualTo(200);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
        await Assert.That(wb.CalcEngine.HasCalculationChain).IsFalse();
    }

    [Test]
    public async Task A_very_long_chain_of_dirty_formulas_is_still_calculated()
    {
        const int length = 30_000;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell(1, 1).Value = 1;
        for (var row = 2; row <= length; row++)
            ws.Cell(row, 1).FormulaA1 = $"A{row - 1}+1";

        await Assert.That(ws.Cell(length, 1).Value).IsEqualTo(length);
        await Assert.That(ws.Cell(length / 2, 1).NeedsRecalculation).IsFalse();
    }

    [Test]
    public async Task Loading_and_reading_formulas_without_a_cached_value_builds_no_tree()
    {
        using var stream = new MemoryStream();
        using (var source = new XLWorkbook())
        {
            var ws = source.AddWorksheet(SheetName);
            ws.Cell("A1").Value = 3;
            ws.Cell("A2").FormulaA1 = "A1*2";
            ws.Cell("A3").FormulaA1 = "A2*2";
            source.SaveAs(stream);
        }

        stream.Position = 0;
        using var wb = new XLWorkbook(stream);
        var loaded = wb.Worksheet(SheetName);

        await Assert.That(loaded.Cell("A3").Value).IsEqualTo(12);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
        await Assert.That(wb.CalcEngine.HasCalculationChain).IsFalse();
    }

    [Test]
    [Arguments("1+")]
    [Arguments("P1+1")]
    public async Task A_save_walks_a_failing_precedent_once_for_a_long_chain_behind_it(string failing)
    {
        const int length = 2000;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("P1").FormulaA1 = failing;
        ws.Cell("B1").FormulaA1 = "P1+1";
        for (var row = 2; row <= length; row++)
            ws.Cell(row, 2).FormulaA1 = $"B{row - 1}+1";

        // A formula outside the failure is still calculated and written.
        ws.Cell("D1").FormulaA1 = "3*3";

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        // Each formula is calculated once: the chain, P1 and D1. Walking the failure again for every
        // dependent would be about length * length / 2.
        await Assert.That(wb.CalcEngine.DemandEvaluationCount).IsLessThanOrEqualTo(length + 4);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(ws.Cell("B" + length).NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("D1").NeedsRecalculation).IsFalse();
    }

    [Test]
    public async Task A_failure_is_not_remembered_outside_a_save()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "B1+1";
        ws.Cell("B1").FormulaA1 = "A1+1";
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();

        ws.Cell("B1").Value = 1;
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(2);
    }

    [Test]
    public async Task A_cycle_deep_in_a_chain_throws_for_the_whole_chain_and_names_the_first_cell()
    {
        const int length = 50;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        for (var row = 1; row < length; row++)
            ws.Cell(row, 1).FormulaA1 = $"A{row + 1}+1";

        ws.Cell(length, 1).FormulaA1 = "A1+1";
        ws.Cell("B1").FormulaA1 = "A25*2";

        var ex = await Assert.That(() => _ = ws.Cell("A1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(ex!.Message).IsEqualTo("Formula in a cell '$Sheet1'!$A1 is part of a cycle.");
        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<XLCircularReferenceException>();
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_legacy_array_formula_over_dirty_cells_is_calculated_on_demand()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "1*1";
        ws.Cell("A2").FormulaA1 = "2*1";
        ws.Cell("A3").FormulaA1 = "3*1";
        ws.Range("B1:B3").FormulaArrayA1 = "A1:A3*2";
        ws.Cell("C1").FormulaA1 = "B2+0";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(4);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("B3").NeedsRecalculation).IsFalse();
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_data_table_precedent_falls_back_to_the_full_recalculation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        var table = XLCellFormula.DataTable1D(new Area(new Point(1, 4)), new Point(1, 1), false, false);
        ((XLWorksheet)ws).Internals.CellsCollection.FormulaSlice.Set(new Point(1, 4), table);
        ws.Cell("B1").FormulaA1 = "D1+1";
        ws.Cell("C1").FormulaA1 = "2+2";

        await Assert.That(() => _ = ws.Cell("B1").Value).Throws<NotImplementedException>();

        // The fallback was the pass of a full recalculation, and it calculated what it could.
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(1);
        await Assert.That(ws.Cell("C1").NeedsRecalculation).IsFalse();
    }

    [Test]
    [Arguments("LET(x,A1,x+1)")]
    [Arguments("LAMBDA(v,v+A1)(1)")]
    [Arguments("LET(f,LAMBDA(v,v+A1),f(1))")]
    public async Task A_let_or_lambda_body_that_reads_a_dirty_cell_calculates_it(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").FormulaA1 = "5*2";
        ws.Cell("B1").FormulaA1 = formula;

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(11);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
    }

    [Test]
    public async Task A_spill_made_by_a_nested_precedent_is_seen_by_a_walk_of_the_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);

        // D1 reads the anchor D40, which spills into D41:D42 while the SUM is walking D1:D100.
        ws.Cell("D1").FormulaA1 = "D40+0";
        ws.Cell("D40").SetDynamicFormulaA1("SEQUENCE(3)");
        ws.Cell("F1").FormulaA1 = "SUM(D1:D100)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(1 + 1 + 2 + 3);
        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(0);
    }

    [Test]
    public async Task A_dirty_anchor_of_a_loaded_workbook_is_calculated_when_a_formula_reads_its_spill()
    {
        using var stream = new MemoryStream();
        using (var source = new XLWorkbook())
        {
            var ws = source.AddWorksheet(SheetName);
            ws.Cell("C1").Value = 1;
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)*C1");
            ws.Cell("E1").FormulaA1 = "A3+0";
            _ = ws.Cell("A1").Value; // the spill is written; E1 is saved with no value
            source.SaveAs(stream);
        }

        stream.Position = 0;
        using var wb = new XLWorkbook(stream);
        var loaded = wb.Worksheet(SheetName);

        // Neither of these builds the dependency tree, so nothing but the engine's own footprints
        // can tell that the spilled A3 is stale.
        loaded.Cell("A1").InvalidateFormula();
        ((XLWorksheet)loaded).Internals.CellsCollection.ValueSlice.SetCellValue(new Point(1, 3), 10);

        await Assert.That(loaded.Cell("E1").Value).IsEqualTo(30);
        await Assert.That(wb.CalcEngine.HasDependencyTree).IsFalse();
    }

    [Test]
    public async Task Recalculate_all_formulas_still_runs_a_full_pass()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SheetName);
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").FormulaA1 = "A1*2";
        ws.Cell("A3").FormulaA1 = "A2*2";

        wb.RecalculateAllFormulas();

        await Assert.That(wb.CalcEngine.PassCount).IsEqualTo(1);
        await Assert.That(ws.Cell("A3").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(8);
    }
}
