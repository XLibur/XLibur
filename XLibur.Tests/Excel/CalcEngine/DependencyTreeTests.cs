using System;
using System.Collections.Generic;
using System.IO;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;
using XLibur.Tests.Excel.IO;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

internal class DependencyTreeTests
{
    #region Add formula to dependency tree

    [Test]
    [MethodDataSource(nameof(AreaDependenciesTestCases))]
    public async Task Area_dependencies_are_extracted_from_formula(string formula, IReadOnlyList<SheetArea> expectedAreas)
    {
        var dependencies = GetDependencies(formula);
        await Assert.That(dependencies.Areas).IsEquivalentTo(expectedAreas);
    }

    [Test]
    [MethodDataSource(nameof(NameDependenciesTestCases))]
    public async Task Name_dependencies_are_kept_for_dependencies_update(string formula, IReadOnlyList<XLName> expectedNames)
    {
        var dependencies = GetDependencies(formula);
        await Assert.That(dependencies.Names).IsEquivalentTo(expectedNames);
    }

    [Test]
    public async Task Name_range_is_added_to_dependencies_of_formula()
    {
        var dependencies = GetDependencies("name + D2", init: wb =>
        {
            wb.DefinedNames.Add("name", "Sheet!$B$4+Sheet!$C$6");
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("D2")),
            new("Sheet", Area.Parse("B4")),
            new("Sheet", Area.Parse("C6"))
        });
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("name")]);
    }

    [Test]
    public async Task Name_range_that_is_reference_is_propagated_to_formula()
    {
        var dependencies = GetDependencies("B3:name", init: wb =>
        {
            wb.DefinedNames.Add("name", "Sheet!$D$7");
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("B3:D7")),
        });
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("name")]);
    }

    [Test]
    public async Task Name_range_can_used_another_name_range()
    {
        var dependencies = GetDependencies("outer", init: wb =>
        {
            wb.DefinedNames.Add("outer", "Sheet!$D$7 + inner");
            wb.DefinedNames.Add("inner", "Sheet!$B$1");
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("D7")),
            new("Sheet", Area.Parse("B1")),
        });
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("outer"), new XLName("inner")]);
    }

    /// <summary>
    /// #489. The parser refuses the text, so the references in it are unknown. The formula gets no
    /// precedents, as a data table's placeholder text gets none, instead of failing the tree it is
    /// added to.
    /// </summary>
    [Test]
    public async Task A_refused_formula_has_no_precedents()
    {
        var dependencies = GetDependencies("'[Book2.xlsx]Sheet1'!A1");
        await Assert.That(dependencies.Areas).IsEmpty();
        await Assert.That(dependencies.Names).IsEmpty();
    }

    /// <summary>
    /// D79. A defined name whose formula refers to itself made the tree follow it for ever and
    /// overflow the stack, which ends the process. A name met again on its own path adds nothing
    /// more; the rest of its formula still does.
    /// </summary>
    [Test]
    public async Task D79_a_name_that_refers_to_itself_is_followed_once()
    {
        var dependencies = GetDependencies("Loop", init: wb => wb.DefinedNames.Add("Loop", "Loop+Sheet!$B$2"));
        await Assert.That(dependencies.Areas).IsEquivalentTo([new SheetArea("Sheet", Area.Parse("B2"))]);
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("Loop")]);
    }

    /// <summary>
    /// D79, through two names that refer to each other.
    /// </summary>
    [Test]
    public async Task D79_names_that_refer_to_each_other_are_each_followed_once()
    {
        var dependencies = GetDependencies("Ping", init: wb =>
        {
            wb.DefinedNames.Add("Ping", "Pong+Sheet!$B$2");
            wb.DefinedNames.Add("Pong", "Ping+Sheet!$C$3");
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("B2")),
            new("Sheet", Area.Parse("C3")),
        });
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("Ping"), new XLName("Pong")]);
    }

    /// <summary>
    /// D79. The guard is the path from the cell formula, not every name seen so far: a name reached
    /// a second time from another branch is followed again. Were it skipped, the range operator
    /// would see only <c>D4</c> and lose <c>B2:D4</c>. The reference comes first in the range,
    /// because <c>Corner:Sheet!$D$4</c> reads as a 3D reference from sheet <c>Corner</c>.
    /// </summary>
    [Test]
    public async Task D79_a_name_reached_from_two_branches_is_followed_from_each()
    {
        var dependencies = GetDependencies("SUM(Corner)+SUM(Sheet!$D$4:Corner)", init: wb =>
        {
            wb.DefinedNames.Add("Corner", "Sheet!$B$2");
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("B2")),
            new("Sheet", Area.Parse("B2:D4")),
        });
    }

    /// <summary>
    /// D79 in a workbook. A write after a read builds the dependency tree, which followed the name
    /// for ever. The write now completes. The cell that uses the name is never evaluated here, so
    /// it is left dirty: evaluating a circular name is a separate fix.
    /// </summary>
    [Test]
    [Arguments("Loop")]
    [Arguments("Ping")]
    public async Task D79_a_write_completes_when_a_cell_uses_a_circular_name(string name)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        if (name == "Loop")
        {
            wb.DefinedNames.Add("Loop", "Loop+1");
        }
        else
        {
            wb.DefinedNames.Add("Ping", "Pong");
            wb.DefinedNames.Add("Pong", "Ping");
        }

        ws.Cell("A1").FormulaA1 = name;
        ws.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);

        ws.Cell("C1").Value = 5;
        ws.Cell("C2").Value = 6;

        await Assert.That(ws.Cell("C2").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
    }

    /// <summary>
    /// D79. A name used twice in one formula still gives the formula its precedents, so a change to
    /// the name's source marks the formula dirty.
    /// </summary>
    [Test]
    public async Task D79_a_name_used_twice_still_marks_its_dependent_dirty()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        wb.DefinedNames.Add("Total", "Sheet!$B$2");
        ws.Cell("B2").Value = 3;
        ws.Cell("A1").FormulaA1 = "Total+Total";
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(6);

        ws.Cell("B2").Value = 4;

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(8);
    }

    [Test]
    public async Task Name_range_that_is_not_a_reference_can_be_added_to_dependency_tree_without_exception()
    {
        var dependencies = GetDependencies("name", init: wb =>
        {
            wb.DefinedNames.Add("name", "1+3");
        });
        await Assert.That(dependencies.Areas).IsEmpty();
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("name")]);
    }

    [Test]
    public async Task Name_range_can_be_sheet_scoped_even_without_specified_sheet()
    {
        // Formula that references a name that is ambiguous between workbook and worksheet scoped one.
        const string formula = "name";
        var dependencies = GetDependencies(formula, init: wb =>
        {
            // Define two names, the local one should be selected
            wb.Worksheet("Sheet").DefinedNames.Add("name", "Sheet!$A$1");
            wb.DefinedNames.Add("name", "Sheet!$B$10");
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("A1"))
        });
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("name")]);
    }

    [Test]
    [Skip("A1 to R1C1 conversion not yet implemented and the name formula must be parsed")]
    public async Task Name_range_that_uses_relative_reference_determines_actual_precedent_areas_through_cell_location()
    {
        var dependencies = GetDependencies("name", "D8", init: wb =>
        {
            wb.DefinedNames.Add("name", "Sheet!B4"); // equivalent of R[3]C[2]
        });
        await Assert.That(dependencies.Areas).IsEquivalentTo(new SheetArea[]
        {
            new("Sheet", Area.Parse("F7")), // D4 (formula cell) + R[3]C[2] (name relative reference) = F7
        });
        await Assert.That(dependencies.Names).IsEquivalentTo([new XLName("name")]);
    }

    #endregion

    #region Remove formula from dependency tree

    [Test]
    public async Task Remove_formula_from_dependency_tree()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var tree = new DependencyTree();
        tree.AddSheetTree(ws);
        var cellFormula = AddFormula(tree, ws, "B3", "=C4");
        await Assert.That(tree.IsEmpty).IsFalse();

        // Remove inserted formula removes the dependent and also removes the precedent
        // area from the tree because there is no formula depending on it.
        tree.RemoveFormula(cellFormula);
        await Assert.That(tree.IsEmpty).IsTrue();

        // Removing already removed formula doesn't throw.
        await Assert.That(() => tree.RemoveFormula(cellFormula)).ThrowsNothing();
        await Assert.That(tree.IsEmpty).IsTrue();
    }

    [Test]
    public async Task Removing_formula_doesnt_remove_precedent_area_from_tree_when_another_formula_depends_on_the_area()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var tree = new DependencyTree();
        tree.AddSheetTree(ws);
        var cellFormulaA1 = AddFormula(tree, ws, "A1", "=C4 + B1");
        var cellFormulaA2 = AddFormula(tree, ws, "A2", "=B1 / C4");
        await Assert.That(tree.IsEmpty).IsFalse();

        // Remove first formula, but the precedent area is still used
        // by second formula so it is not removed.
        tree.RemoveFormula(cellFormulaA1);
        await Assert.That(tree.IsEmpty).IsFalse();

        // Remove second formula
        tree.RemoveFormula(cellFormulaA2);
        await Assert.That(tree.IsEmpty).IsTrue();
    }

    /// <summary>
    /// A removed formula takes its precedent cells and areas out of the tree, and a precedent cell
    /// that another formula reads stays. A precedent cell is kept outside the R-tree (#513).
    /// </summary>
    [Test]
    public async Task Removing_formula_removes_its_precedent_cells_and_areas_from_tree()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var tree = new DependencyTree();
        tree.AddSheetTree(ws);
        var first = AddFormula(tree, ws, "B1", "=A1+SUM(A2:A3)");
        var second = AddFormula(tree, ws, "B2", "=A1");

        tree.RemoveFormula(first);
        MarkDirty(tree, ws, "A2:A3");
        await AssertNotDirty(ws, "B1:B2");

        MarkDirty(tree, ws, "A1");
        await AssertDirty(ws, "B2");
        await AssertNotDirty(ws, "B1");

        tree.RemoveFormula(second);
        await Assert.That(tree.IsEmpty).IsTrue();
    }

    #endregion

    #region Mark dirty

    [Test]
    public async Task Mark_dirty_single_chain_is_fully_marked()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "A2", "=A1");
        AddFormula(tree, ws, "A3", "=A2");
        AddFormula(tree, ws, "A4", "=A3");

        MarkDirty(tree, ws, "A1");
        await AssertDirty(ws, "A2", "A3", "A4");
    }

    [Test]
    public async Task Mark_dirty_split_and_join_is_fully_marked()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "B2", "=B1");
        AddFormula(tree, ws, "C1", "=B2");
        AddFormula(tree, ws, "C3", "=B2");
        AddFormula(tree, ws, "D2", "=C1 + C3");

        MarkDirty(tree, ws, "B1");
        await AssertDirty(ws, "B2", "C1", "C3", "D2");
    }

    [Test]
    public async Task Mark_dirty_uses_correct_sheet()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws1 = wb.AddWorksheet("Sheet1");
        tree.AddSheetTree(ws1);
        var ws2 = wb.AddWorksheet("Sheet2");
        tree.AddSheetTree(ws2);

        // Make a chain, where each cell is on an opposite sheet
        AddFormula(tree, ws1, "B1", "=Sheet2!A1");
        AddFormula(tree, ws2, "C1", "=Sheet1!B1");
        AddFormula(tree, ws1, "D1", "=Sheet2!C1");
        AddFormula(tree, ws2, "E1", "=Sheet1!D1");

        // Formulas on opposite sheet
        AddFormula(tree, ws2, "B1", "=Sheet1!A1");
        AddFormula(tree, ws1, "C1", "=Sheet2!B1");
        AddFormula(tree, ws2, "D1", "=Sheet1!C1");
        AddFormula(tree, ws1, "E1", "=Sheet2!D1");

        MarkDirty(tree, ws2, "A1");
        await AssertDirty(ws1, "B1", "D1");
        await AssertDirty(ws2, "C1", "E1");

        await AssertNotDirty(ws1, "C1", "E1");
        await AssertNotDirty(ws2, "B1", "D1");
    }

    /// <summary>
    /// Spec 40: a formula already dirty for a reason unrelated to this walk (the same primitive
    /// <c>InvalidateFormula</c>, a sheet rename or a reference shift use) must still be traversed,
    /// so anything downstream of it is reached. Before that fix, the walk used "already dirty" as
    /// its own "already visited" marker and stopped at A3, leaving A4 stale.
    /// </summary>
    [Test]
    public async Task Mark_dirty_continues_through_a_cell_already_dirty_for_an_unrelated_reason()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "A2", "=A1");
        AddFormula(tree, ws, "A3", "=A2");
        AddFormula(tree, ws, "A4", "=A3");

        // Simulate an external dirtier (e.g. InvalidateFormula) marking the middle cell dirty
        // before this walk starts.
        ((XLCell)ws.Cell("A3")).Formula!.MarkExplicitlyDirty();

        MarkDirty(tree, ws, "A1");
        await AssertDirty(ws, "A2", "A3", "A4");
    }

    [Test]
    public async Task Mark_dirty_wont_crash_on_cycle()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "B1", "=D1 + A1");
        AddFormula(tree, ws, "C1", "=B1");
        AddFormula(tree, ws, "D1", "=C1");

        // Tail depending on the cycle
        AddFormula(tree, ws, "E1", "=D1");

        MarkDirty(tree, ws, "A1");
        await AssertDirty(ws, "B1", "C1", "D1", "E1");
    }

    [Test]
    public async Task Mark_dirty_affects_precedents_with_partial_overlap()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "D1", "=A1:B3");

        // B3:D4 overlaps with A1:B3 in B3
        MarkDirty(tree, ws, "B3:D4");
        await AssertDirty(ws, "D1");
    }

    [Test]
    public async Task Mark_dirty_can_affect_multiple_chains_at_once()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "B1", "=A1");
        AddFormula(tree, ws, "B2", "=A2");
        AddFormula(tree, ws, "B3", "=A3");

        MarkDirty(tree, ws, "A2:A3");
        await AssertDirty(ws, "B2", "B3");
        await AssertNotDirty(ws, "B1");
    }

    /// <summary>
    /// <see cref="DependencyTree.CreateFrom"/> loads the precedent areas of each sheet into the
    /// R-tree in one bulk load, but <see cref="DependencyTree.AddFormula"/> inserts areas one at a
    /// time (#513). Each formula in column C reads an area of two cells, so the bulk load has enough
    /// areas to build more than one R-tree node. A single cell is kept outside the R-tree, and every
    /// formula in column B reads <c>$D$1</c>, so that one cell has 50 dependents.
    /// </summary>
    [Test]
    public async Task Tree_built_from_a_workbook_marks_dependents_and_accepts_later_changes()
    {
        const int rows = 50;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= rows; row++)
        {
            SetCleanFormula(ws, $"B{row}", $"A{row}*$D$1");
            SetCleanFormula(ws, $"C{row}", $"B{row}+SUM(F{row}:G{row})");
        }

        var tree = DependencyTree.CreateFrom(wb);

        MarkDirty(tree, ws, "A7");
        await AssertDirty(ws, "B7:C7");
        await AssertNotDirty(ws, "B1:C6", $"B8:C{rows}");

        MarkAllClean(ws, rows);
        MarkDirty(tree, ws, "G9");
        await AssertDirty(ws, "C9");
        await AssertNotDirty(ws, $"B1:B{rows}", "C1:C8", $"C10:C{rows}");

        MarkAllClean(ws, rows);
        MarkDirty(tree, ws, "D1");
        await AssertDirty(ws, $"B1:C{rows}");

        // After the bulk load, the tree still takes one formula out and puts another in.
        MarkAllClean(ws, rows);
        tree.RemoveFormula(((XLCell)ws.Cell("B7")).Formula!);
        var replacement = SetCleanFormula(ws, "B7", "A8");
        tree.AddFormula(new SheetArea(ws.Name, Area.Parse("B7")), replacement, wb);

        MarkDirty(tree, ws, "A7");
        await AssertNotDirty(ws, "B7:C7");

        MarkDirty(tree, ws, "A8");
        await AssertDirty(ws, "B7:C8");
    }

    /// <summary>
    /// A precedent cell is kept outside the R-tree (#513). A dirty area looks up each of its cells
    /// when it has no more cells than the sheet has precedent cells, and tests each precedent cell
    /// when it has more. The sheet has three precedent cells, A1, A2 and A3, and one precedent area,
    /// A1:A2.
    /// </summary>
    [Test]
    [Arguments("A2", "B2,C1", "B1,B3")]
    [Arguments("A1:A3", "B1:B3,C1", "")]
    [Arguments("A3:A4", "B3", "B1:B2,C1")]
    [Arguments("A3:A1048576", "B3", "B1:B2,C1")]
    [Arguments("A1:A1048576", "B1:B3,C1", "")]
    [Arguments("B1:D1048576", "", "B1:B3,C1")]
    public async Task Dirty_area_marks_the_dependents_of_the_precedent_cells_in_it(string dirtyArea, string dirty,
        string clean)
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "B1", "=A1");
        AddFormula(tree, ws, "B2", "=A2");
        AddFormula(tree, ws, "B3", "=A3");
        AddFormula(tree, ws, "C1", "=SUM(A1:A2)");

        MarkDirty(tree, ws, dirtyArea);
        await AssertDirty(ws, Split(dirty));
        await AssertNotDirty(ws, Split(clean));

        static string[] Split(string ranges) => ranges.Split(',', StringSplitOptions.RemoveEmptyEntries);
    }

    private static XLCellFormula SetCleanFormula(IXLWorksheet sheet, string address, string formula)
    {
        var cell = (XLCell)sheet.Cell(address);
        cell.Formula = XLCellFormula.NormalA1(formula);
        cell.Formula.MarkClean();
        return cell.Formula;
    }

    private static void MarkAllClean(IXLWorksheet sheet, int rows)
    {
        for (var row = 1; row <= rows; row++)
        {
            ((XLCell)sheet.Cell(row, 2)).Formula!.MarkClean();
            ((XLCell)sheet.Cell(row, 3)).Formula!.MarkClean();
        }
    }

    #region Shared formulas

    /// <summary>
    /// #513. A tree built from a loaded workbook parses each shared formula once, from the R1C1 text
    /// that the loader kept, and resolves it for each cell. Each cell must get the precedents that its
    /// own A1 text gives. The groups have relative, mixed, absolute, other-sheet and defined-name
    /// references, and one group is filled across a row.
    /// </summary>
    [Test]
    public async Task Shared_formula_cells_get_the_precedents_of_their_own_A1_text()
    {
        using var wb = LoadWithSharedFormulas();
        var sheet = wb.Worksheet("Sheet1");
        var tree = DependencyTree.CreateFrom(wb);

        foreach (var address in SharedFormulaCells)
        {
            var cell = (XLCell)sheet.Cell(address);
            var formula = cell.Formula!;
            var formulaArea = new SheetArea(sheet.Name, new Area(cell.SheetPoint, cell.SheetPoint));

            // Without the kept text, the build parses the A1 text, and the comparison proves nothing.
            await Assert.That(formula.TryGetSharedR1C1(cell.SheetPoint, out _)).IsTrue();
            await Assert.That(tree.GetKeptPrecedents(formula))
                .IsEquivalentTo(tree.GetPrecedents(formulaArea, formula, wb).Areas);
        }
    }

    /// <summary>
    /// #513. A row inserted above moves a formula that reads another sheet, and its A1 text stays the
    /// same. Its R1C1 text changes, so the build must not resolve the kept R1C1 text at the new cell.
    /// </summary>
    [Test]
    public async Task Shared_formula_moved_by_a_row_insert_depends_on_what_its_text_reads()
    {
        using var wb = LoadWithSharedFormulas();
        var sheet = wb.Worksheet("Sheet1");
        var other = wb.Worksheet("Sheet2");

        sheet.Row(1).InsertRowsAbove(1);
        var moved = (XLCell)sheet.Cell("F2");
        await Assert.That(moved.FormulaA1).IsEqualTo("Sheet2!A1*2");
        await Assert.That(moved.Formula!.TryGetSharedR1C1(moved.SheetPoint, out _)).IsFalse();
        foreach (var cell in sheet.Range("F2:F5").Cells())
            _ = cell.Value;

        other.Cell("A1").Value = 7;

        await Assert.That(sheet.Cell("F2").NeedsRecalculation).IsTrue();
        await Assert.That(sheet.Cell("F3").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("F2").Value).IsEqualTo(14);
    }

    [Test]
    public async Task Kept_R1C1_text_is_given_only_at_its_cell_and_until_the_A1_text_changes()
    {
        var formula = XLCellFormula.NormalA1("A1*2");
        formula.SetSharedR1C1("RC[-1]*2", new Point(1, 2));

        await Assert.That(formula.TryGetSharedR1C1(new Point(1, 2), out var r1c1)).IsTrue();
        await Assert.That(r1c1).IsEqualTo("RC[-1]*2");
        await Assert.That(formula.TryGetSharedR1C1(new Point(2, 2), out _)).IsFalse();

        formula.UpdateShiftedA1("A2*2");
        await Assert.That(formula.TryGetSharedR1C1(new Point(1, 2), out _)).IsFalse();
    }

    private static readonly string[] SharedFormulaCells =
    [
        "C1", "C2", "C3", "C4", "D1", "D2", "D3", "D4", "E1", "E2", "E3", "E4", "F1", "F2", "F3", "F4",
        "G1", "H1", "I1",
    ];

    /// <summary>
    /// Save a workbook with XLibur, which writes a formula in each cell. Then rewrite the formulas as
    /// shared formulas, as Excel saves them, and load the result.
    /// </summary>
    private static XLWorkbook LoadWithSharedFormulas()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var other = wb.AddWorksheet("Sheet2");
            wb.DefinedNames.Add("Total", "Sheet1!$B$4");
            for (var row = 1; row <= 4; row++)
            {
                ws.Cell(row, 1).Value = row;
                ws.Cell(row, 2).Value = row * 10;
                other.Cell(row, 1).Value = row * 100;
                ws.Cell(row, 3).FormulaA1 = $"A{row}+$B$1+Sheet2!A{row}";
                ws.Cell(row, 4).FormulaA1 = $"SUM(A$1:A{row})";
                ws.Cell(row, 5).FormulaA1 = $"IF(A{row}=1,B{row},Total)";
                ws.Cell(row, 6).FormulaA1 = $"Sheet2!A{row}*2";
            }

            ws.Cell("G1").FormulaA1 = "A1*2";
            ws.Cell("H1").FormulaA1 = "B1*2";
            ws.Cell("I1").FormulaA1 = "C1*2";
            wb.RecalculateAllFormulas();
            wb.SaveAs(package);
        }

        package.RewriteSheet1(xml =>
        {
            xml = ShareFormula(xml, 0, "C1:C4", Rows(row => $"A{row}+$B$1+Sheet2!A{row}"));
            xml = ShareFormula(xml, 1, "D1:D4", Rows(row => $"SUM(A$1:A{row})"));
            xml = ShareFormula(xml, 2, "E1:E4", Rows(row => $"IF(A{row}=1,B{row},Total)"));
            xml = ShareFormula(xml, 3, "F1:F4", Rows(row => $"Sheet2!A{row}*2"));
            return ShareFormula(xml, 4, "G1:I1", ["A1*2", "B1*2", "C1*2"]);

            static string[] Rows(Func<int, string> formula) => [formula(1), formula(2), formula(3), formula(4)];
        });

        package.Position = 0;
        return new XLWorkbook(package);
    }

    /// <summary>
    /// Make the formulas in <paramref name="cellFormulas"/>, in cell order, one shared formula. The
    /// first cell keeps its text and the others refer to it by <paramref name="index"/>.
    /// </summary>
    private static string ShareFormula(string sheetXml, int index, string reference, string[] cellFormulas)
    {
        for (var i = 0; i < cellFormulas.Length; ++i)
        {
            var original = $"<x:f>{cellFormulas[i]}</x:f>";
            var shared = i == 0
                ? $"<x:f t=\"shared\" ref=\"{reference}\" si=\"{index}\">{cellFormulas[i]}</x:f>"
                : $"<x:f t=\"shared\" si=\"{index}\" />";
            var rewritten = sheetXml.Replace(original, shared, StringComparison.Ordinal);
            if (ReferenceEquals(rewritten, sheetXml))
                throw new InvalidOperationException($"'{original}' was not found in the sheet part.");

            sheetXml = rewritten;
        }

        return sheetXml;
    }

    #endregion Shared formulas

    #region Precedents factory

    /// <summary>
    /// #513, change 6. <see cref="PrecedentsFactory"/> collects the precedents while the parser reads a
    /// formula, and <see cref="DependenciesVisitor"/> reads the AST of a shared formula. Both must give
    /// the same precedents, names and unknown flag. Neither throws: text that either cannot read, a
    /// function with the wrong number of arguments included (#543), is a refusal for both. The factory
    /// may ask for the AST only where it cannot read the formula in one pass.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(ParityFormulas))]
    public async Task Precedents_factory_gives_what_the_visitor_gives(string formula)
    {
        using var wb = ParityWorkbook();
        var formulaArea = new SheetArea("Sheet", Area.Parse("D2"));

        var visited = CollectThroughAst(wb, formulaArea, formula);
        var walked = new FormulaDependencies();
        var context = new DependenciesContext(formulaArea, wb, walked);
        var accepted = new PrecedentsFactory().TryCollect(formula, context);

        // The tree clears what a refused walk added, so the formula has no precedents but unknown ones.
        var viaTree = new DependencyTree().GetPrecedents(formulaArea, XLCellFormula.NormalA1(formula), wb);
        await Assert.That(viaTree.Areas).IsEquivalentTo(visited.Areas);
        await Assert.That(viaTree.Names).IsEquivalentTo(visited.Names);
        await Assert.That(viaTree.HasUnknownPrecedents).IsEqualTo(visited.HasUnknownPrecedents);

        if (!accepted)
        {
            await Assert.That(visited.HasUnknownPrecedents).IsTrue();
            await Assert.That(visited.Areas).IsEmpty();
            return;
        }

        if (context.NeedsAst)
        {
            await Assert.That(FormulasThatNeedTheAst).Contains(formula);
            return;
        }

        await Assert.That(walked.Areas).IsEquivalentTo(visited.Areas);
        await Assert.That(walked.Names).IsEquivalentTo(visited.Names);
        await Assert.That(walked.HasUnknownPrecedents).IsEqualTo(visited.HasUnknownPrecedents);
    }

    /// <summary>
    /// #543. A function with the wrong number of arguments is a refusal for the factory and for the
    /// visitor, as <c>1+</c> is. Both threw <see cref="ExpressionParseException"/>, and the tree let it
    /// out of every build. The tree now takes the formula to depend on every cell, and drops the
    /// precedents the walk added before it refused the text.
    /// </summary>
    [Test]
    [Arguments("1+")]
    [Arguments("ABS()")]
    [Arguments("ABS(1,2)")]
    [Arguments("ABS(A1,B1)")]
    public async Task A_wrong_number_of_arguments_is_a_refusal_as_unreadable_text_is(string formula)
    {
        using var wb = ParityWorkbook();
        var formulaArea = new SheetArea("Sheet", Area.Parse("D2"));
        var context = new DependenciesContext(formulaArea, wb, new FormulaDependencies());

        await Assert.That(new PrecedentsFactory().TryCollect(formula, context)).IsFalse();
        await Assert.That(wb.CalcEngine.TryParse(formula, out _)).IsFalse();

        var viaTree = new DependencyTree().GetPrecedents(formulaArea, XLCellFormula.NormalA1(formula), wb);
        await Assert.That(viaTree.HasUnknownPrecedents).IsTrue();
        await Assert.That(viaTree.Areas).IsEmpty();
    }

    /// <summary>
    /// A function from another cell that is not a function, and a defined name that the parser refuses.
    /// The visitor does not read their arguments, and the factory has read them already.
    /// </summary>
    private static readonly string[] FormulasThatNeedTheAst = ["A1(B1+C1)", "Bad", "Bad+1"];

    public static IEnumerable<object[]> ParityFormulas()
    {
        string[] formulas =
        [
            "A1", "$A$1:B2", "A:A", "1:1", "Other!A1", "'Other'!B2:C3", "Sheet3!A1",
            "A1+B1*C1", "-A1", "A1%", "@A1:A3", "A1#", "(A1,B2)", "A1:C3 B2:D4", "A1:C3 E5:F6",
            "A1:Other!B2", "B3:name", "name+D2", "Loop", "Other!local", "local", "[0]!name",
            "SUM(A1:A3,B1)", "IF(A1,B1,C1)", "IF(A1,B1)", "IF(A1,B1:B3,5)", "IF(A1,B1,C1):D5",
            "CHOOSE(A1,B1,C1:C2,3)", "INDEX(A1:C3,B1,C1)", "INDEX(A1:C3,B1):D4", "INDIRECT(A1)",
            "OFFSET(A1,1,1)", "_xlfn.XLOOKUP(A1,B1:B5,C1:C5)", "FOO(A1)", "LOG10(A1)", "A1(B1+C1)",
            "{1,2;3,4}", "#N/A", "Other!#REF!", "\"text\"", "TRUE",
            "Sheet:Other!A1", "[1]Sheet1!A1", "TableName[Second]", "SUM(TableName[[#All],[Second]])",
            "1+", "'[Book2.xlsx]Sheet1'!A1", "Bad", "Bad+1", "ABS()", "ABS(1,2)",
        ];

        foreach (var formula in formulas)
            yield return [formula];
    }

    private static XLWorkbook ParityWorkbook()
    {
        var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet");
        var other = wb.AddWorksheet("Other");
        AddTable(wb);
        wb.DefinedNames.Add("name", "Sheet!$B$4");
        wb.DefinedNames.Add("Loop", "Loop+1");
        // Kept as a load keeps it: the public Add refuses text that the parser refuses.
        wb.DefinedNamesInternal.Add("Bad", "'[Book2.xlsx]Sheet1'!A1", null, validateName: true,
            validateRangeAddress: false);
        other.DefinedNames.Add("local", "Other!$C$3");
        return wb;
    }

    /// <summary>
    /// The precedents as the tree found them before #513 change 6: parse the text into an AST, and visit it.
    /// </summary>
    private static FormulaDependencies CollectThroughAst(XLWorkbook wb, SheetArea formulaArea, string formula)
    {
        var dependencies = new FormulaDependencies();
        if (!wb.CalcEngine.TryParse(formula, out var ast))
        {
            dependencies.MarkPrecedentsUnknown();
            return dependencies;
        }

        var context = new DependenciesContext(formulaArea, wb, dependencies);
        var root = ast.AstRoot.Accept(context, new DependenciesVisitor());
        if (root.IsReference)
            context.AddAreas(root);

        return dependencies;
    }

    #endregion Precedents factory

    #endregion

    #region Rename sheet

    [Test]
    public async Task Sheet_rename_keeps_tree_same_only_with_changed_sheet_name()
    {
        using var wb = new XLWorkbook();
        var renamedSheet = wb.AddWorksheet("Original");
        var unchangedSheet = wb.AddWorksheet("Unchanged");

        renamedSheet.Cell("A1").Value = 1;
        renamedSheet.Cell("A2").Value = 2;
        renamedSheet.Cell("A3").Value = 3;
        renamedSheet.Cell("A4").FormulaA1 = "SUM(Original!A1:A2, A3, Unchanged!A1:A2)";
        unchangedSheet.Cell("A1").Value = 10;
        unchangedSheet.Cell("A2").Value = 20;
        unchangedSheet.Cell("A3").Value = 30;
        unchangedSheet.Cell("A4").FormulaA1 = "SUM(Unchanged!A1:A2, A3, Original!A1:A2)";
        Recalculate();

        renamedSheet.Name = "Renamed";

        await Assert.That(renamedSheet.Cell("A4").FormulaA1).IsEqualTo("SUM(Renamed!A1:A2, A3, Unchanged!A1:A2)");
        await Assert.That(unchangedSheet.Cell("A4").FormulaA1).IsEqualTo("SUM(Unchanged!A1:A2, A3, Renamed!A1:A2)");

        Recalculate();
        await Assert.That(renamedSheet.Cell("A4").NeedsRecalculation).IsFalse();
        await Assert.That(unchangedSheet.Cell("A4").NeedsRecalculation).IsFalse();

        // Both depend on Unchanged!A1
        unchangedSheet.Cell("A1").Value = 110;
        await Assert.That(renamedSheet.Cell("A4").NeedsRecalculation).IsTrue();
        await Assert.That(unchangedSheet.Cell("A4").NeedsRecalculation).IsTrue();
        Recalculate();
        await Assert.That(renamedSheet.Cell("A4").CachedValue).IsEqualTo(136);
        await Assert.That(unchangedSheet.Cell("A4").CachedValue).IsEqualTo(163);

        // Both depend on Renamed!A1
        renamedSheet.Cell("A1").Value = 201;
        await Assert.That(renamedSheet.Cell("A4").NeedsRecalculation).IsTrue();
        await Assert.That(unchangedSheet.Cell("A4").NeedsRecalculation).IsTrue();
        Recalculate();
        await Assert.That(renamedSheet.Cell("A4").CachedValue).IsEqualTo(336);
        await Assert.That(unchangedSheet.Cell("A4").CachedValue).IsEqualTo(363);

        // Only unchanged depends on Unchanged!A3. The renamed formula keeps value.
        unchangedSheet.Cell("A3").Value = 330;
        await Assert.That(renamedSheet.Cell("A4").NeedsRecalculation).IsFalse();
        await Assert.That(unchangedSheet.Cell("A4").NeedsRecalculation).IsTrue();
        Recalculate();
        await Assert.That(renamedSheet.Cell("A4").CachedValue).IsEqualTo(336);
        await Assert.That(unchangedSheet.Cell("A4").CachedValue).IsEqualTo(663);

        // Only renamed depends on Renamed!A3. The unchanged formula keeps value.
        renamedSheet.Cell("A3").Value = 403;
        await Assert.That(renamedSheet.Cell("A4").NeedsRecalculation).IsTrue();
        await Assert.That(unchangedSheet.Cell("A4").NeedsRecalculation).IsFalse();
        Recalculate();
        await Assert.That(renamedSheet.Cell("A4").CachedValue).IsEqualTo(736);
        await Assert.That(unchangedSheet.Cell("A4").CachedValue).IsEqualTo(663);

        void Recalculate()
        {
            // Force recalculation to clear dirty flag. Recalculation always happens for whole
            // calculation chain.
            wb.CalcEngine.Recalculate(wb, null);
        }
    }

    #endregion

    private static XLCellFormula AddFormula(DependencyTree tree, IXLWorksheet sheet, string address, string formula)
    {
        // Set directly, so the cell is not marked as a dirty.
        var cell = (XLCell)sheet.Cell(address);
        cell.Formula = XLCellFormula.NormalA1(formula);
        // Pre-mark the formula as clean so that the dependency-tree MarkDirty walk has a
        // meaningful "started clean -> became dirty" transition to assert against. A
        // freshly-constructed XLCellFormula defaults to dirty (never evaluated).
        cell.Formula.MarkClean();
        var cellArea = new SheetArea(sheet.Name, new Area(cell.SheetPoint, cell.SheetPoint));
        tree.AddFormula(cellArea, cell.Formula, sheet.Workbook);
        return cell.Formula;
    }

    private static void MarkDirty(DependencyTree tree, IXLWorksheet sheet, string range)
    {
        var area = new SheetArea(sheet.Name, Area.Parse(range));
        tree.MarkDirty(area);
    }

    private static async Task AssertDirty(IXLWorksheet sheet, params string[] dirtyRanges)
    {
        await AssertDirtyFlag(true, sheet, dirtyRanges);
    }
    private static async Task AssertNotDirty(IXLWorksheet sheet, params string[] dirtyRanges)
    {
        await AssertDirtyFlag(false, sheet, dirtyRanges);
    }

    private static async Task AssertDirtyFlag(bool expectedDirtyFlag, IXLWorksheet sheet, params string[] dirtyRanges)
    {
        var ws = (XLWorksheet)sheet;
        foreach (var dirtyRange in dirtyRanges)
        {
            foreach (var dirtyCell in ws.Cells(dirtyRange))
            {
                await Assert.That(dirtyCell.Formula?.IsDirty()).IsEqualTo(expectedDirtyFlag);
            }
        }
    }

    #region Structured references

    /// <summary>
    /// A structured reference resolves to the table's current area, so a formula using one must
    /// register that area as a precedent. Until it did, nothing invalidated the formula when the
    /// table's cells changed.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(StructuredReferenceDependencyTestCases))]
    public async Task Structured_reference_is_a_dependency_of_the_area_it_covers(
        string formula,
        string expectedArea)
    {
        var dependencies = GetDependencies(formula, "A1", AddTable);

        await Assert.That(dependencies.Areas)
            .IsEquivalentTo(new SheetArea[] { new("Sheet", Area.Parse(expectedArea)) });
    }

    public static IEnumerable<object[]> StructuredReferenceDependencyTestCases
    {
        get
        {
            // Table occupies E7:H10 — headers on row 7, data rows 8..10.
            yield return ["SUM(TableName[Second])", "F8:F10"];
            yield return ["SUM(TableName[])", "E8:H10"];
            yield return ["SUM(TableName[#All])", "E7:H10"];
            yield return ["SUM(TableName[#Headers])", "E7:H7"];
            yield return ["SUM(TableName[[Second]:[Fourth]])", "F8:H10"];
        }
    }

    /// <summary>
    /// The reference is propagated to the parent node rather than added directly, so an
    /// enclosing range operator can combine it — the contract the other reference nodes follow.
    /// </summary>
    [Test]
    public async Task Structured_reference_is_propagated_to_an_enclosing_range_operator()
    {
        var dependencies = GetDependencies("B3:TableName[Second]", "A1", AddTable);

        await Assert.That(dependencies.Areas)
            .IsEquivalentTo(new SheetArea[] { new("Sheet", Area.Parse("B3:F10")) });
    }

    /// <summary>
    /// An unresolvable reference contributes no precedent rather than failing — the formula is
    /// a <c>#REF!</c>, and whatever later makes the table resolve rebuilds the tree.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(UnresolvableStructuredReferenceTestCases))]
    public async Task Unresolvable_structured_reference_has_no_dependencies(string formula)
    {
        var dependencies = GetDependencies(formula, "A1", AddTable);

        await Assert.That(dependencies.Areas).IsEmpty();
    }

    public static IEnumerable<object[]> UnresolvableStructuredReferenceTestCases
    {
        get
        {
            yield return ["SUM(WrongName[Second])"];
            yield return ["SUM(TableName[NonExistentCol])"];
        }
    }

    /// <summary>
    /// The end-to-end consequence: editing a cell inside the table has to invalidate a formula
    /// that reads the table through a structured reference.
    /// </summary>
    [Test]
    public async Task Editing_a_table_cell_dirties_a_formula_using_a_structured_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        AddTable(wb);

        ws.Cell("F8").Value = 1;
        ws.Cell("F9").Value = 2;
        ws.Cell("F10").Value = 3;
        ws.Cell("A1").FormulaA1 = "SUM(TableName[Second])";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(6);

        ws.Cell("F9").Value = 20;

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(24);
    }

    /// <summary>
    /// A table name is workbook scoped, so the area it resolves to lies on the table's own sheet,
    /// which need not be the sheet holding the formula. Registering the precedent against the
    /// formula's sheet would watch entirely unrelated cells.
    /// </summary>
    [Test]
    public async Task Structured_reference_depends_on_the_sheet_owning_the_table()
    {
        using var wb = new XLWorkbook();
        var report = wb.AddWorksheet("Report");
        wb.AddWorksheet("Data");
        AddTable(wb, "Data");

        var tree = new DependencyTree();
        var cell = report.Cell("A1");
        cell.SetFormulaA1("SUM(TableName[Second])");

        var cellFormula = ((XLCell)cell).Formula!;
        var dependencies = tree.GetPrecedents(new SheetArea(report.Name, cellFormula.Range), cellFormula, wb);

        await Assert.That(dependencies.Areas)
            .IsEquivalentTo(new SheetArea[] { new("Data", Area.Parse("F8:F10")) });
    }

    /// <summary>
    /// A 4x3 table at E7 with headers First..Fourth, matching the layout
    /// <c>StructuredReferenceTests</c> uses.
    /// </summary>
    private static void AddTable(XLWorkbook wb) => AddTable(wb, "Sheet");

    private static void AddTable(XLWorkbook wb, string sheetName)
    {
        var ws = wb.Worksheet(sheetName);
        ws.Cell("E7").Value = "First";
        ws.Cell("F7").Value = "Second";
        ws.Cell("G7").Value = "Third";
        ws.Cell("H7").Value = "Fourth";
        ws.Range("E7:H10").CreateTable("TableName");
    }

    #endregion Structured references

    private static FormulaDependencies GetDependencies(string formula, string formulaAddress = "A1", Action<XLWorkbook> init = null!)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        init?.Invoke(wb);
        var tree = new DependencyTree();
        var cell = ws.Cell(formulaAddress);
        cell.SetFormulaA1(formula);

        var cellFormula = ((XLCell)cell).Formula!;
        var dependencies = tree.GetPrecedents(new SheetArea(ws.Name, cellFormula.Range), cellFormula, wb);
        return dependencies;
    }

    public static IEnumerable<object[]> AreaDependenciesTestCases
    {
        get
        {
            // When a visitor visits a node, there are two choices for found references:
            // * propagate the reference to parent node (in most cases checked by range operator)
            // * add the reference directly to the dependencies

            // A formula that is a simple reference is propagated to the root
            yield return
            [
                "A1",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1"))
                }
            ];

            // References are in a multiple levels of an expression without ref expression or
            // a function are added
            yield return
            [
                "7+A1/(B1+C1)",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1")),
                    new SheetArea("Sheet", Area.Parse("B1")),
                    new SheetArea("Sheet", Area.Parse("C1"))
                }
            ];

            // Unary implicit intersection is propagated
            yield return
            [
                "@A1:A4",
                new[]
                {
                    // Implicit intersection
                    new SheetArea("Sheet", Area.Parse("A1:A4")),
                }
            ];

            // Implicit intersection binds looser than the range operator and takes the rest of it,
            // so this is D3:(@(A1:C2)). The dependency is the whole operand of @, so the range
            // spans all three references.
            yield return
            [
                "D3:@A1:C2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1:D3")),
                }
            ];

            // Unary spill operator propagates a reference
            yield return
            [
                "F2#:A7",
                new[]
                {
                    // This is not correct, but until spill operator works,
                    // but for now it provides best approximate for now.
                    new SheetArea("Sheet", Area.Parse("A2:F7")),
                }
            ];

            // Unary value operators (in this case percent) applied on reference adds it
            yield return
            [
                "4+A4%",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A4")),
                }
            ];

            // Union operation propagates references
            yield return
            [
                "(A1:B2,C1:D2):E3",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1:E3"))
                }
            ];

            // Range operation propagates
            yield return
            [
                // Due to greedy nature, the A1:C4 is the first reference and D2 is the second
                "A1:C4:D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1:D4")),
                }
            ];

            // Range operation with multiple operands
            yield return
            [
                "A1:C4:IF(E10, D2, A10)",
                new[]
                {
                    // E10 is a value argument, thus isn't propagated, only added
                    new SheetArea("Sheet", Area.Parse("E10")),
                    // Areas from same sheet are unified into a single larger area
                    new SheetArea("Sheet", Area.Parse("A1:D10"))
                }
            ];

            // Range operator with multiple combinations
            yield return
            [
                "IF(G4,Sheet!A1,Other!A2):IF(H3,Other!C4,C5)",
                new[]
                {
                    // G4 and H3 are not propagated to range operation, only added
                    new SheetArea("Sheet", Area.Parse("G4")),
                    new SheetArea("Sheet", Area.Parse("H3")),

                    // Largest possible area in each sheet, based on references in the sheet
                    new SheetArea("Sheet", Area.Parse("A1:C5")),
                    new SheetArea("Other", Area.Parse("A2:C4"))
                }
            ];

            // Range operation when an argument isn't a reference doesn't
            // create a range from both, adds
            yield return
            [
                "INDEX({1},1,1):D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("D2")),
                }
            ];

            // Intersection - special case of one area against another area
            yield return
            [
                "A1:C3 B2:D2",
                new[]
                {
                    // In this special case, intersection is evaluated
                    new SheetArea("Sheet", Area.Parse("B2:C2")),
                }
            ];

            // Intersection - multi area operands. Due to complexity, keep
            // original ranges as dependencies.
            yield return
            [
                "A1:E10 IF(TRUE,A1:C3,B2:D2)",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1:C3")),
                    new SheetArea("Sheet", Area.Parse("B2:D2")),
                    new SheetArea("Sheet", Area.Parse("A1:E10")),
                }
            ];

            // Value binary operation on references adds the references
            yield return
            [
                "A1:B2 + A1:C4",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1:B2")),
                    new SheetArea("Sheet", Area.Parse("A1:C4")),
                }
            ];

            // IF function - value is added and true/false values are propagated
            yield return
            [
                "IF(A1,B1,C1):D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1")),
                    new SheetArea("Sheet", Area.Parse("B1:D2")),
                }
            ];

            // IF function, but only false argument is reference
            yield return
            [
                "IF(A1,5,B1):D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1")),
                    new SheetArea("Sheet", Area.Parse("B1:D2")),
                }
            ];

            // IF function, but only true argument is reference and is propagated
            yield return
            [
                "IF(A1,B1):D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1")),
                    new SheetArea("Sheet", Area.Parse("B1:D2")),
                }
            ];

            // INDEX function propagates whole range of first argument
            yield return
            [
                "INDEX(A1:C4,2,5):D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1:D4")),
                }
            ];

            // CHOOSE function adds first argument and propagates remaining arguments
            yield return
            [
                "CHOOSE(A1,B1,5,C1):D2",
                new[]
                {
                    new SheetArea("Sheet", Area.Parse("A1")),
                    new SheetArea("Sheet", Area.Parse("B1:D2")),
                }
            ];

            // Non-ref functions add arguments
            yield return
            [
                "POWER(SomeSheet!C4,Other!B1)",
                new[]
                {
                    new SheetArea("SomeSheet", Area.Parse("C4")),
                    new SheetArea("Other", Area.Parse("B1")),
                }
            ];
        }
    }

    public static IEnumerable<object[]> NameDependenciesTestCases
    {
        get
        {
            yield return
            [
                "WorkbookName  + 5",
                new[] { new XLName("WorkbookName") }
            ];

            yield return
            [
                "Sheet!Name",
                new[] { new XLName("Sheet", "Name") }
            ];
        }
    }
}
