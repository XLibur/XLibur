using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Tests.Excel.IO;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// A sheet is deleted or renamed through one door, the worksheet collection, and every holder of text
/// that names the sheet hears about it (spec 55).
/// </summary>
public class SheetLifecycleTests
{
    private const string BlockedOnTask3 =
        "Spec 55 task 3 (names at every scope) turns this green. It is blocked on Excel-authored "
        + "fixtures the owner has not made yet: delete-before.xlsx/delete-after.xlsx and "
        + "refdelete-before.xlsx/refdelete-after.xlsx in Resource/Other/SheetLifecycle.";

    /// <summary>
    /// D53: <c>wb.Worksheets.Delete</c> skipped <c>IsDeleted</c>, the name fix-up and the calc-engine
    /// purge, so a dependent kept the deleted sheet's value and saved it as its cached value.
    /// </summary>
    [Test]
    public async Task D53_the_collection_delete_does_everything_the_sheet_delete_does()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet1 = wb.AddWorksheet("Sheet1");
            var sheet2 = wb.AddWorksheet("Sheet2");
            sheet1.Cell("A1").Value = 5;
            sheet2.Cell("A1").FormulaA1 = "Sheet1!A1*2";
            wb.DefinedNames.Add("W", "Sheet1!$A$1");
            await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(10);

            wb.Worksheets.Delete("Sheet1");

            await Assert.That(((XLWorksheet)sheet1).IsDeleted).IsTrue();
            await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(XLError.CellReference);
            await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Sheet2").Cell("A1").Value).IsEqualTo(XLError.CellReference);
    }

    /// <summary>
    /// D54: a sheet-scoped name kept pointing at a deleted sheet, through a save and a reload, while
    /// the workbook-scoped control became <c>#REF!</c>.
    /// </summary>
    [Test]
    [Skip(BlockedOnTask3)]
    public async Task D54_a_sheet_scoped_name_on_another_sheet_loses_a_deleted_sheet()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            var sheet2 = wb.AddWorksheet("Sheet2");
            sheet2.DefinedNames.Add("N", "Sheet1!$A$1");
            wb.DefinedNames.Add("W", "Sheet1!$A$1");

            wb.Worksheet("Sheet1").Delete();

            await Assert.That(sheet2.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
            await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Sheet2").DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
    }

    /// <summary>
    /// D55: a rename skipped a sheet-qualified name and a 3D reference in a defined name, while the
    /// control <c>P</c> was renamed.
    /// </summary>
    [Test]
    [Skip(BlockedOnTask3)]
    public async Task D55_a_rename_reaches_sheet_qualified_names_and_3D_references_in_defined_names()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        sheet1.DefinedNames.Add("Local", "Sheet1!$B$1");
        wb.DefinedNames.Add("Q", "Sheet1!Local");
        wb.DefinedNames.Add("ThreeD", "SUM(Sheet1:Sheet3!$A$1)");
        wb.DefinedNames.Add("P", "Sheet1!$A$1");

        sheet1.Name = "Data";

        await Assert.That(RefersTo(wb, "Q")).IsEqualTo("Data!Local");
        await Assert.That(RefersTo(wb, "ThreeD")).IsEqualTo("SUM(Data:Sheet3!$A$1)");
        await Assert.That(RefersTo(wb, "P")).IsEqualTo("Data!$A$1");
    }

    /// <summary>
    /// The order holders hear of a rename or a delete: the calc engine, each sheet's cells
    /// collection, the workbook's names, then each sheet's names.
    /// </summary>
    /// <remarks>
    /// Delete keeps the order rename already had. The calc engine renames its dependency tree on a
    /// rename, and drops the tree and marks every formula dirty on a delete. Neither reads the
    /// formula text the holders after it rewrite, and rewriting a formula marks it dirty on its own,
    /// so the engine and the holders commute on both events.
    /// </remarks>
    [Test]
    public async Task Workbook_listeners_run_in_the_pinned_order()
    {
        using var wb = new XLWorkbook();
        var s = (XLWorksheet)wb.AddWorksheet("S");
        var t = (XLWorksheet)wb.AddWorksheet("T");
        wb.DefinedNames.Add("W", "S!$A$1");
        s.DefinedNames.Add("LS", "T!$A$1");
        t.DefinedNames.Add("LT", "S!$A$1");

        // Each listener is named by reference, not compared as an object: an equivalence assertion
        // over the objects compares them member by member, which is not the question here.
        string Label(IWorkbookListener listener) => listener switch
        {
            XLCalcEngine engine when ReferenceEquals(engine, wb.CalcEngine) => "calc engine",
            XLCellsCollection cells when ReferenceEquals(cells, s.Internals.CellsCollection) => "cells of S",
            XLCellsCollection cells when ReferenceEquals(cells, t.Internals.CellsCollection) => "cells of T",
            XLDefinedName name => $"name {name.Name} ({name.Scope})",
            _ => listener.GetType().Name,
        };

        var labels = wb.WorksheetsInternal.GetWorkbookListeners().Select(Label).ToList();

        // CollectionOrdering.Matching, so that the order this test exists to hold cannot change
        // underneath an order-insensitive assertion.
        await Assert.That(labels).IsEquivalentTo(new[]
        {
            "calc engine",
            "cells of S",
            "cells of T",
            "name W (Workbook)",
            "name LS (Worksheet)",
            "name LT (Worksheet)",
        }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// The calc engine's worst input: an engine in a workbook with no sheets, which has never built
    /// a dependency tree, and one whose tree holds a formula that is already <c>#REF!</c>.
    /// </summary>
    /// <remarks>
    /// A formula the parser refuses cannot be in a dependency tree: building the tree parses every
    /// formula in the workbook, and throws on such a formula. The engine never reads formula text on
    /// a rename or a delete, so the refused formula is the cells collection's worst input, not the
    /// engine's.
    /// </remarks>
    [Test]
    public async Task The_calc_engine_adapter_does_not_throw()
    {
        using var empty = new XLWorkbook();
        IWorkbookListener idle = empty.CalcEngine;

        await Assert.That(() => idle.OnSheetDeleting("Sheet1")).ThrowsNothing();
        await Assert.That(() => idle.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "#REF!+1";
        ws.Cell("A2").FormulaA1 = "Sheet1!#REF!*2";
        ws.Cell("A3").FormulaA1 = "A1*2";
        wb.RecalculateAllFormulas();
        IWorkbookListener engine = wb.CalcEngine;

        await Assert.That(() => engine.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => engine.OnSheetDeleting("Sheet1")).ThrowsNothing();
    }

    /// <summary>
    /// A cells collection's worst input: none at all, a formula the parser refuses, formulas that are
    /// already <c>#REF!</c>, an array formula and a 3D reference.
    /// </summary>
    [Test]
    public async Task The_cells_collection_adapter_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var empty = (XLWorksheet)wb.AddWorksheet("Empty");
        var host = wb.AddWorksheet("Host");
        wb.AddWorksheet("Sheet1");
        host.Cell("A1").FormulaA1 = "'[Book2.xlsx]Sheet1'!A1";
        host.Cell("A2").FormulaA1 = "Sheet1!#REF!+1";
        host.Cell("A3").FormulaA1 = "#REF!*2";
        host.Cell("A4").FormulaA1 = "SUM(Empty:Sheet1!A1)";
        host.Range("B1:B2").FormulaArrayA1 = "Sheet1!A1:A2*2";
        IWorkbookListener none = empty.Internals.CellsCollection;
        IWorkbookListener cells = ((XLWorksheet)host).Internals.CellsCollection;

        await Assert.That(() => none.OnSheetDeleting("Sheet1")).ThrowsNothing();
        await Assert.That(() => none.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => cells.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => cells.OnSheetDeleting("Renamed")).ThrowsNothing();

        // A refused formula is never rewritten (ADR 0002).
        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("'[Book2.xlsx]Sheet1'!A1");
    }

    /// <summary>
    /// A defined name's worst input: a name the parser refuses, one that is already <c>#REF!</c>,
    /// and one whose <c>#REF!</c> still carries the prefix of the sheet being deleted, at both
    /// scopes.
    /// </summary>
    [Test]
    public async Task The_defined_name_adapter_does_not_throw()
    {
        using var package = BookWithRefusedName("SUM(Sheet1!$A$1");
        using var wb = new XLWorkbook(package);
        var other = (XLWorksheet)wb.Worksheet("Other");
        wb.DefinedNames.Add("AlreadyRef", "#REF!");
        wb.DefinedNames.Add("PrefixedRef", "Sheet1!#REF!");
        other.DefinedNames.Add("Local", "Sheet1!#REF!");
        var names = wb.DefinedNamesInternal.Cast<IWorkbookListener>()
            .Concat(other.DefinedNames.Cast<IWorkbookListener>())
            .ToList();

        foreach (var name in names)
        {
            await Assert.That(() => name.OnSheetDeleting("Sheet1")).ThrowsNothing();
            await Assert.That(() => name.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        }

        // A refused formula is never rewritten (ADR 0002).
        await Assert.That(RefersTo(wb, "x")).IsEqualTo("SUM(Sheet1!$A$1");
    }

    [Test]
    public async Task A_rename_changes_the_key_and_the_name_together()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Before");

        ws.Name = "After";

        await Assert.That(ws.Name).IsEqualTo("After");
        await Assert.That(wb.Worksheet("After")).IsSameReferenceAs(ws);
        await Assert.That(wb.Worksheets.Contains("Before")).IsFalse();
    }

    /// <summary>
    /// Renaming a sheet that has been deleted changes only what it is called. The setter used to
    /// look the sheet up by its old name, so it found a sheet added since under that name, moved that
    /// sheet's key to the new name without changing the sheet's own name, and rewrote every formula
    /// that referred to it.
    /// </summary>
    [Test]
    public async Task Renaming_a_deleted_sheet_leaves_a_new_sheet_of_the_same_name_alone()
    {
        using var wb = new XLWorkbook();
        var old = wb.AddWorksheet("Sheet1");
        var keep = wb.AddWorksheet("Keep");
        old.Delete();
        var fresh = wb.AddWorksheet("Sheet1");
        keep.Cell("A1").FormulaA1 = "Sheet1!A1";

        old.Name = "Gone";

        await Assert.That(old.Name).IsEqualTo("Gone");
        await Assert.That(fresh.Name).IsEqualTo("Sheet1");
        await Assert.That(wb.Worksheet("Sheet1")).IsSameReferenceAs(fresh);
        await Assert.That(wb.Worksheets.Contains("Gone")).IsFalse();
        await Assert.That(keep.Cell("A1").FormulaA1).IsEqualTo("Sheet1!A1");
    }

    /// <summary>
    /// A reference to a deleted sheet in a cell formula becomes <c>#REF!</c>, as in Excel, through
    /// either door. Parser 4.0.0 writes a deleted sheet's <c>Sheet1!#REF!</c> as a plain <c>#REF!</c>.
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_cell_formula_that_points_at_a_deleted_sheet_is_rewritten_to_REF(bool throughCollection)
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var host = wb.AddWorksheet("Host");
        wb.AddWorksheet("Sheet 3");
        host.Cell("A1").FormulaA1 = "Sheet1!A1*2";
        host.Cell("A2").FormulaA1 = "SUM(Sheet1!A1:B2)";
        host.Cell("A3").FormulaA1 = "'Sheet 3'!A1+Sheet1!A1";
        host.Cell("A4").FormulaA1 = "Sheet1!#REF!+1";
        host.Cell("A5").FormulaA1 = "sheet1!A1";
        host.Cell("A6").FormulaA1 = "B1+1";

        Delete(wb, sheet1, throughCollection);

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("#REF!*2");
        await Assert.That(host.Cell("A2").FormulaA1).IsEqualTo("SUM(#REF!)");
        await Assert.That(host.Cell("A3").FormulaA1).IsEqualTo("'Sheet 3'!A1+#REF!");
        await Assert.That(host.Cell("A4").FormulaA1).IsEqualTo("#REF!+1");
        await Assert.That(host.Cell("A5").FormulaA1).IsEqualTo("#REF!");
        await Assert.That(host.Cell("A6").FormulaA1).IsEqualTo("B1+1");
        await Assert.That(host.Cell("A1").Value).IsEqualTo(XLError.CellReference);
    }

    /// <summary>A formula the parser refuses keeps its text when a sheet is deleted (ADR 0002).</summary>
    [Test]
    public async Task A_refused_formula_keeps_its_text_when_a_sheet_is_deleted()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = "SUM(Sheet1!A1";
        host.Cell("A2").FormulaA1 = "'[Book2.xlsx]Sheet1'!A1";

        sheet1.Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("SUM(Sheet1!A1");
        await Assert.That(host.Cell("A2").FormulaA1).IsEqualTo("'[Book2.xlsx]Sheet1'!A1");
    }

    /// <summary>
    /// Deleting a sheet and adding one under the same name binds neither a formula nor a
    /// workbook-scoped name to the new sheet.
    /// </summary>
    /// <remarks>
    /// A sheet-scoped name on another sheet still rebinds, because a delete does not reach it yet
    /// (D54). Spec 55 task 3 covers it.
    /// </remarks>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_sheet_added_under_a_deleted_sheets_name_does_not_rebind_the_old_references(bool throughCollection)
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet1.Cell("A1").Value = 5;
        sheet2.Cell("A1").FormulaA1 = "Sheet1!A1*2";
        wb.DefinedNames.Add("W", "Sheet1!$A$1");
        await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(10);

        Delete(wb, sheet1, throughCollection);
        wb.AddWorksheet("Sheet1").Cell("A1").Value = 7;

        await Assert.That(sheet2.Cell("A1").FormulaA1).IsEqualTo("#REF!*2");
        await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(XLError.CellReference);
        await Assert.That(RefersTo(wb, "W")).IsEqualTo("#REF!");
    }

    /// <summary>
    /// A 3D reference that touches the deleted sheet keeps its text in part 1 of spec 55. Excel
    /// narrows the reference when an end sheet is deleted (Q34), and spec 55 task 3 implements that
    /// against an Excel-authored fixture. Until then the reference must not collapse to <c>#REF!</c>,
    /// which is what the parser's default rewrite does.
    /// </summary>
    [Test]
    [Arguments("Sheet1")]
    [Arguments("Sheet2")]
    [Arguments("Sheet3")]
    public async Task A_3D_reference_to_a_deleted_sheet_keeps_its_text_until_spec_55_task_3(string deleted)
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = "SUM(Sheet1:Sheet3!A1)";

        wb.Worksheet(deleted).Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("SUM(Sheet1:Sheet3!A1)");
    }

    /// <summary>
    /// The same holds in a defined name. A name's rewrite reaches a 3D reference only when the name
    /// also refers to the deleted sheet directly, and the parser's default then collapsed the 3D
    /// reference to <c>#REF!</c>, giving <c>#REF!+SUM(#REF!)</c>. Spec 55 task 3 narrows it.
    /// </summary>
    [Test]
    public async Task A_3D_reference_in_a_defined_name_keeps_its_text_until_spec_55_task_3()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        wb.DefinedNames.Add("Mixed", "Sheet1!$A$1+SUM(Sheet1:Sheet3!$A$1)");

        sheet1.Delete();

        await Assert.That(RefersTo(wb, "Mixed")).IsEqualTo("#REF!+SUM(Sheet1:Sheet3!$A$1)");
    }

    /// <summary>Text inside a string is never rewritten (spec 55 non-goals).</summary>
    [Test]
    public async Task Text_inside_a_string_is_not_rewritten_by_a_rename_or_a_delete()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = "INDIRECT(\"Sheet1!A1\")";
        host.Cell("A2").FormulaA1 = "INDIRECT(\"Sheet2!A1\")";

        sheet1.Name = "Renamed";
        sheet2.Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("INDIRECT(\"Sheet1!A1\")");
        await Assert.That(host.Cell("A2").FormulaA1).IsEqualTo("INDIRECT(\"Sheet2!A1\")");
    }

    /// <summary>
    /// A clash with an unsupported sheet is found ignoring case, as sheet names are compared
    /// everywhere else: a workbook cannot hold both <c>Chart</c> and <c>CHART</c>.
    /// </summary>
    [Test]
    public async Task A_chartsheets_name_is_refused_in_any_case()
    {
        using var wb = OpenChartsheetBook();

        await Assert.That(() => wb.Worksheets.Add("CHART")).Throws<ArgumentException>();
        await Assert.That(() => wb.Worksheet("Data").Name = "chart").Throws<ArgumentException>();
    }

    /// <summary>
    /// A refused name leaves the tab order as it was. The positional add used to move every sheet at
    /// or after the position, and only then find that the name was taken.
    /// </summary>
    [Test]
    public async Task A_refused_add_at_a_position_moves_no_sheet()
    {
        using var wb = OpenChartsheetBook();
        var before = Positions(wb);

        await Assert.That(() => wb.Worksheets.Add("Chart", 1)).Throws<ArgumentException>();
        await Assert.That(() => wb.Worksheets.Add("Data", 1)).Throws<ArgumentException>();

        await Assert.That(Positions(wb)).IsEquivalentTo(before, CollectionOrdering.Matching);
    }

    /// <summary>
    /// <c>Add()</c> without a name passes over a name that an unsupported sheet holds, rather than
    /// choosing it and then refusing it.
    /// </summary>
    [Test]
    public async Task Add_without_a_name_skips_a_name_an_unsupported_sheet_holds()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.UnsupportedSheets.Add(new XLWorkbook.UnsupportedSheet { Name = "Sheet2", Position = 2, SheetId = 99 });

        var added = wb.Worksheets.Add();

        await Assert.That(added.Name).IsEqualTo("Sheet3");
    }

    /// <summary>
    /// <c>IXLWorksheet.Delete()</c> deletes the sheet it is called on, not whichever sheet has its name
    /// now. Called again on a deleted sheet, after a sheet was added under the same name, it used to
    /// delete the new sheet: rewrite every formula pointing at it to <c>#REF!</c>, make a name pointing
    /// at it <c>#REF!</c>, and mark it deleted.
    /// </summary>
    [Test]
    public async Task Deleting_a_deleted_sheet_again_leaves_a_new_sheet_of_the_same_name_alone()
    {
        using var wb = new XLWorkbook();
        var old = wb.AddWorksheet("Data");
        var keep = wb.AddWorksheet("Keep");
        old.Delete();
        var fresh = wb.AddWorksheet("Data");
        fresh.Cell("A1").Value = 7;
        keep.Cell("A1").FormulaA1 = "Data!A1*2";
        wb.DefinedNames.Add("W", "Data!$A$1");

        old.Delete();

        await Assert.That(((XLWorksheet)fresh).IsDeleted).IsFalse();
        await Assert.That(wb.Worksheet("Data")).IsSameReferenceAs(fresh);
        await Assert.That(keep.Cell("A1").FormulaA1).IsEqualTo("Data!A1*2");
        await Assert.That(keep.Cell("A1").Value).IsEqualTo(14);
        await Assert.That(RefersTo(wb, "W")).IsEqualTo("Data!$A$1");
    }

    /// <summary>
    /// Deleting a sheet that is already deleted does nothing, as renaming one changes nothing in the
    /// workbook. It used to throw <c>KeyNotFoundException</c>.
    /// </summary>
    [Test]
    public async Task Deleting_a_deleted_sheet_again_does_nothing()
    {
        using var wb = new XLWorkbook();
        var old = wb.AddWorksheet("Data");
        var keep = wb.AddWorksheet("Keep");
        old.Delete();

        await Assert.That(() => old.Delete()).ThrowsNothing();
        await Assert.That(((XLWorksheet)old).IsDeleted).IsTrue();
        await Assert.That(wb.Worksheets.Count).IsEqualTo(1);
        await Assert.That(keep.Position).IsEqualTo(1);
    }

    /// <summary>
    /// The collection names a sheet by the name it has now. After the original <c>Data</c> was deleted
    /// and a new one added, <c>wb.Worksheets.Delete("Data")</c> deletes the new one.
    /// </summary>
    [Test]
    public async Task The_collection_delete_by_name_deletes_the_sheet_that_has_the_name_now()
    {
        using var wb = new XLWorkbook();
        var old = wb.AddWorksheet("Data");
        wb.AddWorksheet("Keep");
        old.Delete();
        var fresh = wb.AddWorksheet("Data");

        wb.Worksheets.Delete("Data");

        await Assert.That(((XLWorksheet)fresh).IsDeleted).IsTrue();
        await Assert.That(wb.Worksheets.Contains("Data")).IsFalse();
    }

    private static XLWorkbook OpenChartsheetBook()
        => new(TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx")));

    private static string[] Positions(XLWorkbook wb)
        => wb.Worksheets.Select(w => $"{w.Name}:{w.Position}")
            .Concat(wb.UnsupportedSheets.Select(s => $"{s.Name}:{s.Position}"))
            .ToArray();

    private static void Delete(XLWorkbook wb, IXLWorksheet sheet, bool throughCollection)
    {
        if (throughCollection)
            wb.Worksheets.Delete(sheet.Name);
        else
            sheet.Delete();
    }

    private static string RefersTo(XLWorkbook wb, string name)
        => wb.DefinedNames.Single(n => n.Name == name).RefersTo;

    /// <summary>
    /// A workbook with sheets <c>Sheet1</c> and <c>Other</c>, and a workbook-scoped name <c>x</c>
    /// whose text is spliced into the file, so the name can hold text the parser refuses.
    /// </summary>
    private static MemoryStream BookWithRefusedName(string refersToText)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = 1;
            wb.AddWorksheet("Other");
            wb.SaveAs(package);
        }

        return package.RewriteWorkbook(xml =>
        {
            var rewritten = xml.Replace("<x:definedNames />",
                $"<x:definedNames><x:definedName name=\"x\">{refersToText}</x:definedName></x:definedNames>");
            if (!rewritten.Contains("definedName name=\"x\"", StringComparison.Ordinal))
                throw new InvalidOperationException("The defined name was not spliced into the workbook part.");

            return rewritten;
        });
    }
}
