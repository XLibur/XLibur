using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.ConditionalFormats;
using XLibur.Tests.Excel.IO;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// A sheet is deleted or renamed through one door, the worksheet collection, and every holder of text
/// that names the sheet hears about it (spec 55). <c>SheetLifecycleFixtureTests</c> checks the holders
/// against workbooks Excel wrote.
/// </summary>
public class SheetLifecycleTests
{
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
    /// the workbook-scoped control became <c>#REF!</c>. The <c>delete-*</c> fixture shows Excel making
    /// a name scoped to another sheet <c>#REF!</c>.
    /// </summary>
    [Test]
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
    /// control <c>P</c> was renamed. The <c>rename-*</c> fixture shows Excel renaming all of them.
    /// </summary>
    [Test]
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
    /// collection, the workbook's names, each sheet's names, then each sheet's conditional formats,
    /// print areas and charts, and last the pivot caches.
    /// </summary>
    /// <remarks>
    /// Delete keeps the order rename already had. The calc engine renames its dependency tree on a
    /// rename, and drops the tree and marks every formula dirty on a delete. Neither reads the
    /// formula text the holders after it rewrite, and rewriting a formula marks it dirty on its own,
    /// so the engine and the holders commute on both events. No holder reads another's text, so the
    /// holders commute with each other too. Which names outlive a deleted sheet is the one question
    /// that reads across holders, and the door settles it before any of them hears of the delete.
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
        string Of(object holder, object ofS) => ReferenceEquals(holder, ofS) ? "S" : "T";
        string Label(IWorkbookListener listener) => listener switch
        {
            XLCalcEngine engine when ReferenceEquals(engine, wb.CalcEngine) => "calc engine",
            XLCellsCollection cells => $"cells of {Of(cells, s.Internals.CellsCollection)}",
            XLDefinedName name => $"name {name.Name} ({name.Scope})",
            XLConditionalFormats formats => $"conditional formats of {Of(formats, s.ConditionalFormats)}",
            XLPrintAreas areas => $"print areas of {Of(areas, s.PageSetup.PrintAreas)}",
            XLCharts charts => $"charts of {Of(charts, s.Charts)}",
            XLPivotCaches caches when ReferenceEquals(caches, wb.PivotCachesInternal) => "pivot caches",
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
            "conditional formats of S",
            "print areas of S",
            "charts of S",
            "conditional formats of T",
            "print areas of T",
            "charts of T",
            "pivot caches",
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

    /// <summary>
    /// A sheet's conditional formats, worst input: none at all, an expression the parser refuses, one
    /// that is already <c>#REF!</c>, a scale whose value is a formula by its type, and an unmodelled
    /// <c>x14</c> rule whose text the parser refuses.
    /// </summary>
    [Test]
    public async Task The_conditional_format_adapter_does_not_throw()
    {
        const string ruleId = "{00000000-0000-0000-0000-000000000001}";
        using var wb = new XLWorkbook();
        var empty = (XLWorksheet)wb.AddWorksheet("Empty");
        IXLWorksheet sheet = wb.AddWorksheet("Host");
        var host = (XLWorksheet)sheet;
        wb.AddWorksheet("Sheet1");
        sheet.Range("A1:A2").AddConditionalFormat().WhenIsTrue("=SUM(Sheet1!A1").Fill.SetBackgroundColor(XLColor.Red);
        sheet.Range("B1:B2").AddConditionalFormat().WhenIsTrue("=#REF!>0").Fill.SetBackgroundColor(XLColor.Red);
        sheet.Range("C1:C3").AddConditionalFormat().ColorScale()
            .Minimum(XLCFContentType.Formula, "Sheet1!$A$1", XLColor.Red);
        host.ConditionalFormats.SeedExtensionRuleFormulas(ruleId, ["SUM(Sheet1!A1", "Sheet1!#REF!"]);
        IWorkbookListener none = empty.ConditionalFormats;
        IWorkbookListener formats = host.ConditionalFormats;

        await Assert.That(() => none.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => none.OnSheetDeleting("Sheet1")).ThrowsNothing();
        await Assert.That(() => formats.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => formats.OnSheetDeleting("Renamed")).ThrowsNothing();

        // A refused formula is never rewritten (ADR 0002); the rest were.
        await Assert.That(host.ConditionalFormats.First().Values[1].Value).IsEqualTo("SUM(Sheet1!A1");
        await Assert.That(host.ConditionalFormats.Last().Values[1].Value).IsEqualTo("#REF!");
        await Assert.That(host.ConditionalFormats.TryGetExtensionRuleFormulas(ruleId, out var kept)).IsTrue();
        await Assert.That(kept).IsEquivalentTo(new[] { "SUM(Sheet1!A1", "#REF!" }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// A print area's worst input: one held as ranges, which has no formula text, and one whose
    /// formula the parser refuses.
    /// </summary>
    [Test]
    public async Task The_print_area_adapter_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var ranges = wb.AddWorksheet("Ranges");
        ranges.PageSetup.PrintAreas.Add("A1:B2");
        var refused = (XLPrintAreas)wb.AddWorksheet("Refused").PageSetup.PrintAreas;
        refused.FormulaReference = "OFFSET(Sheet1!$A$1";
        wb.AddWorksheet("Sheet1");
        IWorkbookListener[] listeners = [(XLPrintAreas)ranges.PageSetup.PrintAreas, refused];

        foreach (var listener in listeners)
        {
            await Assert.That(() => listener.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
            await Assert.That(() => listener.OnSheetDeleting("Renamed")).ThrowsNothing();
        }

        // A refused formula is never rewritten (ADR 0002).
        await Assert.That(refused.FormulaReference).IsEqualTo("OFFSET(Sheet1!$A$1");
    }

    /// <summary>
    /// A sheet's charts, worst input: none at all, a series whose reference the parser refuses, one
    /// already <c>#REF!</c>, and one with no references.
    /// </summary>
    [Test]
    public async Task The_chart_adapter_does_not_throw()
    {
        using var wb = new XLWorkbook();
        IWorkbookListener none = (XLCharts)wb.AddWorksheet("None").Charts;
        var host = wb.AddWorksheet("Host");
        wb.AddWorksheet("Sheet1");
        var chart = host.Charts.Add(XLChartType.ColumnClustered);
        chart.Series.Add("refused", "SUM(Sheet1!$A$1");
        chart.Series.Add("already", "#REF!", "Sheet1!#REF!");
        chart.Series.Add("empty", "");
        IWorkbookListener charts = (XLCharts)host.Charts;

        await Assert.That(() => none.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => none.OnSheetDeleting("Sheet1")).ThrowsNothing();
        await Assert.That(() => charts.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => charts.OnSheetDeleting("Renamed")).ThrowsNothing();

        // A refused reference is never rewritten (ADR 0002).
        await Assert.That(chart.Series.First().ValueReferences).IsEqualTo("SUM(Sheet1!$A$1");
        await Assert.That(chart.Series.ElementAt(1).CategoryReferences).IsEqualTo("#REF!");
    }

    /// <summary>
    /// The pivot caches' worst input: none at all, a cache whose source is a table, and one whose
    /// source is a range on the sheet renamed and then deleted.
    /// </summary>
    [Test]
    public async Task The_pivot_cache_adapter_does_not_throw()
    {
        using var empty = new XLWorkbook();
        IWorkbookListener none = empty.PivotCachesInternal;

        await Assert.That(() => none.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => none.OnSheetDeleting("Sheet1")).ThrowsNothing();

        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Sheet1");
        data.Cell("A1").Value = "Name";
        data.Cell("A2").Value = "a";
        data.Range("A1:A2").CreateTable("Names");
        data.Cell("C1").Value = "Amount";
        data.Cell("C2").Value = 1;
        wb.PivotCaches.Add(data.Range("A1:A2"));
        wb.PivotCaches.Add(data.Range("C1:C2"));
        IWorkbookListener caches = wb.PivotCachesInternal;

        await Assert.That(() => caches.OnSheetRenamed("Sheet1", "Renamed")).ThrowsNothing();
        await Assert.That(() => caches.OnSheetDeleting("Renamed")).ThrowsNothing();
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

    /// <summary>
    /// A sheet error on the deleted sheet becomes a plain <c>#REF!</c>, however the sheet's name is
    /// written, and never <c>#REF!#REF!</c>. Parser 4.0.0 writes it that way (fork #54), and the
    /// changelog's parser entry relies on it.
    /// </summary>
    [Test]
    [Arguments("Sheet1", "Sheet1!#REF!+1")]
    [Arguments("Sheet1", "'Sheet1'!#REF!+1")]
    [Arguments("Sheet 1", "'Sheet 1'!#REF!+1")]
    [Arguments("It's", "'It''s'!#REF!+1")]
    public async Task A_sheet_error_on_the_deleted_sheet_becomes_a_plain_REF(string sheetName, string formula)
    {
        using var wb = new XLWorkbook();
        var deleted = wb.AddWorksheet(sheetName);
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = formula;

        deleted.Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("#REF!+1");
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
    /// Deleting a sheet and adding one under the same name binds no formula and no name, at either
    /// scope, to the new sheet (acceptance criterion 5).
    /// </summary>
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
        sheet2.DefinedNames.Add("N", "Sheet1!$A$1");
        await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(10);

        Delete(wb, sheet1, throughCollection);
        wb.AddWorksheet("Sheet1").Cell("A1").Value = 7;

        await Assert.That(sheet2.Cell("A1").FormulaA1).IsEqualTo("#REF!*2");
        await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(XLError.CellReference);
        await Assert.That(RefersTo(wb, "W")).IsEqualTo("#REF!");
        await Assert.That(sheet2.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
    }

    /// <summary>
    /// A 3D reference with the deleted sheet at one end narrows to the sheets left, by tab order, as
    /// Excel does (Q34; the <c>delete-*</c> fixture). A sheet deleted from between the ends leaves the
    /// reference as it is. Part 1 of spec 55 left such a reference as it was; task 3 narrows it.
    /// </summary>
    [Test]
    [Arguments("Sheet1", "SUM(Sheet2:Sheet3!A1)")]
    [Arguments("Sheet2", "SUM(Sheet1:Sheet3!A1)")]
    [Arguments("Sheet3", "SUM(Sheet1:Sheet2!A1)")]
    public async Task A_3D_reference_narrows_when_the_sheet_at_one_end_is_deleted(string deleted, string expected)
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = "SUM(Sheet1:Sheet3!A1)";

        wb.Worksheet(deleted).Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo(expected);
    }

    /// <summary>
    /// A 3D reference narrowed to a single sheet is written as a reference to that sheet, quoted where
    /// the name needs it, as the <c>delete-*</c> fixture shows: <c>SUM(First:Last!$A$1)</c> became
    /// <c>SUM(Last!$A$1)</c>. Its ends can be written in either order. One with the deleted sheet at
    /// both ends has nothing left.
    /// </summary>
    [Test]
    [Arguments("SUM(Sheet1:Sheet2!A1)", "SUM(Sheet2!A1)")]
    [Arguments("SUM(Sheet2:Sheet1!A1)", "SUM(Sheet2!A1)")]
    [Arguments("SUM(Sheet1:Sheet3!A1)+Sheet1!B1", "SUM(Sheet2:Sheet3!A1)+#REF!")]
    [Arguments("SUM(Sheet1:Sheet1!A1)", "SUM(#REF!)")]
    public async Task A_3D_reference_narrowed_to_one_sheet_is_written_as_that_sheets(string formula, string expected)
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = formula;

        sheet1.Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo(expected);
    }

    /// <summary>
    /// A 3D reference narrowed to a sheet whose name needs quotes is quoted as a plain sheet
    /// reference is, the way the parser quotes one.
    /// </summary>
    [Test]
    public async Task A_3D_reference_narrowed_to_a_sheet_whose_name_needs_quotes_is_quoted()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("My Sheet");
        var host = wb.AddWorksheet("Host");
        host.Cell("A1").FormulaA1 = "SUM(Sheet1:'My Sheet'!A1)";

        sheet1.Delete();

        await Assert.That(host.Cell("A1").FormulaA1).IsEqualTo("SUM('My Sheet'!A1)");
    }

    /// <summary>
    /// The same narrowing in a defined name. Part 1 of spec 55 left the 3D reference as it was here.
    /// </summary>
    [Test]
    public async Task A_3D_reference_in_a_defined_name_narrows()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        wb.DefinedNames.Add("Mixed", "Sheet1!$A$1+SUM(Sheet1:Sheet3!$A$1)");

        sheet1.Delete();

        await Assert.That(RefersTo(wb, "Mixed")).IsEqualTo("#REF!+SUM(Sheet2:Sheet3!$A$1)");
    }

    /// <summary>
    /// A name scoped to the deleted sheet outlives it only if a cell formula on another sheet or a
    /// defined name refers to it. It moves to workbook scope, with the text the delete leaves it, and
    /// each reference to it, <c>Data!Used</c>, becomes <c>[0]!Used</c>. The <c>scoped-delete-*</c> and
    /// <c>delete-*</c> fixtures show both.
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_name_scoped_to_the_deleted_sheet_outlives_it_only_if_something_refers_to_it(bool throughCollection)
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var other = wb.AddWorksheet("Other");
        data.DefinedNames.Add("Alone", "Data!$A$1");
        data.DefinedNames.Add("Used", "Data!$B$1", "a comment");
        data.DefinedNames.Add("Named", "Data!$D$1");
        other.Cell("A1").FormulaA1 = "Data!Used";
        other.DefinedNames.Add("L", "Data!Named*2");

        Delete(wb, data, throughCollection);

        var names = wb.DefinedNames.Select(n => $"{n.Name} = {n.RefersTo}").Order(StringComparer.Ordinal).ToList();
        await Assert.That(names).IsEquivalentTo(new[] { "Named = #REF!", "Used = #REF!" }, CollectionOrdering.Matching);
        await Assert.That(wb.DefinedNames.Single(n => n.Name == "Used").Comment).IsEqualTo("a comment");
        await Assert.That(other.Cell("A1").FormulaA1).IsEqualTo("[0]!Used");
        await Assert.That(other.Cell("A1").Value).IsEqualTo(XLError.CellReference);
        await Assert.That(other.DefinedNames.Single().RefersTo).IsEqualTo("[0]!Named*2");
    }

    /// <summary>
    /// A reference to a name that outlived its sheet, <c>[0]!Used</c>, reads the name's value and
    /// follows its precedents. <c>[0]</c> is Excel's notation for this workbook. It used to read
    /// <c>#REF!</c> whatever the name referred to, because any book prefix was taken as another
    /// workbook, and a save cached that <c>#REF!</c>.
    /// </summary>
    [Test]
    public async Task A_reference_to_a_name_that_outlived_its_sheet_reads_the_names_value()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            data.DefinedNames.Add("Used", "Other!$B$1");
            other.Cell("B1").Value = 5;
            other.Cell("A1").FormulaA1 = "Data!Used";
            await Assert.That(other.Cell("A1").Value).IsEqualTo(5);

            data.Delete();

            await Assert.That(other.Cell("A1").FormulaA1).IsEqualTo("[0]!Used");
            await Assert.That(RefersTo(wb, "Used")).IsEqualTo("Other!$B$1");
            await Assert.That(other.Cell("A1").Value).IsEqualTo(5);

            other.Cell("B1").Value = 7;

            await Assert.That(other.Cell("A1").Value).IsEqualTo(7);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>()
                .Single(c => c.CellReference?.Value == "A1");
            await Assert.That(cell.CellFormula!.Text).IsEqualTo("[0]!Used");
            await Assert.That(cell.CellValue?.Text).IsEqualTo("7");
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Other").Cell("A1").FormulaA1).IsEqualTo("[0]!Used");
        await Assert.That(reloaded.Worksheet("Other").Cell("A1").Value).IsEqualTo(7);
    }

    /// <summary>
    /// <c>[0]!Name</c> loaded from a file reads the workbook-scoped name, and a change to the name's
    /// precedent reaches it.
    /// </summary>
    [Test]
    public async Task A_book_zero_name_loaded_from_a_file_reads_the_workbook_scoped_name()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var other = wb.AddWorksheet("Other");
            wb.DefinedNames.Add("Used", "Other!$B$1");
            other.Cell("B1").Value = 5;
            other.Cell("A1").FormulaA1 = "[0]!Used";
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        var sheet = reloaded.Worksheet("Other");
        reloaded.RecalculateAllFormulas();

        await Assert.That(sheet.Cell("A1").FormulaA1).IsEqualTo("[0]!Used");
        await Assert.That(sheet.Cell("A1").Value).IsEqualTo(5);

        sheet.Cell("B1").Value = 8;

        await Assert.That(sheet.Cell("A1").Value).IsEqualTo(8);
    }

    /// <summary>
    /// Unverified, and a call recorded in spec 55's Results: only cell formulas on the other sheets
    /// and defined names count as referring to a name scoped to the deleted sheet. A conditional
    /// format that refers to one does not keep it, and its reference becomes <c>#REF!</c> like any
    /// other. No fixture holds such a reference; a further fixture would settle it.
    /// </summary>
    [Test]
    public async Task Unverified_a_conditional_format_does_not_keep_a_scoped_name_alive()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var other = wb.AddWorksheet("Other");
        data.DefinedNames.Add("Used", "Data!$B$1");
        other.Range("C1:C1").AddConditionalFormat().WhenIsTrue("=Data!Used>0").Fill.SetBackgroundColor(XLColor.Red);

        data.Delete();

        await Assert.That(wb.DefinedNames.Any(n => n.Name == "Used")).IsFalse();
        await Assert.That(other.ConditionalFormats.Single().Values[1].Value).IsEqualTo("#REF!>0");
    }

    /// <summary>
    /// Unverified, and a call recorded in spec 55's Results: when a workbook-scoped name already holds
    /// the name, the workbook-scoped one is left as it is, the one scoped to the deleted sheet goes,
    /// and a reference to it becomes <c>#REF!</c>. In <c>scoped-delete-*</c> nothing referred to the
    /// sheet's <c>Clash</c>, so the fixture does not settle this.
    /// </summary>
    [Test]
    public async Task Unverified_a_workbook_scoped_name_of_the_same_name_keeps_its_place()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var other = wb.AddWorksheet("Other");
        data.DefinedNames.Add("Clash", "Data!$C$1");
        wb.DefinedNames.Add("Clash", "Other!$B$1");
        other.Cell("A1").FormulaA1 = "Data!Clash";

        data.Delete();

        await Assert.That(wb.DefinedNames.Select(n => $"{n.Name} = {n.RefersTo}"))
            .IsEquivalentTo(new[] { "Clash = Other!$B$1" }, CollectionOrdering.Matching);
        await Assert.That(other.Cell("A1").FormulaA1).IsEqualTo("#REF!");
    }

    /// <summary>
    /// A print area kept as formula text goes with its own sheet, as the <c>delete-*</c> fixture
    /// shows. On another sheet a print area is a name scoped to that sheet, and a reference in it to
    /// the deleted sheet becomes <c>#REF!</c>, as it does in any such name (D54). No fixture holds a
    /// print area on another sheet that refers to the deleted one, so that half follows the rule for
    /// names.
    /// </summary>
    [Test]
    public async Task A_print_area_goes_with_its_own_sheet()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            ((XLPrintAreas)data.PageSetup.PrintAreas).FormulaReference = "OFFSET(Data!$A$1,0,0,4,2)";
            var otherArea = (XLPrintAreas)other.PageSetup.PrintAreas;
            otherArea.FormulaReference = "OFFSET(Data!$A$1,0,0,4,2)";

            data.Delete();

            await Assert.That(otherArea.FormulaReference).IsEqualTo("OFFSET(#REF!,0,0,4,2)");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        var printAreas = document.WorkbookPart!.Workbook!.DefinedNames!.Elements<S.DefinedName>()
            .Where(n => n.Name == "_xlnm.Print_Area")
            .Select(n => n.Text)
            .ToList();
        await Assert.That(printAreas).IsEquivalentTo(new[] { "OFFSET(#REF!,0,0,4,2)" }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// An internal hyperlink keeps its location when the sheet it names is renamed or deleted, as the
    /// <c>rename-*</c> and <c>delete-*</c> fixtures show Excel doing. D67 is not a defect, and
    /// hyperlinks have no listener.
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task An_internal_hyperlink_keeps_its_location(bool delete)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            other.Cell("D1").SetHyperlink(new XLHyperlink("Data!A2"));

            if (delete)
                data.Delete();
            else
                data.Name = "Renamed";

            await Assert.That(other.Cell("D1").GetHyperlink().InternalAddress).IsEqualTo("Data!A2");
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Other").Cell("D1").GetHyperlink().InternalAddress).IsEqualTo("Data!A2");
    }

    /// <summary>
    /// Saving a loaded workbook after a sheet delete keeps a name on a sheet whose name ends with the
    /// deleted one's. A save-time pass dropped every name whose text held <c>Data!</c>, which also
    /// matches <c>OtherData!</c>; the workbook part's names were written again from the model right
    /// after it, so it never took effect, and it is gone.
    /// </summary>
    [Test]
    public async Task Saving_after_a_delete_keeps_a_name_on_a_sheet_whose_name_ends_with_the_deleted_one()
    {
        using var original = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data");
            wb.AddWorksheet("OtherData");
            wb.DefinedNames.Add("N", "OtherData!$A$1");
            wb.DefinedNames.Add("W", "Data!$A$1");
            wb.SaveAs(original);
        }

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(original))
        {
            wb.Worksheet("Data").Delete();
            wb.SaveAs(saved);
        }

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(RefersTo(reloaded, "N")).IsEqualTo("OtherData!$A$1");
        await Assert.That(RefersTo(reloaded, "W")).IsEqualTo("#REF!");
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
