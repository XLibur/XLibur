using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Drawings;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// The order sheet listeners run in. Spec 33 replaces nine hardcoded calls with an enumeration;
/// this is what proves the enumeration did not reorder them.
/// <para>
/// The order the enumeration replaces, read off <c>XLWorksheetRangeShifter.Shift&lt;TAxis&gt;</c> at
/// the branch point (<c>806d69f7</c>), was:
/// </para>
/// <list type="number">
///   <item>merged-range straddle split</item>
///   <item>defined names, every sheet's, then the workbook's</item>
///   <item>conditional formats</item>
///   <item>data validations (sqref)</item>
///   <item>data-validation criteria formulas, every sheet's</item>
///   <item>page breaks</item>
///   <item>sparkline cleanup</item>
///   <item>calc engine</item>
///   <item>hyperlinks</item>
/// </list>
/// <para>
/// Spec 26 task 8 reconciled the row/column discrepancy in steps 6 and 7 — <c>ShiftColumns</c> ran
/// page breaks then sparklines and <c>ShiftRows</c> the reverse — having established that the two
/// commute. This test records that outcome; it does not decide it. One order now serves both axes,
/// so this test takes no axis parameter.
/// </para>
/// </summary>
public class SheetListenerOrderTests
{
    [Test]
    public async Task Sheet_listeners_run_in_the_pinned_order()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");

        var names = ws.GetSheetListeners().Select(l => l.GetType().Name).ToList();

        // CollectionOrdering.Matching, not the default: an order-insensitive assertion here would
        // pin the set and let the order this test exists to hold change underneath it.
        await Assert.That(names).IsEquivalentTo(new[]
        {
            nameof(MergedRangeSplitListener),
            nameof(XLDefinedNames),      // this sheet's
            nameof(XLDefinedNames),      // the workbook's
            nameof(XLConditionalFormats),
            nameof(XLDataValidations),   // sqref for this sheet, then criteria formulas
            nameof(XLPageSetup),
            nameof(XLSparklineGroups),
            nameof(XLCalcEngine),
            nameof(XLHyperlinks),
            nameof(DrawingAnchorListener),
            nameof(XLSheetView),
        }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// The workbook-scoped listeners are yielded once per sheet, so the enumeration grows with the
    /// workbook while the sheet-scoped ones stay at one apiece. Two sheets, so two
    /// <c>XLDefinedNames</c> entries plus the workbook's, and two <c>XLDataValidations</c>.
    /// </summary>
    [Test]
    public async Task Workbook_scoped_listeners_are_yielded_once_per_sheet()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        wb.AddWorksheet("T");

        var names = ws.GetSheetListeners().Select(l => l.GetType().Name).ToList();

        await Assert.That(names).IsEquivalentTo(new[]
        {
            nameof(MergedRangeSplitListener),
            nameof(XLDefinedNames),      // sheet S
            nameof(XLDefinedNames),      // sheet T
            nameof(XLDefinedNames),      // the workbook's
            nameof(XLConditionalFormats),
            nameof(XLDataValidations),   // sheet S
            nameof(XLDataValidations),   // sheet T
            nameof(XLPageSetup),
            nameof(XLSparklineGroups),
            nameof(XLCalcEngine),
            nameof(XLHyperlinks),
            nameof(DrawingAnchorListener),
            nameof(XLSheetView),
        }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// One listener per pivot table, matching how <c>XLWorksheets.GetWorkbookListeners</c> yields
    /// one per defined name. A sheet with no pivot tables yields none, which is why the two tests
    /// above do not list any.
    /// </summary>
    [Test]
    public async Task Every_pivot_table_on_the_sheet_is_a_listener()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        source.Cell("A1").Value = "Name";
        source.Cell("A2").Value = "Alice";
        source.Cell("B1").Value = "Amount";
        source.Cell("B2").Value = 1;

        var ws = (XLWorksheet)wb.AddWorksheet("Pivot");
        ws.PivotTables.Add("one", ws.Cell("A1")!, source.Range("A1:B2")!);
        ws.PivotTables.Add("two", ws.Cell("F1")!, source.Range("A1:B2")!);

        var names = ws.GetSheetListeners().Select(l => l.GetType().Name).ToList();

        await Assert.That(names.Count(n => n == nameof(XLPivotTable))).IsEqualTo(2);
        await Assert.That(names[^1]).IsEqualTo(nameof(XLPivotTable));
        await Assert.That(names[^3]).IsEqualTo(nameof(XLSheetView));
    }
}
