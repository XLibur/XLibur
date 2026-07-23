using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Tests.Excel.ConditionalFormats;

// Regression coverage for ClosedXML issue #2850: inserting rows/columns must shift every
// conditional-format (and data-validation) range by exactly the inserted amount. The original
// defect doubled the offset for rules whose shifted target address collided with another
// existing rule's range, because the range repository handed back the same aliased instance
// which was then shifted a second time by the blanket auto-shift.
[TestFixture]
public class ConditionalFormatRangeShiftTests
{
    [Test]
    public void InsertRowsAbove_ShiftsCfRangesByExactAmount()
    {
        // Layout from the issue: multiple rules per row, and rows whose shifted target collides
        // with another rule's row (13+10 == 23, which already hosts rules).
        var rows = new[] { 12, 12, 12, 13, 13, 13, 16, 16, 16, 17, 17, 17, 23, 23, 31, 31, 32, 34, 34, 35 };
        const int inserted = 10;

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        foreach (var r in rows)
            ws.Cell(r, 11 /* K */).AddConditionalFormat().WhenGreaterThan(1).Fill.SetBackgroundColor(XLColor.Red);

        ws.Row(13).InsertRowsAbove(inserted);

        var expected = rows.Select(r => r < 13
            ? $"K12:K{12 + inserted}"      // row above the insertion expands to swallow inserted rows
            : $"K{r + inserted}:K{r + inserted}"); // rows at/below the insertion move down

        var actual = ws.ConditionalFormats.Select(cf => cf.Ranges.Single().RangeAddress.ToString());
        Assert.That(actual, Is.EqualTo(expected.ToList()));
    }

    [Test]
    public void InsertColumnsBefore_ShiftsCfRangesByExactAmount()
    {
        // Column analogue: C+10 == M, which already hosts rules.
        var cols = new[] { 2, 2, 3, 3, 6, 13, 13, 15 }; // B, C, F, M, O
        const int inserted = 10;

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        foreach (var c in cols)
            ws.Cell(20, c).AddConditionalFormat().WhenGreaterThan(1).Fill.SetBackgroundColor(XLColor.Red);

        ws.Column(3).InsertColumnsBefore(inserted);

        var actual = ws.ConditionalFormats
            .Select(cf => cf.Ranges.Single().RangeAddress)
            .Select(a => (a.FirstAddress.ColumnNumber, a.LastAddress.ColumnNumber))
            .ToList();

        var expected = cols.Select(c => c < 3
            ? (c, c + inserted)          // column before the insertion expands
            : (c + inserted, c + inserted)) // columns at/after the insertion move right
            .ToList();

        Assert.That(actual, Is.EqualTo(expected));
    }

    [Test]
    public void InsertRowsAbove_ShiftsMultiAreaCf_ExtendsAndShiftsTogether()
    {
        // A single CF covering two disjoint areas. Inserting rows inside the first must extend it,
        // while the second (below the insertion) shifts down — the value-typed area transform
        // handles both in one pass.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cf = ws.Range("A5:A7").AddConditionalFormat();
        cf.WhenGreaterThan(1).Fill.SetBackgroundColor(XLColor.Red);
        // Coverage is stored as an XLAreaList; extend it with a second disjoint area.
        ((XLConditionalFormat)cf).SetAreas(new XLAreaList(new List<XLSheetRange>
        {
            XLSheetRange.Parse("A5:A7"),
            XLSheetRange.Parse("C10:C12"),
        }));

        ws.Row(6).InsertRowsAbove(3);

        var areas = ws.ConditionalFormats.Single().Ranges
            .Select(r => r.RangeAddress.ToString())
            .OrderBy(s => s)
            .ToList();

        // A5:A7 spans the insertion at row 6 -> extends to A5:A10; C10:C12 is below -> C13:C15.
        Assert.That(areas, Is.EqualTo(new[] { "A5:A10", "C13:C15" }));
    }
}
