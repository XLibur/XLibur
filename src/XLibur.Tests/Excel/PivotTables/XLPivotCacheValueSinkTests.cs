using System;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// The records and the shared items of a cache field are both filled through one
/// <c>AddCellValue</c> dispatch (#621), so every cell value type must come out the same in both.
/// </summary>
public class XLPivotCacheValueSinkTests
{
    [Test]
    public async Task Every_cell_value_type_is_stored_alike_in_records_and_shared_items()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Field";
        // A2 stays blank.
        ws.Cell("A3").Value = true;
        ws.Cell("A4").Value = 1.5;
        ws.Cell("A5").Value = "text";
        ws.Cell("A6").Value = XLError.DivisionByZero;
        ws.Cell("A7").Value = new DateTime(2024, 1, 2);
        ws.Cell("A8").Value = new TimeSpan(14, 30, 0);

        var cache = (XLPivotCache)wb.PivotCaches.Add(ws.Range("A1:A8"));
        var fieldValues = cache.GetFieldValues(0);

        XLCellValue[] expected =
        [
            Blank.Value,
            true,
            1.5,
            "text",
            XLError.DivisionByZero,
            new DateTime(2024, 1, 2),
            // A TimeSpan is a plain OLE Automation date in the cache, see XLPivotCacheValue.ToCacheDateTime.
            new DateTime(1899, 12, 30, 14, 30, 0),
        ];

        await Assert.That(fieldValues.GetCellValues().ToArray()).IsEquivalentTo(expected, CollectionOrdering.Matching);
        await Assert.That(fieldValues.SharedItems.GetCellValues().ToArray()).IsEquivalentTo(expected, CollectionOrdering.Matching);
    }
}
