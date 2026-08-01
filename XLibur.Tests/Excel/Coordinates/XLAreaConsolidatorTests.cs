using System.Collections.Generic;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Coordinates;

internal class XLAreaConsolidatorTests
{
    [Test]
    [Arguments("", "")] // Empty stays empty
    [Arguments("B2:C3", "B2:C3")] // Single area passes through
    [Arguments("A1:B2 A3:B4", "A1:B4")] // Vertically adjacent, same columns - merge
    [Arguments("A1:B2 C1:D2", "A1:D2")] // Horizontally adjacent, same rows - merge
    [Arguments("A1:C2 B1:D2", "A1:D2")] // Overlapping - merge
    [Arguments("A1:C1 E1:G1 A3:C3 E3:G3", "A1:C1 E1:G1 A3:C3 E3:G3")] // Sparse - no merge
    public async Task Consolidate_merges_overlapping_and_adjacent_areas(string areaListText, string expected)
    {
        await Assert.That(Parse(areaListText).GetConsolidated().ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    public async Task Consolidate_matches_ClosedXML_baseline()
    {
        // Ported from ClosedXML RangesConsolidationTests.ConsolidateRangesSameWorksheet, whose
        // IXLRanges engine runs the same bitmask algorithm as XLAreaConsolidator.
        var input = Parse("A1:E3 A4:B10 E2:F12 C6:I8 G9 C9:D9 H9 I9:I13 C4:D5");

        var result = input.GetConsolidated().ToSpaceList();

        await Assert.That(result).IsEqualTo("A1:E9 F2:F12 G6:I9 A10:B10 E10:E12 I10:I13");
    }

    private static XLAreaList Parse(string spaceList)
    {
        if (spaceList.Length == 0)
            return XLAreaList.Empty;

        var list = new List<Area>();
        foreach (var reference in spaceList.Split(' '))
            list.Add(Area.Parse(reference));

        return new XLAreaList(list);
    }
}
