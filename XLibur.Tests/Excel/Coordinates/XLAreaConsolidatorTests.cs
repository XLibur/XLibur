using System.Collections.Generic;
using XLibur.Excel.Coordinates;
using NUnit.Framework;

namespace XLibur.Tests.Excel.Coordinates;

[TestFixture]
internal class XLAreaConsolidatorTests
{
    [TestCase("", ExpectedResult = "")] // Empty stays empty
    [TestCase("B2:C3", ExpectedResult = "B2:C3")] // Single area passes through
    [TestCase("A1:B2 A3:B4", ExpectedResult = "A1:B4")] // Vertically adjacent, same columns - merge
    [TestCase("A1:B2 C1:D2", ExpectedResult = "A1:D2")] // Horizontally adjacent, same rows - merge
    [TestCase("A1:C2 B1:D2", ExpectedResult = "A1:D2")] // Overlapping - merge
    [TestCase("A1:C1 E1:G1 A3:C3 E3:G3", ExpectedResult = "A1:C1 E1:G1 A3:C3 E3:G3")] // Sparse - no merge
    public string Consolidate_merges_overlapping_and_adjacent_areas(string areaListText)
    {
        return Parse(areaListText).GetConsolidated().ToSpaceList();
    }

    [Test]
    public void Consolidate_matches_ClosedXML_baseline()
    {
        // Ported from ClosedXML RangesConsolidationTests.ConsolidateRangesSameWorksheet, whose
        // IXLRanges engine runs the same bitmask algorithm as XLAreaConsolidator.
        var input = Parse("A1:E3 A4:B10 E2:F12 C6:I8 G9 C9:D9 H9 I9:I13 C4:D5");

        var result = input.GetConsolidated().ToSpaceList();

        Assert.AreEqual("A1:E9 F2:F12 G6:I9 A10:B10 E10:E12 I10:I13", result);
    }

    private static XLAreaList Parse(string spaceList)
    {
        if (spaceList.Length == 0)
            return XLAreaList.Empty;

        var list = new List<XLSheetRange>();
        foreach (var reference in spaceList.Split(' '))
            list.Add(XLSheetRange.Parse(reference));

        return new XLAreaList(list);
    }
}
