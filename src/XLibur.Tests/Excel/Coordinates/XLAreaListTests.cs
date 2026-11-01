using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Coordinates;
// Ported from ClosedXML's XLAreaListTests, adapted to XLibur's Area/Point.
// Covers the value-typed sqref transforms that back conditional-format / data-validation
// coverage (see XLAreaList). Resource-file baseline comparisons are intentionally omitted.
internal class XLAreaListTests
{
    [Test]
    [Arguments("A1:C3", "A1", "B1:C3 A2:A4")]
    [Arguments("A1:C3", "B1", "A1:A3 C1:C3 B2:B4")]
    [Arguments("A1:C3", "C1", "A1:B3 C2:C4")]
    [Arguments("A1:C3", "A2", "A1:C1 B2:C3 A2:A4")]
    [Arguments("A1:C3", "B2", "A1:C1 A2:A3 C2:C3 B2:B4")]
    [Arguments("A1:C3", "C2", "A1:C1 A2:B3 C2:C4")]
    [Arguments("A1:C3", "A3", "A1:C2 B3:C3 A3:A4")]
    [Arguments("A1:C3", "B3", "A1:C2 A3 C3 B3:B4")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:B3 C3:C4")]
    [Arguments("B1:D3", "A1:A3", "B1:D3")] // Insert to left side - don't move
    [Arguments("A2:C4", "A1:C1", "A3:C5")] // Insert to top side - shift
    [Arguments("A2:C4", "A2:C2", "A3:C5")] // Insert to top edge - shift
    [Arguments("A2:C4", "A1", "B2:C4 A3:A5")] // Insert to top side - shift
    [Arguments("A1:C3", "D1:D3", "A1:C3")] // Insert to right side - don't move
    [Arguments("A1:C3", "A4:C5", "A1:C5")] // Insert to bottom edge - extend
    [Arguments("A1:C3", "A4", "A1:C3 A4")] // Insert to bottom side - extend
    [Arguments("A1:C3", "B4:E5", "A1:C3 B4:C5")] // Insert to bottom edge (inserted area out of bounds) - extend
    [Arguments("A1048576", "A1048576", "")] // Push out of sheet
    [Arguments("A1048575:A1048576", "A1048575", "A1048576")] // Partially push out of sheet
    [Arguments("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public async Task InsertAndShiftDown(string areaList, string insertedArea, string expected)
    {
        var list = new XLAreaList(Area.Parse(areaList));
        var result = list.InsertAndShiftDown(Area.Parse(insertedArea));

        await Assert.That(result.ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1:C3", "A1", "A2:C3 B1:D1")]
    [Arguments("A1:C3", "B1", "A2:C3 A1:D1")]
    [Arguments("A1:C3", "C1", "A2:C3 A1:D1")]
    [Arguments("A1:C3", "A2", "A1:C1 A3:C3 B2:D2")]
    [Arguments("A1:C3", "B2", "A1:C1 A3:C3 A2:D2")]
    [Arguments("A1:C3", "C2", "A1:C1 A3:C3 A2:D2")]
    [Arguments("A1:C3", "A3", "A1:C2 B3:D3")]
    [Arguments("A1:C3", "B3", "A1:C2 A3:D3")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:D3")]
    [Arguments("A1:C3", "A1:A3", "B1:D3")] // Insert to left edge - shift, don't extend
    [Arguments("A2:C4", "A1", "A2:C4")] // Insert to top side - don't move
    [Arguments("A1:C3", "D1:D3", "A1:D3")] // Insert to right edge - extend
    [Arguments("A1:C3", "D2:E10", "A1:C3 D2:E3")] // Insert to right edge (inserted area out of bounds) - extend
    [Arguments("A1:C3", "E1:E3", "A1:C3")] // Insert to right side - don't move
    [Arguments("A1:C3", "A4", "A1:C3")] // Insert to bottom side - don't move
    [Arguments("XFD1", "XFD1", "")] // Push out of sheet
    [Arguments("XFC1:XFD1", "XFC1", "XFD1")] // Partially push out of sheet
    [Arguments("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public async Task InsertAndShiftRight(string areaList, string insertedArea, string expected)
    {
        var list = new XLAreaList(Area.Parse(areaList));
        var result = list.InsertAndShiftRight(Area.Parse(insertedArea));

        await Assert.That(result.ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1:C3", "A1", "B1:C3 A1:A2")]
    [Arguments("A1:C3", "B1", "A1:A3 C1:C3 B1:B2")]
    [Arguments("A1:C3", "C1", "A1:B3 C1:C2")]
    [Arguments("A1:C3", "A2", "A1:C1 B2:C3 A2")]
    [Arguments("A1:C3", "B2", "A1:C1 A2:A3 C2:C3 B2")]
    [Arguments("A1:C3", "C2", "A1:C1 A2:B3 C2")]
    [Arguments("A1:C3", "A3", "A1:C2 B3:C3")]
    [Arguments("A1:C3", "B3", "A1:C2 A3 C3")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:B3")]
    [Arguments("B1:D3", "A1:A3", "B1:D3")] // Delete on the left side - don't move
    [Arguments("A2:C4", "A1:C1", "A1:C3")] // Delete on top side - shift
    [Arguments("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [Arguments("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move
    [Arguments("A1:A3", "A1:D5", "")] // Delete completely
    [Arguments("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public async Task DeleteAndShiftUp(string areaList, string deletedArea, string expected)
    {
        var list = new XLAreaList(Area.Parse(areaList));
        var result = list.DeleteAndShiftUp(Area.Parse(deletedArea));

        await Assert.That(result.ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1:C3", "A1", "A2:C3 A1:B1")]
    [Arguments("A1:C3", "B1", "A2:C3 A1 B1")]
    [Arguments("A1:C3", "C1", "A2:C3 A1:B1")]
    [Arguments("A1:C3", "A2", "A1:C1 A3:C3 A2:B2")]
    [Arguments("A1:C3", "B2", "A1:C1 A3:C3 A2 B2")]
    [Arguments("A1:C3", "C2", "A1:C1 A3:C3 A2:B2")]
    [Arguments("A1:C3", "A3", "A1:C2 A3:B3")]
    [Arguments("A1:C3", "B3", "A1:C2 A3 B3")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:B3")]
    [Arguments("B1:D3", "A1:A3", "A1:C3")] // Delete on the left side - shift
    [Arguments("A2:C4", "A1", "A2:C4")] // Delete on top side - don't move
    [Arguments("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [Arguments("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move
    [Arguments("A1:A3", "A1:D5", "")] // Delete completely
    [Arguments("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public async Task DeleteAndShiftLeft(string areaList, string deletedArea, string expected)
    {
        var list = new XLAreaList(Area.Parse(areaList));
        var result = list.DeleteAndShiftLeft(Area.Parse(deletedArea));

        await Assert.That(result.ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", true)]
    [Arguments("A1:C3", "B2", true)]
    [Arguments("B2:C3", "A2", false)]
    [Arguments("A1:C2 B3:C3", "A3", false)]
    public async Task IntersectsWith_determines_intersection_with_any_area(string areaListText, string areaText, bool expected)
    {
        var areaList = Parse(areaListText);
        var area = Area.Parse(areaText);
        await Assert.That(areaList.IntersectsWith(area)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", "A1")]
    [Arguments("A1:C3", "B2", "A1:C3")]
    [Arguments("A1:C3", "B2:D4", "A1:C3")]
    [Arguments("A1 C1", "A1:C1", "A1 C1")]
    [Arguments("A1 C1", "B1", "")]
    [Arguments("A1 C1", "B1:D2", "C1")]
    public async Task IntersectingWith_returns_areas_intersecting_with_the_other_area(string areaListText, string areaText, string expected)
    {
        var areaList = Parse(areaListText);
        var area = Area.Parse(areaText);
        await Assert.That(new XLAreaList(areaList.IntersectingWith(area).ToList()).ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "B1", "A1")]
    [Arguments("A1:E5", "C3:C4", "A1:E2 A5:E5 A3:B4 D3:E4")]
    [Arguments("B2:C5 B9 C4:D7", "C4:C5", "B2:C3 B4:B5 B9 C6:D7 D4:D5")]
    public async Task Excluding_returns_area_list_without_excluded(string areaListText, string excludedAreaText, string expected)
    {
        var areaList = Parse(areaListText);
        var excludedArea = Area.Parse(excludedAreaText);
        await Assert.That(areaList.Excluding(excludedArea).ToSpaceList()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", "A1", "A1")] // Copy from same point to the same point
    [Arguments("A1", "B5", "A1", "B5")] // Copy to different point
    [Arguments("B2", "D2", "A1:C3", "E3")] // Intersected area not in corner and shifted doesn't start at target
    [Arguments("D3:G6", "A1", "E4:F5", "A1:B2")]
    [Arguments("B2", "XFD1048576", "A1:C3", null)] // Copied area out of sheet
    public async Task TryCopyAreaTo_return_list_of_intersecting_areas_shifted_to_target(string areaListText, string targetPointText, string areaToCopyText, string? expected)
    {
        var areaList = Parse(areaListText);
        var targetPoint = Point.Parse(targetPointText);
        var areaToCopy = Area.Parse(areaToCopyText);
        await Assert.That(areaList.TryCopyAreaTo(targetPoint, areaToCopy, out var result) ? result.ToSpaceList() : null).IsEqualTo(expected);
    }

    private static XLAreaList Parse(string spaceList)
    {
        var list = new List<Area>();
        foreach (var reference in spaceList.Split(' '))
            list.Add(Area.Parse(reference));

        return new XLAreaList(list);
    }
}
