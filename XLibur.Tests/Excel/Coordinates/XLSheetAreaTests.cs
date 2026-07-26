using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Coordinates;

public class XLSheetAreaTests
{
    [Test]
    public async Task Sheet_name_is_compared_case_insensitive()
    {
        var upperCase = new XLBookArea("NAME", new XLSheetRange(1, 2, 3, 4));
        var lowerCase = new XLBookArea("name", new XLSheetRange(1, 2, 3, 4));
        await Assert.That(lowerCase.GetHashCode()).IsEqualTo(upperCase.GetHashCode());
        await Assert.That(lowerCase).IsEqualTo(upperCase);
    }

    [Test]
    public async Task Intersection_produces_range_intersection_in_same_sheet()
    {
        var sheetArea1 = new XLBookArea("SHEET", XLSheetRange.Parse("A1:C3"));
        var sheetArea2 = new XLBookArea("sheet", XLSheetRange.Parse("B2:D4"));
        var otherSheetArea = new XLBookArea("Other", XLSheetRange.Parse("B2:D4"));

        var sameSheetIntersection = sheetArea1.Intersect(sheetArea2);
        await Assert.That(sameSheetIntersection).IsEqualTo(new XLBookArea("sheet", XLSheetRange.Parse("B2:C3")));

        var differentSheetIntersection = sheetArea1.Intersect(otherSheetArea);
        await Assert.That(differentSheetIntersection).IsNull();
    }
}
