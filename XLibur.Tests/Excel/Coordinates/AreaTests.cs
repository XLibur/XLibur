using System;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Coordinates;

public class AreaTests
{
    [Test]
    [Arguments("A1", 1, 1, 1, 1)]
    [Arguments("A1:Z100", 1, 1, 100, 26)]
    [Arguments("BD14:EG256", 14, 56, 256, 137)]
    [Arguments("A1:XFD1048576", 1, 1, 1048576, 16384)]
    [Arguments("XFD1048576", 1048576, 16384, 1048576, 16384)]
    [Arguments("XFD1048576:XFD1048576", 1048576, 16384, 1048576, 16384)]
    public async Task ParseCellRefsAccordingToGrammar(string refText, int firstRow, int firstCol, int lastRow, int lastCol)
    {
        var reference = Area.Parse(refText);
        await Assert.That(reference.FirstPoint.Row).IsEqualTo(firstRow);
        await Assert.That(reference.FirstPoint.Column).IsEqualTo(firstCol);
        await Assert.That(reference.LastPoint.Row).IsEqualTo(lastRow);
        await Assert.That(reference.LastPoint.Column).IsEqualTo(lastCol);
    }

    [Test]
    [Arguments("")]
    [Arguments("A1:")]
    [Arguments(":A1")]
    [Arguments("A1: A1")]
    [Arguments(" A1:A1")]
    [Arguments("A1:A1 ")]
    [Arguments("B1:A1")]
    [Arguments("A2:A1")]
    public async Task InvalidInputsAreNotParsed(string invalidRef)
    {
        await Assert.That(() => Area.Parse(invalidRef)).Throws<FormatException>();
    }

    [Test]
    [Arguments("A1:A1", "A1")]
    [Arguments("DO974:LAR2487", "DO974:LAR2487")]
    [Arguments("XFD1048576:XFD1048576", "XFD1048576")]
    [Arguments("XFD1048575:XFD1048576", "XFD1048575:XFD1048576")]
    public async Task CanFormatToString(string cellRef, string expected)
    {
        var r = Area.Parse(cellRef);
        await Assert.That(r.ToString()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", "A1")]
    [Arguments("A1", "B3", "A1:B3")]
    [Arguments("C2", "B3", "B2:C3")]
    [Arguments("I6:J9", "L7", "I6:L9")]
    [Arguments("B2:B4", "A3:C3", "A2:C4")]
    [Arguments("B2:C3", "E5:F6", "B2:F6")]
    public async Task RangeOperation(string leftOperand, string rightOperand, string expectedRange)
    {
        var left = Area.Parse(leftOperand);
        var right = Area.Parse(rightOperand);
        var expected = Area.Parse(expectedRange);

        await Assert.That(left.Range(right)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", "A1")]
    [Arguments("A1", "A2", null)]
    [Arguments("B1:B3", "A2:C2", "B2")]
    [Arguments("A1:A3", "B2:C2", null)]
    [Arguments("A1:D6", "B2:C3", "B2:C3")]
    [Arguments("A1:C6", "B4:E10", "B4:C6")]
    public async Task IntersectOperation(string leftOperand, string rightOperand, string expectedRange)
    {
        var left = Area.Parse(leftOperand);
        var right = Area.Parse(rightOperand);
        var expected = expectedRange is null ? (Area?)null : Area.Parse(expectedRange);

        await Assert.That(left.Intersect(right)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", true)]
    [Arguments("A1", "A2", false)]
    [Arguments("B1:B3", "A2:C2", true)]
    [Arguments("A1:A3", "B2:C2", false)]
    [Arguments("A1:D6", "B2:C3", true)]
    [Arguments("A1:C6", "B4:E10", true)]
    public async Task Intersects_checks_whether_the_range_has_intersection_with_another(string leftOperand, string rightOperand, bool expected)
    {
        var left = Area.Parse(leftOperand);
        var right = Area.Parse(rightOperand);

        await Assert.That(left.Intersects(right)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("A1", "A1", true)]
    [Arguments("B1:C3", "B1:C3", true)]
    [Arguments("A1:D4", "B2:C3", true)]
    [Arguments("B3:C3", "B2:C3", false)]
    [Arguments("A2:C2", "B2:C3", false)]
    public async Task Overlaps_checks_whether_left_fully_overlaps_right(string leftOperand, string rightOperand, bool expected)
    {
        var left = Area.Parse(leftOperand);
        var right = Area.Parse(rightOperand);

        await Assert.That(left.Overlaps(right)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("C4:F8", "C1:F3", "C4:F8")] // Inserted area is fully above
    [Arguments("C4:F8", "A9:G12", "C4:F8")] // Inserted area is fully below
    [Arguments("C4:F8", "G1:H5", "C4:F8")] // Inserted are is fully to the right
    [Arguments("C4:F8", "C1:D11", "E4:H8")] // Inserted area at the left column of the area
    [Arguments("C4:F8", "A1:B8", "E4:H8")] // Inserted area is fully to the left
    [Arguments("C4:F8", "D4:E8", "C4:H8")] // Inserted into the area
    [Arguments("C4:F8", "D2:I8", "C4:L8")] // Inside the area, overlapping = extend
    [Arguments("C4:F8", "F4:F8", "C4:G8")] // Last column of the area, overlapping = extend
    [Arguments("XFD1", "XFB1", null)] // Completely pushed out of the range
    [Arguments("XFA1:XFD1", "XEZ1:XFA1", "XFC1:XFD1")] // Partially pushed out of the range
    [Arguments("XFA1:XFD1", "XFB1:XFC1", "XFA1:XFD1")] // Extend below last row
    public async Task TryInsertAreaAndShiftRight_without_partial_cover(string original, string inserted, string repositioned)
    {
        var originalArea = Area.Parse(original);
        var insertedArea = Area.Parse(inserted);
        var repositionedArea = repositioned is not null ? Area.Parse(repositioned) : (Area?)null;

        var success = originalArea.TryInsertAreaAndShiftRight(insertedArea, out var result);

        await Assert.That(success).IsTrue();
        await Assert.That(result).IsEqualTo(repositionedArea);
    }

    [Test]
    [Arguments("C4:F8", "B3:B4")] // Partially above
    [Arguments("C4:F8", "B5:C7")] // In the middle
    [Arguments("C4:F8", "A5:B9")] // Partially below
    public async Task TryInsertAreaAndShiftRight_with_partial_cover(string original, string inserted)
    {
        var originalArea = Area.Parse(original);
        var insertedArea = Area.Parse(inserted);

        await Assert.That(originalArea.TryInsertAreaAndShiftRight(insertedArea, out var result)).IsFalse();
    }

    [Test]
    [Arguments("D6:G10", "A1:C15", "D6:G10")] // Inserted are is fully to the left
    [Arguments("D6:G10", "H1:K15", "D6:G10")] // Inserted are is fully to the right
    [Arguments("D6:G10", "A11:K15", "D6:G10")] // Inserted are is fully below
    [Arguments("D6:G10", "D6:G11", "D12:G16")] // Inserted area at the top row of the area
    [Arguments("D6:G10", "C4:H7", "D10:G14")] // Inserted above the area
    [Arguments("D6:G10", "D7:G9", "D6:G13")] // Inserted into the area
    [Arguments("D6:G10", "A7:H9", "D6:G13")] // Inside the area, overlapping = extend
    [Arguments("D6:G10", "D10:G11", "D6:G12")] // Last row of the area, overlapping = extend
    [Arguments("A1048576", "A1048575", null)] // Completely pushed out of the range
    [Arguments("A1048574:A1048576", "A1048570:A1048571", "A1048576")] // Partially pushed out of the range
    [Arguments("A1048570:A1048572", "A1048571:A1048576", "A1048570:A1048576")] // Extend below last row
    public async Task TryInsertAreaAndShiftDown_without_partial_cover(string original, string inserted, string repositioned)
    {
        var originalArea = Area.Parse(original);
        var insertedArea = Area.Parse(inserted);
        var repositionedArea = repositioned is not null ? Area.Parse(repositioned) : (Area?)null;

        var success = originalArea.TryInsertAreaAndShiftDown(insertedArea, out var result);

        await Assert.That(success).IsTrue();
        await Assert.That(result).IsEqualTo(repositionedArea);
    }

    [Test]
    [Arguments("D6:G10", "A6:E6")] // Left
    [Arguments("D6:G10", "D5:D5")] // Above
    [Arguments("D6:G10", "E7:H15")] // Right
    public async Task TryInsertAreaAndShiftDown_with_partial_cover(string original, string inserted)
    {
        var originalArea = Area.Parse(original);
        var insertedArea = Area.Parse(inserted);

        await Assert.That(originalArea.TryInsertAreaAndShiftDown(insertedArea, out var result)).IsFalse();
    }

    [Test]
    [Arguments("E4:G4", "B3:C5", "C4:E4")] // Deleted area fully to the left with overlapping width
    [Arguments("E4:G4", "A2:D5", "A4:C4")] // The deleted are ends exactly at the column to the left of the area
    [Arguments("E4:G4", "F1:F7", "E4:F4")] // The deleted is fully within the area, but not at left/right column
    [Arguments("E4:G4", "E4:G4", null)] // Delete are exactly covers the area
    [Arguments("E4:G4", "A1:Z9", null)] // Delete fully covers the area
    [Arguments("E4:G4", "H1:K10", "E4:G4")] // The deleted is fully to the right of the area.
    [Arguments("E4:G4", "G3:H5", "E4:F4")] // The deleted partially intersects the area and is to the right.
    [Arguments("D4:E4", "A5:F9", "D4:E4")] // Deleted area is fully downward
    [Arguments("D4:E4", "A1:F3", "D4:E4")] // Deleted area is fully upwards
    [Arguments("D4:E4", "A5:F10", "D4:E4")] // Partial deletion is below -> not affected
    public async Task TryDeleteAreaAndShiftLeft_without_partial_cover(string original, string deleted, string repositioned)
    {
        var originalArea = Area.Parse(original);
        var deletedArea = Area.Parse(deleted);
        var repositionedArea = repositioned is not null ? Area.Parse(repositioned) : (Area?)null;

        var success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out var result);

        await Assert.That(success).IsTrue();
        await Assert.That(result).IsEqualTo(repositionedArea);
    }

    [Test]
    [Arguments("D4:E8", "A1:B5")] // Partial left
    [Arguments("D4:E8", "D2:E7")] // Partial inside
    [Arguments("D4:E8", "C4:D6")] // Partial left and inside
    public async Task TryDeleteAreaAndShiftLeft_with_partial_cover(string original, string deleted)
    {
        var originalArea = Area.Parse(original);
        var deletedArea = Area.Parse(deleted);
        var success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out var result);

        await Assert.That(success).IsFalse();
        await Assert.That(result).IsNull();
    }

    [Test]
    [Arguments("B5:B8", "A1:C3", "B2:B5")] // Deleted area fully above (with a row space) with overlapping width
    [Arguments("B5:B8", "A2:C4", "B2:B5")] // The deleted are ends exactly at the row above the area
    [Arguments("B5:B8", "A6:C7", "B5:B6")] // The deleted is fully within the area, but not at top/bottom row
    [Arguments("B5:B8", "A5:C8", null)] // Delete are exactly covers the area
    [Arguments("B5:B8", "A4:C9", null)] // Delete fully covers the area
    [Arguments("B5:B8", "A9:C10", "B5:B8")] // The deleted is fully below the area.
    [Arguments("B5:B8", "A6:C10", "B5:B5")] // The deleted partially intersects the area and is below.
    [Arguments("B5:B8", "A1:A10", "B5:B8")] // Deleted area is fully on the left
    [Arguments("B5:B8", "C1:C10", "B5:B8")] // Deleted area is fully on the right
    [Arguments("B5:D8", "B9:C10", "B5:D8")] // Partial deletion is below -> not affected
    public async Task TryDeleteAreaAndShiftUp_without_partial_cover(string leftOperand, string deleted, string expected)
    {
        var originalArea = Area.Parse(leftOperand);
        var deletedArea = Area.Parse(deleted);
        var expectedResult = expected is not null ? Area.Parse(expected) : (Area?)null;

        var success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out var result);

        await Assert.That(success).IsTrue();
        await Assert.That(result).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments("B5:D8", "A1:B3")] // Partial above
    [Arguments("B5:D8", "C6:D8")] // Partial inside
    [Arguments("B5:D8", "B1:B6")] // Partial above and inside
    public async Task TryDeleteAreaAndShiftUp_with_partial_cover(string leftOperand, string deleted)
    {
        var originalArea = Area.Parse(leftOperand);
        var deletedArea = Area.Parse(deleted);
        var success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out var result);

        await Assert.That(success).IsFalse();
        await Assert.That(result).IsNull();
    }
}
