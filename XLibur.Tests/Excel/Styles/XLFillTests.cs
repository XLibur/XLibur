using System.IO;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class XLFillTests
{
    [Test]
    public async Task BackgroundColorSetsPattern()
    {
        var fill = new XLFill { BackgroundColor = XLColor.Blue };
        await Assert.That(fill.PatternType).IsEqualTo(XLFillPatternValues.Solid);
    }

    [Test]
    public async Task BackgroundNoColorSetsPatternNone()
    {
        var fill = new XLFill { BackgroundColor = XLColor.Automatic };
        await Assert.That(fill.PatternType).IsEqualTo(XLFillPatternValues.None);
    }

    [Test]
    public async Task BackgroundPatternEqualCheck()
    {
        var fill1 = new XLFill { BackgroundColor = XLColor.Blue };
        var fill2 = new XLFill { BackgroundColor = XLColor.Blue };
        await Assert.That(fill1.Equals(fill2)).IsTrue();
        await Assert.That(fill2.GetHashCode()).IsEqualTo(fill1.GetHashCode());
    }

    [Test]
    public async Task BackgroundPatternNotEqualCheck()
    {
        var fill1 = new XLFill { PatternType = XLFillPatternValues.Solid, BackgroundColor = XLColor.Blue };
        var fill2 = new XLFill { PatternType = XLFillPatternValues.Solid, BackgroundColor = XLColor.Red };
        await Assert.That(fill1.Equals(fill2)).IsFalse();
        await Assert.That(fill2.GetHashCode()).IsNotEqualTo(fill1.GetHashCode());
    }

    [Test]
    public async Task FillsWithTransparentColorEqual()
    {
        var fill1 = new XLFill { BackgroundColor = XLColor.ElectricUltramarine, PatternType = XLFillPatternValues.None };
        var fill2 = new XLFill { BackgroundColor = XLColor.EtonBlue, PatternType = XLFillPatternValues.None };
        var fill3 = new XLFill { BackgroundColor = XLColor.FromIndex(64) };
        var fill4 = new XLFill { BackgroundColor = XLColor.Automatic };

        await Assert.That(fill1.Equals(fill2)).IsTrue();
        await Assert.That(fill1.Equals(fill3)).IsTrue();
        await Assert.That(fill1.Equals(fill4)).IsTrue();
        await Assert.That(fill2.GetHashCode()).IsEqualTo(fill1.GetHashCode());
        await Assert.That(fill3.GetHashCode()).IsEqualTo(fill1.GetHashCode());
        await Assert.That(fill4.GetHashCode()).IsEqualTo(fill1.GetHashCode());
    }

    [Test]
    public async Task SolidFillsWithDifferentPatternColorEqual()
    {
        var fill1 = new XLFill
        {
            PatternType = XLFillPatternValues.Solid,
            BackgroundColor = XLColor.Red,
            PatternColor = XLColor.Blue
        };

        var fill2 = new XLFill
        {
            PatternType = XLFillPatternValues.Solid,
            BackgroundColor = XLColor.Red,
            PatternColor = XLColor.Green
        };

        await Assert.That(fill1.Equals(fill2)).IsTrue();
        await Assert.That(fill2.GetHashCode()).IsEqualTo(fill1.GetHashCode());
    }

    [Test]
    public async Task BackgroundWithConditionalFormat()
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Test");
        worksheet.Cell(2, 2).SetValue("Text");
        var cf = worksheet.Cell(2, 2).AddConditionalFormat();
        var style = cf.WhenNotBlank();
        style = style
            .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XLColor.Blue);

        await Assert.That(style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);

        await Assert.That(style.Border.BottomBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(style.Border.TopBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(style.Border.LeftBorderColor).IsEqualTo(XLColor.Blue);
        await Assert.That(style.Border.RightBorderColor).IsEqualTo(XLColor.Blue);
    }

    [Test]
    public async Task LoadAndSaveTransparentBackgroundFill()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\TransparentBackgroundFill\inputfile.xlsx"));
        using var ms = new MemoryStream();
        await TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);
            wb.SaveAs(ms);
            return wb;
        }, @"Other\StyleReferenceFiles\TransparentBackgroundFill\TransparentBackgroundFill.xlsx");
    }

    [Test]
    public async Task ReservedFills_ReplaceWithPredefinedValues()
    {
        // If attribute or whole predefined fill is missing from the file, save predefined values
        await TestHelper.LoadAndAssert(async wb =>
        {
            await Assert.That(wb.Worksheets.Count).IsGreaterThan(0);
        }, @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Input.xlsx");

        await TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Input.xlsx",
            @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Output.xlsx");
    }

    [Test]
    public async Task ReservedFills_MoveFillsFromReservedPositions()
    {
        // If the input doesn't have expected fill values at the reserved position s0 and 1 (can only happen
        // for non-excel sources, excel always has correct values), put expected fill at 0 and 1, but save original
        // fills to different positions if they are used.
        await TestHelper.LoadAndAssert(async wb =>
        {
            await Assert.That(wb.Worksheets.Count).IsGreaterThan(0);
        }, @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Input.xlsx");

        await TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Input.xlsx",
            @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Output.xlsx");
    }
}
