using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// A style facade built without a style, over a key of its own, keeps that key when a property is
/// set through it.
/// </summary>
/// <remarks>
/// Such a facade gets an empty style that holds the default component, not the one it was built
/// over. Its <c>Modify</c> must therefore rebuild from its own key, not take the value back off
/// the style as an attached facade does (#621) - otherwise setting one property resets every
/// other to its default. A rich text run's font is the case that exists today; the other four
/// facades accept a null style in their constructors too.
/// </remarks>
public class DetachedStyleFacadeTests
{
    [Test]
    public async Task Font_KeepsItsOwnKey()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet().Cell(1, 1).Style.Font;
        source.FontName = "Arial";

        var font = new XLFont(source);
        font.Bold = true;

        await Assert.That(font.FontName).IsEqualTo("Arial");
        await Assert.That(font.Bold).IsTrue();
    }

    [Test]
    public async Task Fill_KeepsItsOwnKey()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet().Cell(1, 1).Style.Fill;
        // Not solid: a solid fill's key ignores its pattern colour, so the write would not show.
        source.PatternType = XLFillPatternValues.DarkGray;

        var fill = new XLFill(null, source);
        fill.PatternColor = XLColor.Blue;

        await Assert.That(fill.PatternType).IsEqualTo(XLFillPatternValues.DarkGray);
        await Assert.That(fill.PatternColor).IsEqualTo(XLColor.Blue);
    }

    [Test]
    public async Task Alignment_KeepsItsOwnKey()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet().Cell(1, 1).Style.Alignment;
        source.WrapText = true;

        var alignment = new XLAlignment(null, source);
        alignment.Vertical = XLAlignmentVerticalValues.Top;

        await Assert.That(alignment.WrapText).IsTrue();
        await Assert.That(alignment.Vertical).IsEqualTo(XLAlignmentVerticalValues.Top);
    }

    [Test]
    public async Task Protection_KeepsItsOwnKey()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet().Cell(1, 1).Style.Protection;
        source.Locked = false;

        var protection = new XLProtection(null, source);
        protection.Hidden = true;

        await Assert.That(protection.Locked).IsFalse();
        await Assert.That(protection.Hidden).IsTrue();
    }

    /// <summary>
    /// Every number format setter replaces the whole key, so there is nothing of its own to keep;
    /// this pins only that the detached path writes the value.
    /// </summary>
    [Test]
    public async Task NumberFormat_TakesTheValue()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet().Cell(1, 1).Style.NumberFormat;
        source.Format = "0.00";

        var numberFormat = new XLNumberFormat(null, source);
        numberFormat.NumberFormatId = 14;

        await Assert.That(numberFormat.NumberFormatId).IsEqualTo(14);
        await Assert.That(numberFormat.Format).IsEqualTo(string.Empty);
    }
}
