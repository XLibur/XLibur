using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class StyleChangeTests
{
    [Test]
    public async Task ChangeFontColorDoesNotAffectOtherProperties()
    {
        using var wb = new XLWorkbook();
        // Arrange
        var ws = wb.AddWorksheet("Sheet1");
        var a1 = ws.Cell("A1");
        var a2 = ws.Cell("A2");
        var b1 = ws.Cell("B1");
        var b2 = ws.Cell("B2");

        ws.Range("A1:B2").Value = "Test";

        a1.Style.Fill.BackgroundColor = XLColor.Red;
        a2.Style.Fill.BackgroundColor = XLColor.Green;
        b1.Style.Fill.BackgroundColor = XLColor.Blue;
        b2.Style.Fill.BackgroundColor = XLColor.Pink;

        a1.Style.Font.FontName = "Arial";
        a2.Style.Font.FontName = "Times New Roman";
        b1.Style.Font.FontName = "Calibri";
        b2.Style.Font.FontName = "Cambria";

        // Act
        ws.Range("A1:B2").Style.Font.FontColor = XLColor.PowderBlue;

        // Assert
        await Assert.That(ws.Cell("A1").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("A2").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Cell("B1").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Blue);
        await Assert.That(ws.Cell("B2").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Pink);

        await Assert.That(ws.Cell("A1").Style.Font.FontName).IsEqualTo("Arial");
        await Assert.That(ws.Cell("A2").Style.Font.FontName).IsEqualTo("Times New Roman");
        await Assert.That(ws.Cell("B1").Style.Font.FontName).IsEqualTo("Calibri");
        await Assert.That(ws.Cell("B2").Style.Font.FontName).IsEqualTo("Cambria");

        await Assert.That(ws.Cell("A1").Style.Font.FontColor).IsEqualTo(XLColor.PowderBlue);
        await Assert.That(ws.Cell("A2").Style.Font.FontColor).IsEqualTo(XLColor.PowderBlue);
        await Assert.That(ws.Cell("B1").Style.Font.FontColor).IsEqualTo(XLColor.PowderBlue);
        await Assert.That(ws.Cell("B2").Style.Font.FontColor).IsEqualTo(XLColor.PowderBlue);
    }

    [Test]
    public async Task ChangeDetachedStyleAlignment()
    {
        var style = XLStyle.Default;

        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

        await Assert.That(style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Justify);
    }

    [Test]
    public async Task ChangeDetachedStyleBorder()
    {
        var style = XLStyle.Default;

        style.Border.DiagonalBorder = XLBorderStyleValues.Double;

        await Assert.That(style.Border.DiagonalBorder).IsEqualTo(XLBorderStyleValues.Double);
    }

    [Test]
    public async Task ChangeDetachedStyleFill()
    {
        var style = XLStyle.Default;

        style.Fill.BackgroundColor = XLColor.Red;

        await Assert.That(style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task ChangeDetachedStyleFont()
    {
        var style = XLStyle.Default;

        style.Font.FontSize = 50;

        await Assert.That(style.Font.FontSize).IsEqualTo(50);
    }

    [Test]
    public async Task ChangeDetachedStyleNumberFormat()
    {
        var style = XLStyle.Default;

        style.NumberFormat.Format = "YYYY";

        await Assert.That(style.NumberFormat.Format).IsEqualTo("YYYY");
    }

    [Test]
    public async Task ChangeDetachedStyleProtection()
    {
        var style = XLStyle.Default;

        style.Protection.Hidden = true;

        await Assert.That(style.Protection.Hidden).IsTrue();
    }

    [Test]
    public async Task ChangeAttachedStyleAlignment()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var a1 = ws.Cell("A1");

        a1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;

        await Assert.That(a1.Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Justify);
    }
}
