using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class BorderTests
{
    [Test]
    public async Task OutsideBorder_OnDetachedStyle_SetsAllFourSides()
    {
        var style = XLWorkbook.DefaultStyle;
        style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        style.Border.OutsideBorderColor = XLColor.Black;

        await Assert.That(style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(style.Border.LeftBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(style.Border.RightBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(style.Border.TopBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(style.Border.BottomBorderColor).IsEqualTo(XLColor.Black);
    }

    [Test]
    public async Task OutsideBorder_OnDetachedStyle_AppliedToCellWorks()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var style = XLWorkbook.DefaultStyle;
        style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell("A1").Style = style;

        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("A1").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("A1").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("A1").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(ws.Cell("A1").Style.Border.RightBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(ws.Cell("A1").Style.Border.TopBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(ws.Cell("A1").Style.Border.BottomBorderColor).IsEqualTo(XLColor.Black);
    }

    [Test]
    public async Task InsideBorder_OnDetachedStyle_SetsAllFourSides()
    {
        var style = XLWorkbook.DefaultStyle;
        style.Border.InsideBorder = XLBorderStyleValues.Thin;
        style.Border.InsideBorderColor = XLColor.Red;

        await Assert.That(style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(style.Border.RightBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(style.Border.BottomBorderColor).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task SetInsideBorderPreservesOutsideBorders()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cells("B2:C2").Style
            .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
            .Border.SetOutsideBorderColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.5));

        // Check pre-conditions
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("B2").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(ws.Cell("B2").Style.Border.RightBorderColor.ThemeColor).IsEqualTo(XLThemeColor.Accent1);

        ws.Range("B2:C2").Style.Border.SetInsideBorder(XLBorderStyleValues.None);

        await Assert.That(ws.Cell("B2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("B2").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C2").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(ws.Cell("C2").Style.Border.RightBorderColor.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
    }
}
