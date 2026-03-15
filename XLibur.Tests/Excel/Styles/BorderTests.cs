using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.Styles;

public class BorderTests
{
    [Test]
    public void OutsideBorder_OnDetachedStyle_SetsAllFourSides()
    {
        var style = XLWorkbook.DefaultStyle;
        style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        style.Border.OutsideBorderColor = XLColor.Black;

        Assert.AreEqual(XLBorderStyleValues.Thick, style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, style.Border.RightBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, style.Border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, style.Border.BottomBorder);
        Assert.AreEqual(XLColor.Black, style.Border.LeftBorderColor);
        Assert.AreEqual(XLColor.Black, style.Border.RightBorderColor);
        Assert.AreEqual(XLColor.Black, style.Border.TopBorderColor);
        Assert.AreEqual(XLColor.Black, style.Border.BottomBorderColor);
    }

    [Test]
    public void OutsideBorder_OnDetachedStyle_AppliedToCellWorks()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var style = XLWorkbook.DefaultStyle;
        style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        style.Border.OutsideBorderColor = XLColor.Black;

        ws.Cell("A1").Style = style;

        Assert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("A1").Style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("A1").Style.Border.RightBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("A1").Style.Border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, ws.Cell("A1").Style.Border.BottomBorder);
        Assert.AreEqual(XLColor.Black, ws.Cell("A1").Style.Border.LeftBorderColor);
        Assert.AreEqual(XLColor.Black, ws.Cell("A1").Style.Border.RightBorderColor);
        Assert.AreEqual(XLColor.Black, ws.Cell("A1").Style.Border.TopBorderColor);
        Assert.AreEqual(XLColor.Black, ws.Cell("A1").Style.Border.BottomBorderColor);
    }

    [Test]
    public void InsideBorder_OnDetachedStyle_SetsAllFourSides()
    {
        var style = XLWorkbook.DefaultStyle;
        style.Border.InsideBorder = XLBorderStyleValues.Thin;
        style.Border.InsideBorderColor = XLColor.Red;

        Assert.AreEqual(XLBorderStyleValues.Thin, style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.Thin, style.Border.RightBorder);
        Assert.AreEqual(XLBorderStyleValues.Thin, style.Border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.Thin, style.Border.BottomBorder);
        Assert.AreEqual(XLColor.Red, style.Border.LeftBorderColor);
        Assert.AreEqual(XLColor.Red, style.Border.RightBorderColor);
        Assert.AreEqual(XLColor.Red, style.Border.TopBorderColor);
        Assert.AreEqual(XLColor.Red, style.Border.BottomBorderColor);
    }

    [Test]
    public void SetInsideBorderPreservesOutsideBorders()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cells("B2:C2").Style
            .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
            .Border.SetOutsideBorderColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.5));

        // Check pre-conditions
        Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("B2").Style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("B2").Style.Border.RightBorder);
        Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor);
        Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.RightBorderColor.ThemeColor);

        ws.Range("B2:C2").Style.Border.SetInsideBorder(XLBorderStyleValues.None);

        Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("B2").Style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.None, ws.Cell("B2").Style.Border.RightBorder);
        Assert.AreEqual(XLBorderStyleValues.None, ws.Cell("C2").Style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.Thin, ws.Cell("C2").Style.Border.RightBorder);
        Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor);
        Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("C2").Style.Border.RightBorderColor.ThemeColor);
    }
}
