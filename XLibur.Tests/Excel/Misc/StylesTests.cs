using System;
using System.Collections.Generic;
using XLibur.Excel;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class StylesTests
{
    private static void SetupBorders(IXLRange range)
    {
        range.FirstRow().Cell(1).Style.Border.TopBorder = XLBorderStyleValues.None;
        range.FirstRow().Cell(2).Style.Border.TopBorder = XLBorderStyleValues.Thick;
        range.FirstRow().Cell(3).Style.Border.TopBorder = XLBorderStyleValues.Double;

        range.LastRow().Cell(1).Style.Border.BottomBorder = XLBorderStyleValues.None;
        range.LastRow().Cell(2).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        range.LastRow().Cell(3).Style.Border.BottomBorder = XLBorderStyleValues.Double;

        range.FirstColumn().Cell(1).Style.Border.LeftBorder = XLBorderStyleValues.None;
        range.FirstColumn().Cell(2).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
        range.FirstColumn().Cell(3).Style.Border.LeftBorder = XLBorderStyleValues.Double;

        range.LastColumn().Cell(1).Style.Border.RightBorder = XLBorderStyleValues.None;
        range.LastColumn().Cell(2).Style.Border.RightBorder = XLBorderStyleValues.Thick;
        range.LastColumn().Cell(3).Style.Border.RightBorder = XLBorderStyleValues.Double;
    }

    [Test]
    public async Task InsideBorderTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var range = ws.Range("B2:D4");

        SetupBorders(range);

        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorderColor = XLColor.Red;

        var center = range.Cell(2, 2);

        await Assert.That(center.Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(center.Style.Border.BottomBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(center.Style.Border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(center.Style.Border.RightBorderColor).IsEqualTo(XLColor.Red);

        await Assert.That(range.FirstRow().Cell(1).Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(range.FirstRow().Cell(2).Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(range.FirstRow().Cell(3).Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Double);

        await Assert.That(range.LastRow().Cell(1).Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(range.LastRow().Cell(2).Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(range.LastRow().Cell(3).Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Double);

        await Assert.That(range.FirstColumn().Cell(1).Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(range.FirstColumn().Cell(2).Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(range.FirstColumn().Cell(3).Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Double);

        await Assert.That(range.LastColumn().Cell(1).Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(range.LastColumn().Cell(2).Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(range.LastColumn().Cell(3).Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Double);
    }

    [Test]
    public async Task ResolveThemeColors()
    {
        using var wb = new XLWorkbook();
        var color = wb.Theme.ResolveThemeColor(XLThemeColor.Accent1).Color.ToHex();
        await Assert.That(color).IsEqualTo("FF4F81BD");

        color = wb.Theme.ResolveThemeColor(XLThemeColor.Background1).Color.ToHex();
        await Assert.That(color).IsEqualTo("FFFFFFFF");
    }

    // NUnit's [Theory] fed this from the enum automatically; TUnit needs it explicit.
    public static IEnumerable<XLThemeColor> AllThemeColors() => Enum.GetValues<XLThemeColor>();

    [Test]
    [MethodDataSource(nameof(AllThemeColors))]
    public async Task CanResolveAllThemeColors(XLThemeColor themeColor)
    {
        var theme = new XLWorkbook().Theme;
        var color = theme.ResolveThemeColor(themeColor);
        await Assert.That(color).IsNotNull();
    }

    [Test]
    public async Task SetStyleViaRowReference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Style
            .Font.SetFontSize(8)
            .Font.SetFontColor(XLColor.Green)
            .Font.SetBold(true);

        var row = ws.Row(1);
        ws.Cell(1, 1).Value = "Test";
        row.Cell(2).Value = "Test";
        row.Cells(3, 3).Value = "Test";

        foreach (var cell in ws.CellsUsed())
        {
            await Assert.That(ws.Cell("A1").Style.Font.FontSize).IsEqualTo(8);
            await Assert.That(ws.Cell("B1").Style.Font.FontColor).IsEqualTo(XLColor.Green);
            await Assert.That(ws.Cell("C1").Style.Font.Bold).IsTrue();
        }
    }
}
