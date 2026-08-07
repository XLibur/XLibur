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

    /// <summary>
    /// A border facade holds the interned border it last resolved, and takes it off the style rather
    /// than interning the new key itself. These hold onto one facade across the write and the read,
    /// so a facade left pointing at the pre-change value fails rather than being hidden by the next
    /// <c>.Style.Border</c> building a fresh one.
    /// </summary>
    [Test]
    public async Task Border_facade_on_a_range_reflects_the_change_it_applied()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var border = ws.Range("A1:C3").Style.Border;
        border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("A1").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(ws.Cell("C3").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
    }

    [Test]
    public async Task Border_facade_on_a_worksheet_reflects_a_compound_change_it_applied()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var border = ws.Style.Border;
        border.OutsideBorder = XLBorderStyleValues.Thick;

        await Assert.That(border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
    }

    /// <summary>
    /// A cell is a cell container, so a compound setter reaches the fast path that resolves the
    /// style and reads the interned component back off it.
    /// </summary>
    [Test]
    public async Task Border_facade_on_a_cell_reflects_a_compound_change_it_applied()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var border = ws.Cell("E5").Style.Border;
        border.OutsideBorder = XLBorderStyleValues.Double;
        border.OutsideBorderColor = XLColor.Red;

        await Assert.That(border.TopBorder).IsEqualTo(XLBorderStyleValues.Double);
        await Assert.That(border.BottomBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("E5").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Double);
        await Assert.That(ws.Cell("E5").Style.Border.RightBorderColor).IsEqualTo(XLColor.Red);
    }
}
