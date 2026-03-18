using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.Styles;

[TestFixture]
public class BatchStyleTests
{
    [Test]
    public void Batch_SetMultipleProperties_AppliesAllAtOnce()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.Font.Bold = true;
            s.Font.Italic = true;
            s.Font.FontSize = 14;
            s.Fill.BackgroundColor = XLColor.Red;
            s.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        });

        Assert.IsTrue(cell.Style.Font.Bold);
        Assert.IsTrue(cell.Style.Font.Italic);
        Assert.AreEqual(14, cell.Style.Font.FontSize);
        Assert.AreEqual(XLColor.Red, cell.Style.Fill.BackgroundColor);
        Assert.AreEqual(XLAlignmentHorizontalValues.Center, cell.Style.Alignment.Horizontal);
    }

    [Test]
    public void Batch_NoChanges_DoesNotModifyStyle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");
        var originalStyle = cell.Style;

        cell.Style.Batch(s =>
        {
            // Set to same default values — no actual change
        });

        // Style should remain default
        Assert.AreEqual(XLStyle.Default.Key, ((XLStyle)cell.Style).Key);
    }

    [Test]
    public void Batch_ReturnsSameStyleInstance()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");
        var style = cell.Style;

        var result = style.Batch(s => s.Font.Bold = true);

        Assert.AreSame(style, result);
    }

    [Test]
    public void Batch_SetBorderProperties()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.Border.TopBorder = XLBorderStyleValues.Thin;
            s.Border.BottomBorder = XLBorderStyleValues.Thick;
            s.Border.LeftBorder = XLBorderStyleValues.Dashed;
            s.Border.RightBorder = XLBorderStyleValues.Double;
        });

        Assert.AreEqual(XLBorderStyleValues.Thin, cell.Style.Border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, cell.Style.Border.BottomBorder);
        Assert.AreEqual(XLBorderStyleValues.Dashed, cell.Style.Border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.Double, cell.Style.Border.RightBorder);
    }

    [Test]
    public void Batch_SetNumberFormatAndProtection()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.NumberFormat.Format = "#,##0.00";
            s.Protection.Locked = false;
            s.Protection.Hidden = true;
        });

        Assert.AreEqual("#,##0.00", cell.Style.NumberFormat.Format);
        Assert.IsFalse(cell.Style.Protection.Locked);
        Assert.IsTrue(cell.Style.Protection.Hidden);
    }

    [Test]
    public void Batch_FluentSetters_Work()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.Font.SetBold().Font.SetItalic().Font.SetFontSize(16);
        });

        Assert.IsTrue(cell.Style.Font.Bold);
        Assert.IsTrue(cell.Style.Font.Italic);
        Assert.AreEqual(16, cell.Style.Font.FontSize);
    }

    [Test]
    public void Batch_OnRange_FallsBackToNormalBehavior()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var range = ws.Range("A1:C3");

        range.Style.Batch(s =>
        {
            s.Font.Bold = true;
            s.Fill.BackgroundColor = XLColor.Blue;
        });

        // All cells in range should have the style applied
        foreach (var cell in range.Cells())
        {
            Assert.IsTrue(cell.Style.Font.Bold);
            Assert.AreEqual(XLColor.Blue, cell.Style.Fill.BackgroundColor);
        }
    }

    [Test]
    public void Batch_MatchesIndividualPropertySets()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Set via batch
        ws.Cell("A1").Style.Batch(s =>
        {
            s.Font.Bold = true;
            s.Font.FontSize = 12;
            s.Fill.BackgroundColor = XLColor.Green;
            s.Alignment.WrapText = true;
            s.Border.OutsideBorder = XLBorderStyleValues.Thin;
        });

        // Set individually
        var cell = ws.Cell("B1");
        cell.Style.Font.Bold = true;
        cell.Style.Font.FontSize = 12;
        cell.Style.Fill.BackgroundColor = XLColor.Green;
        cell.Style.Alignment.WrapText = true;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

        // Both cells should have identical style keys
        var keyA = ((XLStyle)ws.Cell("A1").Style).Key;
        var keyB = ((XLStyle)ws.Cell("B1").Style).Key;
        Assert.AreEqual(keyA, keyB);
    }

    [Test]
    public void BatchModify_WithKeyLambda_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        ((XLStyle)cell.Style).BatchModify(k => k with
        {
            Font = k.Font with { Bold = true, FontSize = 12.0 },
            Alignment = k.Alignment with { WrapText = true },
        });

        Assert.IsTrue(cell.Style.Font.Bold);
        Assert.AreEqual(12.0, cell.Style.Font.FontSize);
        Assert.IsTrue(cell.Style.Alignment.WrapText);
    }

    [Test]
    public void Batch_IncludeQuotePrefix()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.IncludeQuotePrefix = true;
            s.Font.Bold = true;
        });

        Assert.IsTrue(cell.Style.IncludeQuotePrefix);
        Assert.IsTrue(cell.Style.Font.Bold);
    }
}
