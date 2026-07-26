using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class BatchStyleTests
{
    [Test]
    public async Task Batch_SetMultipleProperties_AppliesAllAtOnce()
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

        await Assert.That(cell.Style.Font.Bold).IsTrue();
        await Assert.That(cell.Style.Font.Italic).IsTrue();
        await Assert.That(cell.Style.Font.FontSize).IsEqualTo(14);
        await Assert.That(cell.Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(cell.Style.Alignment.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
    }

    [Test]
    public async Task Batch_NoChanges_DoesNotModifyStyle()
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
        await Assert.That(((XLStyle)cell.Style).Key).IsEqualTo(XLStyle.Default.Key);
    }

    [Test]
    public async Task Batch_ReturnsSameStyleInstance()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");
        var style = cell.Style;

        var result = style.Batch(s => s.Font.Bold = true);

        await Assert.That(result).IsSameReferenceAs(style);
    }

    [Test]
    public async Task Batch_SetBorderProperties()
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

        await Assert.That(cell.Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thin);
        await Assert.That(cell.Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(cell.Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Dashed);
        await Assert.That(cell.Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Double);
    }

    [Test]
    public async Task Batch_SetNumberFormatAndProtection()
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

        await Assert.That(cell.Style.NumberFormat.Format).IsEqualTo("#,##0.00");
        await Assert.That(cell.Style.Protection.Locked).IsFalse();
        await Assert.That(cell.Style.Protection.Hidden).IsTrue();
    }

    [Test]
    public async Task Batch_FluentSetters_Work()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.Font.SetBold().Font.SetItalic().Font.SetFontSize(16);
        });

        await Assert.That(cell.Style.Font.Bold).IsTrue();
        await Assert.That(cell.Style.Font.Italic).IsTrue();
        await Assert.That(cell.Style.Font.FontSize).IsEqualTo(16);
    }

    [Test]
    public async Task Batch_OnRange_FallsBackToNormalBehavior()
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
            await Assert.That(cell.Style.Font.Bold).IsTrue();
            await Assert.That(cell.Style.Fill.BackgroundColor).IsEqualTo(XLColor.Blue);
        }
    }

    [Test]
    public async Task Batch_MatchesIndividualPropertySets()
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
        await Assert.That(keyB).IsEqualTo(keyA);
    }

    [Test]
    public async Task BatchModify_WithKeyLambda_Works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        ((XLStyle)cell.Style).BatchModify(k => k with
        {
            Font = k.Font with { Bold = true, FontSize = 12.0 },
            Alignment = k.Alignment with { WrapText = true },
        });

        await Assert.That(cell.Style.Font.Bold).IsTrue();
        await Assert.That(cell.Style.Font.FontSize).IsEqualTo(12.0);
        await Assert.That(cell.Style.Alignment.WrapText).IsTrue();
    }

    [Test]
    public async Task Batch_IncludeQuotePrefix()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Batch(s =>
        {
            s.IncludeQuotePrefix = true;
            s.Font.Bold = true;
        });

        await Assert.That(cell.Style.IncludeQuotePrefix).IsTrue();
        await Assert.That(cell.Style.Font.Bold).IsTrue();
    }
}
