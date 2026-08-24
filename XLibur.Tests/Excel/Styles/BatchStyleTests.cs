using System;
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

    /// <summary>
    /// A component facade the caller obtained before the batch must report what the batch set. The
    /// style's own getters resync on every access, so only a retained facade can go stale.
    /// </summary>
    [Test]
    public async Task Facade_retained_across_a_batch_reports_the_batched_value()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        var font = cell.Style.Font;
        cell.Style.Batch(s => s.Font.Bold = true);

        await Assert.That(font.Bold).IsTrue();
    }

    /// <summary>
    /// The damaging half of the same staleness: a write through a retained facade rebuilds the
    /// component key from whatever that facade holds, so a stale one silently drops everything the
    /// batch had set.
    /// </summary>
    [Test]
    public async Task Writing_through_a_facade_retained_across_a_batch_keeps_the_batched_value()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        var font = cell.Style.Font;
        cell.Style.Batch(s => s.Font.FontSize = 20);
        font.Italic = true;

        await Assert.That(cell.Style.Font.FontSize).IsEqualTo(20);
        await Assert.That(cell.Style.Font.Italic).IsTrue();
    }

    /// <summary>
    /// Refreshing the border facade after a batch must not drop a colour left pending by that same
    /// batch. A colour assigned to a styleless edge is held until the edge is given a style, and
    /// the direct path keeps it across exactly this transition - so the batch path must too.
    /// </summary>
    [Test]
    public async Task A_colour_left_pending_by_a_batch_still_applies_afterwards()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var batched = ws.Cell("A1");
        batched.Style.Batch(s =>
        {
            s.Border.LeftBorder = XLBorderStyleValues.Thin;
            s.Border.TopBorderColor = XLColor.Red;
        });
        batched.Style.Border.TopBorder = XLBorderStyleValues.Thin;

        var direct = ws.Cell("B1");
        direct.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        direct.Style.Border.TopBorderColor = XLColor.Red;
        direct.Style.Border.TopBorder = XLBorderStyleValues.Thin;

        await Assert.That(BorderSignature(batched)).IsEqualTo(BorderSignature(direct));
    }

    /// <summary>
    /// Every <see cref="IXLBorder"/> property must reach the same cell style whether it is assigned
    /// directly or inside a <see cref="IXLStyle.Batch"/>. Before spec 23 these ran through two
    /// independent implementations of <c>IXLBorder</c> - <c>XLBorder</c> and <c>XLDeferredBorder</c> -
    /// and <c>InsideBorder</c>/<c>InsideBorderColor</c> disagreed: a single cell has no interior
    /// edges, so the direct path is a no-op, while the deferred path set all four.
    /// </summary>
    [Test]
    [Arguments("OutsideBorder")]
    [Arguments("InsideBorder")]
    [Arguments("LeftBorder")]
    [Arguments("RightBorder")]
    [Arguments("TopBorder")]
    [Arguments("BottomBorder")]
    [Arguments("DiagonalBorder")]
    public async Task Batch_and_direct_assignment_agree_for_every_border_property(string property)
    {
        const XLBorderStyleValues value = XLBorderStyleValues.Thick;

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var direct = ws.Cell("A1");
        ApplyBorder(direct.Style.Border, property, value);

        var batched = ws.Cell("B1");
        batched.Style.Batch(s => ApplyBorder(s.Border, property, value));

        await Assert.That(BorderSignature(batched)).IsEqualTo(BorderSignature(direct));
    }

    /// <summary>Colour variants of the same parity property.</summary>
    [Test]
    [Arguments("OutsideBorderColor")]
    [Arguments("InsideBorderColor")]
    [Arguments("LeftBorderColor")]
    [Arguments("RightBorderColor")]
    [Arguments("TopBorderColor")]
    [Arguments("BottomBorderColor")]
    [Arguments("DiagonalBorderColor")]
    public async Task Batch_and_direct_assignment_agree_for_every_border_colour(string property)
    {
        var value = XLColor.Red;

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var direct = ws.Cell("A1");
        ApplyBorderColor(direct.Style.Border, property, value);

        var batched = ws.Cell("B1");
        batched.Style.Batch(s => ApplyBorderColor(s.Border, property, value));

        await Assert.That(BorderSignature(batched)).IsEqualTo(BorderSignature(direct));
    }

    private static void ApplyBorder(IXLBorder border, string property, XLBorderStyleValues value)
    {
        switch (property)
        {
            case "OutsideBorder": border.OutsideBorder = value; break;
            case "InsideBorder": border.InsideBorder = value; break;
            case "LeftBorder": border.LeftBorder = value; break;
            case "RightBorder": border.RightBorder = value; break;
            case "TopBorder": border.TopBorder = value; break;
            case "BottomBorder": border.BottomBorder = value; break;
            case "DiagonalBorder": border.DiagonalBorder = value; break;
            default: throw new ArgumentOutOfRangeException(nameof(property), property, null);
        }
    }

    private static void ApplyBorderColor(IXLBorder border, string property, XLColor value)
    {
        switch (property)
        {
            case "OutsideBorderColor": border.OutsideBorderColor = value; break;
            case "InsideBorderColor": border.InsideBorderColor = value; break;
            case "LeftBorderColor": border.LeftBorderColor = value; break;
            case "RightBorderColor": border.RightBorderColor = value; break;
            case "TopBorderColor": border.TopBorderColor = value; break;
            case "BottomBorderColor": border.BottomBorderColor = value; break;
            case "DiagonalBorderColor": border.DiagonalBorderColor = value; break;
            default: throw new ArgumentOutOfRangeException(nameof(property), property, null);
        }
    }

    /// <summary>
    /// Reads the cell's border straight off the style key, so the comparison cannot be satisfied by
    /// two facades that merely agree with themselves.
    /// </summary>
    private static string BorderSignature(IXLCell cell)
    {
        var k = ((XLStyle)cell.Style).Key.Border;
        return string.Join('|',
            k.LeftBorder, k.RightBorder, k.TopBorder, k.BottomBorder, k.DiagonalBorder,
            k.LeftBorderColor, k.RightBorderColor, k.TopBorderColor, k.BottomBorderColor,
            k.DiagonalBorderColor, k.DiagonalUp, k.DiagonalDown);
    }
}
