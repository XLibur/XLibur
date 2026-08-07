using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Utils;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// An edge with no border style has no colour to draw with, and Excel writes none, so the colour
/// stored against such an edge is dropped when a border key is interned.
/// </summary>
/// <remarks>
/// The colour was never honoured - equality already treated two styleless edges as the same edge
/// whatever their colours - but the key that won the intern, and so the colour a read reported, was
/// whichever key reached the repository first. These tests pin the collapsed form down.
/// </remarks>
public class XLBorderKeyNormalizationTests
{
    [Test]
    public async Task Colour_set_on_a_styleless_edge_reads_back_as_black()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        await Assert.That(cell.Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);

        cell.Style.Border.LeftBorderColor = XLColor.Red;

        await Assert.That(cell.Style.Border.LeftBorderColor).IsEqualTo(XLColor.Black);
    }

    [Test]
    public async Task Two_cells_that_differ_only_in_a_styleless_edge_colour_share_one_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var a1 = ws.Cell("A1");
        a1.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        a1.Style.Border.LeftBorderColor = XLColor.Red;
        a1.Style.Border.LeftBorder = XLBorderStyleValues.None;

        var a2 = ws.Cell("A2");
        a2.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        a2.Style.Border.LeftBorderColor = XLColor.Blue;
        a2.Style.Border.LeftBorder = XLBorderStyleValues.None;

        await Assert.That(a1.Style).IsEqualTo(a2.Style);
        await Assert.That(a1.Style.Border.LeftBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(a2.Style.Border.LeftBorderColor).IsEqualTo(XLColor.Black);
    }

    /// <summary>
    /// Normalization must not touch an edge that has a style, including when the colour is assigned
    /// before the style.
    /// </summary>
    [Test]
    public async Task Colour_on_a_styled_edge_survives_whichever_order_it_is_assigned_in()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var styleFirst = ws.Cell("A1");
        styleFirst.Style.Border.TopBorder = XLBorderStyleValues.Thick;
        styleFirst.Style.Border.TopBorderColor = XLColor.Red;

        var colourFirst = ws.Cell("A2");
        colourFirst.Style.Border.TopBorderColor = XLColor.Red;
        colourFirst.Style.Border.TopBorder = XLBorderStyleValues.Thick;

        await Assert.That(styleFirst.Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);

        // Assigning the colour first sets it on a styleless edge, where it is dropped; the style
        // that follows does not bring it back. Documented rather than asserted as desirable.
        await Assert.That(colourFirst.Style.Border.TopBorderColor).IsEqualTo(XLColor.Black);
    }

    [Test]
    public async Task Every_edge_is_normalized_independently()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.LeftBorderColor = XLColor.Red;
        cell.Style.Border.RightBorderColor = XLColor.Blue;
        cell.Style.Border.TopBorder = XLBorderStyleValues.Double;
        cell.Style.Border.TopBorderColor = XLColor.Green;
        cell.Style.Border.BottomBorderColor = XLColor.Yellow;
        cell.Style.Border.DiagonalBorderColor = XLColor.Orange;

        var border = cell.Style.Border;
        await Assert.That(border.LeftBorderColor).IsEqualTo(XLColor.Red);
        await Assert.That(border.TopBorderColor).IsEqualTo(XLColor.Green);
        await Assert.That(border.RightBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(border.BottomBorderColor).IsEqualTo(XLColor.Black);
        await Assert.That(border.DiagonalBorderColor).IsEqualTo(XLColor.Black);
    }

    /// <summary>
    /// A side's style and colour are independent attributes in the schema, so a file may state a
    /// colour for an edge it gives no style. Such a border must still be recognised as the border it
    /// is equivalent to, or the styles part gains a duplicate <c>&lt;border&gt;</c> on every save.
    /// </summary>
    [Test]
    public async Task Border_read_with_a_colour_but_no_style_matches_the_plain_default()
    {
        var border = new Border
        {
            LeftBorder = new LeftBorder(new Color { Rgb = "FFFF0000" }),
        };

        var converted = OpenXmlHelper.BorderToXLibur(border, XLBorderValue.Default.Key);

        await Assert.That(converted.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(converted).IsEqualTo(XLBorderValue.Default.Key);
    }
}
