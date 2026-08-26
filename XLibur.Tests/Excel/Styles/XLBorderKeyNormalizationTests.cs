using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Utils;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// An edge with no border style has no colour to draw with, and Excel writes none, so the colour
/// held against such an edge is dropped whenever a border key reaches a repository or an
/// <c>XLStyleKey</c>.
/// </summary>
/// <remarks>
/// The colour was never honoured at that level - equality already treated two styleless edges as
/// the same edge whatever their colours - but the key that won the intern, and so the colour a read
/// reported, was whichever key reached the repository first. These tests pin the collapsed form
/// down.
/// <para>
/// A colour assigned to a still-styleless edge is not simply lost, though: <c>XLBorder</c> holds it
/// as pending and applies it if the edge is later given a style on the same facade - see
/// <see cref="Colour_assigned_before_the_style_is_kept_once_a_style_follows"/> and its siblings.
/// </para>
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
    /// Normalization must leave alone the colour of an edge that has a style.
    /// </summary>
    [Test]
    public async Task Colour_assigned_after_the_style_is_kept()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;
        cell.Style.Border.TopBorderColor = XLColor.Red;

        await Assert.That(cell.Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// Assigning the colour before the style used to lose it: the colour landed on a styleless
    /// edge, was collapsed to black there, and the style that followed had nothing left to combine
    /// it with. <c>XLBorder</c> now holds such a colour as pending and applies it once the edge is
    /// given a style, so the two assignments give the same result in either order.
    /// </summary>
    /// <remarks>
    /// Exercised through <c>cell.Style.Border</c> twice, as the property would ordinarily be
    /// written, rather than through one held reference - see
    /// <see cref="Colour_assigned_before_the_style_survives_two_separate_dot_Border_accesses"/> for
    /// why that distinction matters here specifically.
    /// </remarks>
    [Test]
    public async Task Colour_assigned_before_the_style_is_kept_once_a_style_follows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A2");

        var border = cell.Style.Border;
        border.TopBorderColor = XLColor.Red;
        border.TopBorder = XLBorderStyleValues.Thick;

        await Assert.That(cell.Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(cell.Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// The pending colour is held on the <c>XLBorder</c> facade, and <c>cell.Style</c> and
    /// <c>Style.Border</c> both cache their facade per cell, returning the same instance on repeat
    /// access rather than a fresh one. Without that, the ordinary two-statement idiom below - each
    /// statement fetching its own facade rather than sharing a held reference - would lose the
    /// colour exactly as it did before the pending mechanism existed.
    /// </summary>
    [Test]
    public async Task Colour_assigned_before_the_style_survives_two_separate_dot_Border_accesses()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A2");

        cell.Style.Border.TopBorderColor = XLColor.Red;
        cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;

        await Assert.That(cell.Style.Border.TopBorderColor).IsEqualTo(XLColor.Red);
    }

    /// <summary>
    /// A pending colour is a one-shot intent, consumed by the first real style transition it sees -
    /// whether that transition applies it (edge goes from <c>None</c> to a style) or makes it moot
    /// (edge is explicitly cleared back to <c>None</c> before ever being given a style). Once
    /// consumed it does not reappear on some later, unrelated transition.
    /// </summary>
    [Test]
    public async Task Colour_pending_on_an_edge_does_not_survive_an_explicit_transition_through_None()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A2");

        var border = cell.Style.Border;
        border.TopBorderColor = XLColor.Red; // TopBorder is still None: held as pending.
        border.TopBorder = XLBorderStyleValues.Thin; // None -> Thin: pending Red is applied and consumed.
        border.TopBorder = XLBorderStyleValues.None; // Thin -> None: Red is normalized away, same as any styled-to-None edge.
        border.TopBorder = XLBorderStyleValues.Thick; // None -> Thick again: nothing pending this time.

        await Assert.That(cell.Style.Border.TopBorderColor).IsEqualTo(XLColor.Black);
    }

    /// <summary>
    /// A colour held pending against one ground truth must not be applied to a different one: if
    /// the border changes for a reason other than this facade's own writes - here, replacing the
    /// whole style - a stale pending colour is discarded rather than resurrected on the next style
    /// change.
    /// </summary>
    [Test]
    public async Task Colour_pending_on_an_edge_is_discarded_if_the_border_changes_from_elsewhere()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A2");

        // Gives the cell a non-default border, so replacing the whole style below is a real,
        // detectable change rather than a same-key no-op.
        cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

        var border = cell.Style.Border;
        border.TopBorderColor = XLColor.Red; // TopBorder is still None: held as pending.

        // Replaces the whole style, including the border, out from under the held facade.
        cell.Style = XLWorkbook.DefaultStyle;

        cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;

        await Assert.That(cell.Style.Border.TopBorderColor).IsEqualTo(XLColor.Black);
    }

    /// <summary>
    /// The compound edge setters apply all four non-diagonal edges as one write and do not use the
    /// pending mechanism - see the remarks on the pending-colour fields in <c>XLBorder</c>. A colour
    /// assigned this way before the matching style is still collapsed to black, same as it was
    /// before the pending mechanism existed for the single-edge properties.
    /// </summary>
    [Test]
    public async Task Compound_edge_colour_assigned_before_the_compound_style_is_dropped()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A2");

        cell.Style.Border.OutsideBorderColor = XLColor.Red;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

        await Assert.That(cell.Style.Border.TopBorderColor).IsEqualTo(XLColor.Black);
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

        var converted = StyleDecoder.BorderKey(border, XLBorderValue.Default.Key);

        await Assert.That(converted.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(converted).IsEqualTo(XLBorderValue.Default.Key);
    }
}
