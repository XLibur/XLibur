using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// A picture loaded from a two-cell anchor reports the size of its <c>from</c>/<c>to</c> box, and
/// an <c>editAs="absolute"</c> one the position of its <c>from</c> marker, not its <c>a:xfrm</c>
/// (issue #782).
/// </summary>
/// <remarks>
/// <para>
/// Every fixture is an 8×8 px picture whose <c>a:xfrm</c> holds offset 0,0 and the native size, as
/// XLibur writes it, so the expected values below can only come from the anchor.
/// </para>
/// <para>
/// The expected values were measured in Excel (Microsoft 365, over COM) from the same files: the
/// shape's <c>Left</c>, <c>Top</c>, <c>Width</c> and <c>Height</c>. COM reports points. The Excel
/// that measured them runs at 120 DPI, where it lays the grid out in whole pixels of that DPI, so
/// these tests load at 120 DPI and the points are converted at 120 / 72: Excel's 180 pt is 300 px.
/// The fixtures give every row and column an explicit height or width, or a custom sheet default,
/// because Excel derives an unstated default from the font at its own DPI, which XLibur does not.
/// </para>
/// </remarks>
public class PictureAnchorSizeTests
{
    private const string Drawing1 = "xl/drawings/drawing1.xml";

    private static readonly LoadOptions AtExcelDpi = new() { Dpi = new Point(120, 120) };

    private static readonly string[] EditAsValues = ["oneCell", "absolute", "twoCell"];

    [Test]
    // Columns B 114 px and C 186 px; rows 8, 10 and 11 30 px and row 9 50 px.
    [Arguments("grid", 300, 140)]
    // Column C and row 10 are hidden, and add nothing.
    [Arguments("hidden", 114, 110)]
    // From B8 + (10, 5) px to D12 + (20, 10) px.
    [Arguments("offsets", 310, 145)]
    // A sheet whose only widths are its default: 141 px columns and 35 px rows.
    [Arguments("defaults", 282, 105)]
    // Rows of 25 pt and 10.1 pt are 41 px and 16 px: Excel drops the part of a pixel.
    [Arguments("rounding", 300, 128)]
    // A from offset wider and taller than its cell ends at the cell's edge.
    [Arguments("clamp", 196, 120)]
    public async Task A_picture_reports_the_size_of_its_anchor_box(string fixture, int width, int height)
    {
        foreach (var editAs in EditAsValues)
        {
            var picture = Load(Case(fixture, editAs), AtExcelDpi);

            await Assert.That($"{editAs}: {picture.Width}x{picture.Height}")
                .IsEqualTo($"{editAs}: {width}x{height}");
        }
    }

    [Test]
    [Arguments("grid", 114, 210)]
    [Arguments("hidden", 114, 210)]
    [Arguments("offsets", 124, 215)]
    [Arguments("defaults", 282, 70)]
    [Arguments("clamp", 228, 240)]
    public async Task An_absolute_picture_reports_the_position_of_its_from_marker(string fixture, int left, int top)
    {
        var picture = Load(Case(fixture, "absolute"), AtExcelDpi);

        await Assert.That(picture.Placement).IsEqualTo(XLPicturePlacement.FreeFloating);
        await Assert.That($"{picture.Left},{picture.Top}").IsEqualTo($"{left},{top}");
    }

    [Test]
    public async Task A_box_with_no_visible_width_falls_back_to_the_xfrm()
    {
        // Columns B and C are hidden, so Excel shows the picture 0 px wide. A size of nothing is
        // not one the picture could be written back with, so the width comes from a:xfrm (8 px at
        // 96 DPI, 10 px at 120). The height still comes from the anchor.
        foreach (var editAs in EditAsValues)
        {
            var picture = Load(Case("allHidden", editAs), AtExcelDpi);

            await Assert.That($"{editAs}: {picture.Width}x{picture.Height}")
                .IsEqualTo($"{editAs}: 10x140");
        }
    }

    [Test]
    [Arguments("oneCell", XLPicturePlacement.Move)]
    [Arguments("absolute", XLPicturePlacement.FreeFloating)]
    [Arguments("twoCell", XLPicturePlacement.MoveAndSize)]
    public async Task The_issue_fixture_reports_its_anchor_box_at_96_dpi(string editAs, XLPicturePlacement placement)
    {
        // The #780 fixture: B8 to D12 on a sheet with no stated column width and rows of 15 pt,
        // which is 2 columns of 64 px and 4 rows of 20 px at 96 DPI. Excel, at its 120 DPI, measured
        // the box as 96 x 57.6 pt, which is its own 2 columns of 80 px and 4 rows of 24 px.
        var picture = Load(Case("issue", editAs), new LoadOptions());

        await Assert.That(picture.Placement).IsEqualTo(placement);
        await Assert.That($"{picture.Width}x{picture.Height}").IsEqualTo("128x80");
    }

    [Test]
    public async Task The_issue_fixture_places_an_absolute_picture_at_its_from_marker()
    {
        // Column A is 64 px, and rows 1 to 7 are 20 px each.
        var picture = Load(Case("issue", "absolute"), new LoadOptions());

        await Assert.That($"{picture.Left},{picture.Top}").IsEqualTo("64,140");
    }

    [Test]
    [Arguments("oneCell")]
    [Arguments("absolute")]
    [Arguments("twoCell")]
    public async Task Moving_the_picture_keeps_the_size_it_was_shown_at(string editAs)
    {
        // The repro of #782: MoveTo drops the kept anchor, and the one-cell anchor written in its
        // place takes the picture's size, which must be the 128 x 80 px box and not the 8 px xfrm.
        using var saved = LoadAndSave(Case("issue", editAs), ws => ws.Picture("Logo").MoveTo(ws.Cell("F2")));

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:oneCellAnchor>");
        await Assert.That(drawing).Contains($"<xdr:ext cx=\"{128 * 9525}\" cy=\"{80 * 9525}\" />");
    }

    [Test]
    public async Task Scaling_the_picture_scales_the_size_it_was_shown_at()
    {
        using var saved = LoadAndSave(Case("issue", "oneCell"), ws => ws.Picture("Logo").Scale(0.5));

        await Assert.That(saved.ReadPart(Drawing1))
            .Contains($"<xdr:ext cx=\"{64 * 9525}\" cy=\"{40 * 9525}\" />");
    }

    [Test]
    [Arguments("oneCell")]
    [Arguments("absolute")]
    public async Task Reading_the_anchor_size_keeps_the_loaded_anchor(string editAs)
    {
        // #780: a picture nobody edits is written back into the anchor it came from.
        using var saved = LoadAndSave(Case("offsets", editAs), _ => { });

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains($"<xdr:twoCellAnchor editAs=\"{editAs}\">");
        await Assert.That(drawing).Contains("<a:ext cx=\"76200\" cy=\"76200\" />");
    }

    private static IXLPicture Load(MemoryStream package, LoadOptions options) =>
        new XLWorkbook(package, options).Worksheet(1).Picture("Logo");

    private static MemoryStream LoadAndSave(MemoryStream source, Action<IXLWorksheet> edit)
    {
        using (source)
        {
            var saved = new MemoryStream();
            using var wb = new XLWorkbook(source);
            edit(wb.Worksheet(1));
            wb.SaveAs(saved);
            return saved;
        }
    }

    /// <summary>The fixtures measured in Excel. Pixel offsets are at 96 DPI, where XLibur wrote them.</summary>
    private static MemoryStream Case(string name, string editAs) => name switch
    {
        "issue" => Fixture(editAs, _ => { }, "B8", default, "D12", default),
        "grid" => Fixture(editAs, Grid, "B8", default, "D12", default),
        "hidden" => Fixture(editAs, ws =>
        {
            Grid(ws);
            ws.Column(3).Hide();
            ws.Row(10).Hide();
        }, "B8", default, "D12", default),
        "offsets" => Fixture(editAs, Grid, "B8", new Point(8, 4), "D12", new Point(16, 8)),
        "defaults" => Fixture(editAs, ws =>
        {
            ws.RowHeight = 21;
            ws.ColumnWidth = 15;
        }, "C3", default, "E6", default),
        "rounding" => Fixture(editAs, ws =>
        {
            Grid(ws);
            ws.Row(8).Height = 25;
            ws.Row(9).Height = 10.1;
            ws.Row(10).Height = 25;
        }, "B8", default, "D12", default),
        "clamp" => Fixture(editAs, Grid, "B8", new Point(200, 60), "D12", new Point(8, 8)),
        "allHidden" => Fixture(editAs, ws =>
        {
            Grid(ws);
            ws.Column(2).Hide();
            ws.Column(3).Hide();
        }, "B8", default, "D12", default),
        _ => throw new ArgumentOutOfRangeException(nameof(name)),
    };

    /// <summary>
    /// Rows of 18 pt and columns 12 characters wide by default, with column C 20 characters wide and
    /// row 9 30 pt high.
    /// </summary>
    private static void Grid(IXLWorksheet ws)
    {
        ws.RowHeight = 18;
        ws.ColumnWidth = 12;
        ws.Column(3).Width = 20;
        ws.Row(9).Height = 30;
    }

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    /// <summary>
    /// An 8×8 px picture in a two-cell anchor with the given <c>editAs</c>, its <c>a:xfrm</c> at
    /// offset 0,0 and the native size, as XLibur writes it.
    /// </summary>
    private static MemoryStream Fixture(string editAs, Action<IXLWorksheet> grid,
        string from, Point fromOffset, string to, Point toOffset)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            grid(ws);
            ws.AddPicture(new MemoryStream(Png), "Logo")
                .MoveTo(ws.Cell(from), fromOffset, ws.Cell(to), toOffset);
            wb.SaveAs(package);
        }

        return package.RewritePart(Drawing1,
            xml => xml.Replace("<xdr:twoCellAnchor>", $"<xdr:twoCellAnchor editAs=\"{editAs}\">"));
    }
}
