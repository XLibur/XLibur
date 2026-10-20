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
    // A range to the last column is the sheet's default: B and C are 114 px each, as Excel converts
    // a width of 12.7109375 at a digit width of 9 px, not 115 px.
    [Arguments("rangeDefault", 228)]
    // B and C state the same width as the range after them: 82 px each, not 83.
    [Arguments("rangeExplicit", 164)]
    // B and C state 12.710625, 114 px, which is not the 115 px default of defaultColWidth="12.78"
    // although the model holds the same width for both.
    [Arguments("explicitAtDefault", 228)]
    // C onwards is 0 wide, so the box is column B alone.
    [Arguments("rangeZero", 114)]
    // A width of 0 that is not hidden is drawn as 0 px, the same as a hidden column.
    [Arguments("zeroWidth", 80)]
    // Excel takes the digit width from the first font of the stylesheet, Calibri 11 here, although
    // the default cell format uses Arial 20: 115 px columns.
    [Arguments("cellFont", 230)]
    // ... and although the Normal style uses Arial 20.
    [Arguments("normalStyleFont", 230)]
    // baseColWidth="12" is 12 digits and padding, rounded up to a multiple of 8 px: 120 px.
    [Arguments("cellFontBaseWidth", 240)]
    // No stated width is baseColWidth 8: 80 px.
    [Arguments("cellFontNoWidth", 160)]
    public async Task A_picture_measures_its_columns_as_excel_draws_them(string fixture, int width)
    {
        // Measured in Excel at 120 DPI. Only the width: the rows of these sheets are at a default
        // height the file does not state, which Excel derives from the font.
        var picture = Load(StatedColumnsCase(fixture), AtExcelDpi);

        await Assert.That(picture.Width).IsEqualTo(width);
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
    public async Task Resizing_a_row_inside_a_one_cell_picture_keeps_the_size_it_was_shown_at()
    {
        // Excel keeps a oneCell picture 96 x 57.6 pt when row 9 under it is made 40 pt high. The
        // kept markers would now stretch it, so the anchor is written from the model, whose size
        // must be the 128 x 80 px box and not the 8 px xfrm.
        using var saved = LoadAndSave(Case("issue", "oneCell"), ws => ws.Row(9).Height = 40);

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:oneCellAnchor>");
        await Assert.That(drawing).Contains("<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row>");
        await Assert.That(drawing).Contains($"<xdr:ext cx=\"{128 * 9525}\" cy=\"{80 * 9525}\" />");
    }

    [Test]
    public async Task Resizing_a_row_above_an_absolute_picture_keeps_where_it_was_shown()
    {
        // Excel keeps an absolute picture at 48, 100.8 pt when row 3 above it is made 40 pt high.
        // The anchor written from the model must hold the box's position and size, not the xfrm's
        // 0,0 and 8 px.
        using var saved = LoadAndSave(Case("issue", "absolute"), ws => ws.Row(3).Height = 40);

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:absoluteAnchor>");
        await Assert.That(drawing).Contains($"<xdr:pos x=\"{64 * 9525}\" y=\"{140 * 9525}\" />");
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

    [Test]
    public async Task A_picture_in_an_aptos_narrow_workbook_reports_the_size_excel_shows_it_at()
    {
        // picture-aptos-narrow.xlsx: Excel 16.0 (build 20430) at 120 DPI, a new workbook (Aptos Narrow
        // 11, no stated column width, rows of 14.4 pt), Shapes.AddPicture at left 100 pt, top 50 pt,
        // 150 x 75 pt. Excel writes a oneCell anchor from C4 to F9 and draws columns of 48 pt (80 px:
        // its ColumnWidth 8.11 is a digit width of 9 px) and rows of 24 px. The box is 250 x 125 px
        // at 120 DPI, whether or not the machine has Aptos Narrow installed (#793).
        var picture = Load(AptosNarrowPicture(), AtExcelDpi, "Picture 2");

        await Assert.That($"{picture.Width}x{picture.Height}").IsEqualTo("250x125");
    }

    [Test]
    public async Task A_picture_in_an_aptos_narrow_workbook_measures_its_columns_at_96_dpi()
    {
        // At 96 DPI Aptos Narrow 11 has a digit width of 7 px, so a column is 64 px and the three
        // columns and the offsets of the box are Excel's 150 pt, 200 px. Only the width: the rows
        // are at a default height the file does not state, which Excel derives from the font.
        var picture = Load(AptosNarrowPicture(), new LoadOptions(), "Picture 2");

        await Assert.That(picture.Width).IsEqualTo(200);
    }

    [Test]
    public async Task Moving_a_picture_in_an_aptos_narrow_workbook_keeps_the_size_excel_shows_it_at()
    {
        using var saved = LoadAndSave(AptosNarrowPicture(), AtExcelDpi,
            ws => ws.Picture("Picture 2").MoveTo(ws.Cell("H2")));

        // 7620 EMU is a pixel at 120 DPI.
        await Assert.That(saved.ReadPart(Drawing1))
            .Contains($"<xdr:ext cx=\"{250 * 7620}\" cy=\"{125 * 7620}\" />");
    }

    [Test]
    public async Task Moving_a_picture_in_an_aptos_narrow_workbook_keeps_its_width_at_96_dpi()
    {
        // The 200 px the box measures at 96 DPI (above), by the rule: Excel on this machine runs at
        // 120 DPI and cannot be measured at 96. The height is left unpinned, as above. 9525 EMU is a
        // pixel at 96 DPI.
        using var saved = LoadAndSave(AptosNarrowPicture(), ws => ws.Picture("Picture 2").MoveTo(ws.Cell("H2")));

        await Assert.That(saved.ReadPart(Drawing1)).Contains($"<xdr:ext cx=\"{200 * 9525}\" ");
    }

    private static MemoryStream AptosNarrowPicture()
    {
        using var resource = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\Drawings\picture-aptos-narrow.xlsx"));
        var package = new MemoryStream();
        resource.CopyTo(package);
        package.Position = 0;
        return package;
    }

    private static IXLPicture Load(MemoryStream package, LoadOptions options, string name = "Logo") =>
        new XLWorkbook(package, options).Worksheet(1).Picture(name);

    private static MemoryStream LoadAndSave(MemoryStream source, Action<IXLWorksheet> edit) =>
        LoadAndSave(source, new LoadOptions(), edit);

    private static MemoryStream LoadAndSave(MemoryStream source, LoadOptions options, Action<IXLWorksheet> edit)
    {
        using (source)
        {
            var saved = new MemoryStream();
            using var wb = new XLWorkbook(source, options);
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

    private const string StatedRowHeight = "<x:sheetFormatPr defaultRowHeight=\"15\" />";

    /// <summary>
    /// The B8:D12 <c>oneCell</c> picture of <c>issue</c>, on a sheet whose <c>sheetFormatPr</c>,
    /// <c>cols</c> and fonts are written as Excel or another writer may state them.
    /// </summary>
    private static MemoryStream StatedColumnsCase(string name) => name switch
    {
        "rangeDefault" => WithColumns(StatedRowHeight,
            "<x:col min=\"1\" max=\"16384\" width=\"12.7109375\" customWidth=\"1\" />"),
        "rangeExplicit" => WithColumns(StatedRowHeight,
            "<x:col min=\"2\" max=\"3\" width=\"9.140625\" customWidth=\"1\" />" +
            "<x:col min=\"4\" max=\"16384\" width=\"9.140625\" customWidth=\"1\" />"),
        "explicitAtDefault" => WithColumns("<x:sheetFormatPr defaultColWidth=\"12.78\" defaultRowHeight=\"15\" />",
            "<x:col min=\"2\" max=\"3\" width=\"12.710625\" customWidth=\"1\" />"),
        "rangeZero" => WithColumns(StatedRowHeight,
            "<x:col min=\"2\" max=\"2\" width=\"12.7109375\" customWidth=\"1\" />" +
            "<x:col min=\"3\" max=\"16384\" width=\"0\" customWidth=\"1\" />"),
        "zeroWidth" => WithColumns(StatedRowHeight, "<x:col min=\"3\" max=\"3\" width=\"0\" customWidth=\"1\" />"),
        "cellFont" => WithArialCellFormat(WithColumns(
            "<x:sheetFormatPr defaultColWidth=\"12.78\" defaultRowHeight=\"15\" />", null)),
        "normalStyleFont" => WithArialNormalStyle(WithColumns(
            "<x:sheetFormatPr defaultColWidth=\"12.78\" defaultRowHeight=\"15\" />", null)),
        "cellFontBaseWidth" => WithArialCellFormat(WithColumns(
            "<x:sheetFormatPr baseColWidth=\"12\" defaultRowHeight=\"15\" />", null)),
        "cellFontNoWidth" => WithArialCellFormat(WithColumns(StatedRowHeight, null)),
        _ => throw new ArgumentOutOfRangeException(nameof(name)),
    };

    private static MemoryStream WithColumns(string sheetFormat, string? columns) =>
        Case("issue", "oneCell").RewriteSheet1(xml => xml
            .Replace("<x:sheetFormatPr defaultRowHeight=\"15\" />", sheetFormat)
            .Replace("<x:sheetData", columns is null ? "<x:sheetData" : $"<x:cols>{columns}</x:cols><x:sheetData"));

    private const string Calibri11 =
        "<x:font><x:sz val=\"11\" /><x:color rgb=\"FF000000\" /><x:name val=\"Calibri\" /><x:family val=\"2\" /></x:font>";

    private const string Arial20 =
        "<x:font><x:sz val=\"20\" /><x:color rgb=\"FF000000\" /><x:name val=\"Arial\" /><x:family val=\"2\" /></x:font>";

    /// <summary>The first font is Calibri 11, the default cell format Arial 20.</summary>
    private static MemoryStream WithArialCellFormat(MemoryStream package) => package.RewritePart("xl/styles.xml",
        xml => WithStyles(xml,
            "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
            "<x:cellXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" /></x:cellXfs>",
            "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>"));

    /// <summary>The first font is Calibri 11, the Normal style (the second style format) Arial 20.</summary>
    private static MemoryStream WithArialNormalStyle(MemoryStream package) => package.RewritePart("xl/styles.xml",
        xml => WithStyles(xml,
            "<x:cellStyleXfs count=\"2\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" />" +
            "<x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
            "<x:cellXfs count=\"2\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" />" +
            "<x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"1\" /></x:cellXfs>",
            "<x:cellStyles count=\"2\"><x:cellStyle name=\"Plain\" xfId=\"0\" />" +
            "<x:cellStyle name=\"Normal\" xfId=\"1\" builtinId=\"0\" /></x:cellStyles>"));

    private static string WithStyles(string xml, string cellStyleXfs, string cellXfs, string cellStyles) =>
        StylesPart.ReplaceElements(xml,
            ("fonts", $"<x:fonts count=\"2\">{Calibri11}{Arial20}</x:fonts>"),
            ("cellStyleXfs", cellStyleXfs),
            ("cellXfs", cellXfs),
            ("cellStyles", cellStyles));

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
