using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// A picture Excel anchors as <c>xdr:twoCellAnchor</c> with <c>editAs="oneCell"</c> or
/// <c>editAs="absolute"</c> keeps that anchor through a load and save (issue #780).
/// </summary>
/// <remarks>
/// Excel places and sizes every two-cell anchor from its <c>from</c>/<c>to</c> markers, whatever
/// <c>editAs</c> says; <c>editAs</c> only decides how the picture follows a later grid edit. The
/// fixture's <c>a:xfrm</c> says 8×8 px while the markers span B8 to D12, which is what XLibur and
/// ClosedXML write, so a writer that sizes the anchor from <c>xfrm</c> shrinks the picture to 8 px.
/// </remarks>
public class PictureEditAsRoundTripTests
{
    private const string Drawing1 = "xl/drawings/drawing1.xml";

    [Test]
    [Arguments("oneCell")]
    [Arguments("absolute")]
    public async Task An_untouched_picture_keeps_its_two_cell_anchor_and_edit_as(string editAs)
    {
        using var saved = LoadAndSave(PictureFromB8ToD12(editAs), _ => { });

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains($"<xdr:twoCellAnchor editAs=\"{editAs}\">");
        await Assert.That(drawing).DoesNotContain("oneCellAnchor");
        await Assert.That(drawing).DoesNotContain("absoluteAnchor");
        await Assert.That(Markers(drawing)).IsEqualTo("from 1,7 to 3,11");
    }

    [Test]
    [Arguments("oneCell")]
    [Arguments("absolute")]
    public async Task An_untouched_picture_keeps_its_shape_properties(string editAs)
    {
        // The real offset and size, as Excel writes them, so a regenerated xdr:pic would differ.
        const string ExcelXfrm = "<a:xfrm><a:off x=\"609600\" y=\"1280160\" /><a:ext cx=\"1219200\" cy=\"731520\" /></a:xfrm>";
        using var source = PictureFromB8ToD12(editAs)
            .RewritePart(Drawing1, xml => Regex.Replace(xml, "<a:xfrm>.*?</a:xfrm>", ExcelXfrm));
        var before = Xfrm(source.ReadPart(Drawing1));
        await Assert.That(before).IsEqualTo(ExcelXfrm);

        using var saved = LoadAndSave(source, _ => { });

        await Assert.That(Xfrm(saved.ReadPart(Drawing1))).IsEqualTo(before);
    }

    [Test]
    [Arguments("oneCell", XLPicturePlacement.Move)]
    [Arguments("absolute", XLPicturePlacement.FreeFloating)]
    public async Task A_kept_anchor_reloads_with_the_same_placement(string editAs, XLPicturePlacement placement)
    {
        using var saved = LoadAndSave(PictureFromB8ToD12(editAs), _ => { });

        using var wb = new XLWorkbook(saved);
        await Assert.That(wb.Worksheet(1).Picture("Logo").Placement).IsEqualTo(placement);
    }

    [Test]
    public async Task Inserting_rows_above_a_one_cell_picture_moves_both_corners()
    {
        // Excel moves an editAs="oneCell" picture down and keeps its size, so both markers move by
        // the rows inserted above it.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Row(3).InsertRowsAbove(3));

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:twoCellAnchor editAs=\"oneCell\">");
        await Assert.That(Markers(drawing)).IsEqualTo("from 1,10 to 3,14");
    }

    [Test]
    public async Task Inserting_rows_inside_a_one_cell_picture_keeps_its_span()
    {
        // Rows inserted under the picture do not stretch it: the to marker moves by the from
        // marker's shift, which here is none.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Row(10).InsertRowsAbove(2));

        await Assert.That(Markers(saved.ReadPart(Drawing1))).IsEqualTo("from 1,7 to 3,11");
    }

    [Test]
    public async Task Inserting_columns_left_of_a_one_cell_picture_moves_both_corners()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Column(1).InsertColumnsBefore(2));

        await Assert.That(Markers(saved.ReadPart(Drawing1))).IsEqualTo("from 3,7 to 5,11");
    }

    [Test]
    public async Task Inserting_rows_above_an_absolute_picture_does_not_move_it()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute"), ws => ws.Row(3).InsertRowsAbove(3));

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:twoCellAnchor editAs=\"absolute\">");
        await Assert.That(Markers(drawing)).IsEqualTo("from 1,7 to 3,11");
    }

    [Test]
    public async Task Renaming_an_untouched_picture_keeps_its_anchor()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Picture("Logo").Name = "Brand");

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:twoCellAnchor editAs=\"oneCell\">");
        await Assert.That(drawing).Contains("name=\"Brand\"");
        await Assert.That(drawing).DoesNotContain("name=\"Logo\"");
    }

    [Test]
    public async Task Moving_the_picture_writes_the_anchor_from_the_model()
    {
        // Once the caller places the picture, the model decides the anchor: MoveTo(cell) is a
        // one-cell anchor at the new cell.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Picture("Logo").MoveTo(ws.Cell("F2")));

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:oneCellAnchor>");
        await Assert.That(drawing).DoesNotContain("editAs");
        await Assert.That(drawing).Contains("<xdr:from><xdr:col>5</xdr:col>");
    }

    [Test]
    public async Task Changing_the_placement_writes_the_anchor_from_the_model()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute"),
            ws => ws.Picture("Logo").Placement = XLPicturePlacement.Move);

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:oneCellAnchor>");
        await Assert.That(drawing).DoesNotContain("editAs");
    }

    [Test]
    public async Task Resizing_the_picture_writes_the_anchor_from_the_model()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Picture("Logo").Width = 40);

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:oneCellAnchor>");
        await Assert.That(drawing).Contains($"cx=\"{40 * 9525}\"");
    }

    [Test]
    public async Task An_explicit_two_cell_picture_still_stretches_with_inserted_rows()
    {
        // editAs="twoCell" is the default spelled out; the model says it back, so nothing is kept
        // and both corners move with the grid.
        using var saved = LoadAndSave(PictureFromB8ToD12("twoCell"), ws => ws.Row(10).InsertRowsAbove(2));

        await Assert.That(Markers(saved.ReadPart(Drawing1))).IsEqualTo("from 1,7 to 3,13");
    }

    [Test]
    public async Task Resizing_a_row_inside_a_one_cell_picture_writes_the_anchor_from_the_model()
    {
        // The kept markers would stretch the picture over the taller row; Excel keeps its size.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Row(9).Height = 40);

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:oneCellAnchor>");
        await Assert.That(drawing).DoesNotContain("editAs");
    }

    [Test]
    public async Task Resizing_a_column_inside_a_one_cell_picture_writes_the_anchor_from_the_model()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Column("C").Width = 30);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:oneCellAnchor>");
    }

    [Test]
    public async Task Resizing_a_row_above_a_one_cell_picture_keeps_its_anchor()
    {
        // A one-cell picture hangs off its top-left cell, so rows above it change neither its
        // markers nor its size.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws => ws.Row(2).Height = 40);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:twoCellAnchor editAs=\"oneCell\">");
    }

    [Test]
    public async Task Resizing_a_row_below_a_picture_keeps_its_anchor()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute"), ws => ws.Row(20).Height = 40);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:twoCellAnchor editAs=\"absolute\">");
    }

    [Test]
    public async Task Resizing_a_row_above_an_absolute_picture_writes_the_anchor_from_the_model()
    {
        // The kept markers would move the picture down with the taller row; Excel leaves it where
        // it was.
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute"), ws => ws.Row(1).Height = 40);

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:absoluteAnchor>");
        await Assert.That(drawing).DoesNotContain("editAs");
    }

    [Test]
    public async Task Hiding_a_column_left_of_an_absolute_picture_writes_the_anchor_from_the_model()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute"), ws => ws.Column("A").Hide());

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:absoluteAnchor>");
    }

    [Test]
    public async Task Changing_the_default_row_height_writes_the_anchor_from_the_model()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute"), ws => ws.RowHeight = 30);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:absoluteAnchor>");
    }

    [Test]
    public async Task Inserting_taller_rows_inside_a_one_cell_picture_writes_the_anchor_from_the_model()
    {
        // Keeping the span would put a 40 pt row inside the picture, so it would stretch.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell"), ws =>
        {
            ws.Row(10).InsertRowsAbove(1);
            ws.Row(10).Height = 40;
        });

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:oneCellAnchor>");
    }

    [Test]
    public async Task Resizing_a_row_above_an_absolute_picture_among_sized_rows_writes_the_anchor_from_the_model()
    {
        // Every row of the sheet has a set height, so the comparison has a size to check for each
        // row from the top of the sheet to the picture's bottom row.
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute", RowsOf20Pt(1, 30)), ws => ws.Row(5).Height = 40);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:absoluteAnchor>");
    }

    [Test]
    public async Task Resizing_a_row_below_an_absolute_picture_among_sized_rows_keeps_its_anchor()
    {
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute", RowsOf20Pt(1, 30)), ws => ws.Row(25).Height = 40);

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:twoCellAnchor editAs=\"absolute\">");
    }

    [Test]
    public async Task Inserting_rows_above_a_one_cell_picture_over_sized_rows_keeps_its_anchor()
    {
        // The sized rows under the picture move down with it, so its markers still count across
        // rows of the same heights.
        using var saved = LoadAndSave(PictureFromB8ToD12("oneCell", RowsOf20Pt(8, 12)), ws => ws.Row(3).InsertRowsAbove(3));

        var drawing = saved.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:twoCellAnchor editAs=\"oneCell\">");
        await Assert.That(Markers(drawing)).IsEqualTo("from 1,10 to 3,14");
    }

    [Test]
    public async Task Inserting_rows_above_an_absolute_picture_over_sized_rows_writes_the_anchor_from_the_model()
    {
        // The picture stays where it is, but the sized rows under it move down, so its kept markers
        // now count across rows of other heights.
        using var saved = LoadAndSave(PictureFromB8ToD12("absolute", RowsOf20Pt(8, 12)), ws => ws.Row(3).InsertRowsAbove(3));

        await Assert.That(saved.ReadPart(Drawing1)).Contains("<xdr:absoluteAnchor>");
    }

    [Test]
    public async Task Resizing_a_row_between_two_absolute_pictures_rewrites_only_the_one_below_it()
    {
        using var source = AbsolutePicturesAboveAndBelowRow20();

        using var saved = LoadAndSave(source, ws => ws.Row(20).Height = 40);

        var drawing = saved.ReadPart(Drawing1);
        var kept = Regex.Match(drawing, "<xdr:twoCellAnchor editAs=\"absolute\">.*?</xdr:twoCellAnchor>").Value;
        var rewritten = Regex.Match(drawing, "<xdr:absoluteAnchor>.*?</xdr:absoluteAnchor>").Value;
        await Assert.That(kept).Contains("name=\"Above\"");
        await Assert.That(rewritten).Contains("name=\"Below\"");
    }

    [Test]
    public async Task A_second_save_keeps_the_anchor()
    {
        using var source = PictureFromB8ToD12("oneCell");
        using var wb = new XLWorkbook(source);
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        wb.SaveAs(first);
        wb.SaveAs(second);

        var drawing = second.ReadPart(Drawing1);
        await Assert.That(drawing).Contains("<xdr:twoCellAnchor editAs=\"oneCell\">");
        await Assert.That(Markers(drawing)).IsEqualTo("from 1,7 to 3,11");
    }

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    /// <summary>
    /// An 8×8 px picture anchored B8 to D12 the way Excel anchors an inserted picture: a two-cell
    /// anchor whose <c>editAs</c> says how it follows the grid. This is the repro of ClosedXML #2895.
    /// </summary>
    private static MemoryStream PictureFromB8ToD12(string editAs, Action<IXLWorksheet>? setUp = null)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            setUp?.Invoke(ws);
            ws.AddPicture(new MemoryStream(Png), "Logo").MoveTo(ws.Cell("B8"), ws.Cell("D12"));
            wb.SaveAs(package);
        }

        return WithEditAs(package, editAs);
    }

    /// <summary>
    /// Two absolute pictures on a sheet whose rows 1 to 40 are 20 pt tall: <c>Above</c> from B2 to
    /// D6, and <c>Below</c> from B30 to D34.
    /// </summary>
    private static MemoryStream AbsolutePicturesAboveAndBelowRow20()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            RowsOf20Pt(1, 40)(ws);
            ws.AddPicture(new MemoryStream(Png), "Above").MoveTo(ws.Cell("B2"), ws.Cell("D6"));
            ws.AddPicture(new MemoryStream(Png), "Below").MoveTo(ws.Cell("B30"), ws.Cell("D34"));
            wb.SaveAs(package);
        }

        return WithEditAs(package, "absolute");
    }

    private static Action<IXLWorksheet> RowsOf20Pt(int first, int last) => ws =>
    {
        for (var row = first; row <= last; row++)
            ws.Row(row).Height = 20;
    };

    private static MemoryStream WithEditAs(MemoryStream package, string editAs) =>
        package.RewritePart(Drawing1,
            xml => xml.Replace("<xdr:twoCellAnchor>", $"<xdr:twoCellAnchor editAs=\"{editAs}\">"));

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

    /// <summary>The zero-based column and row of the anchor's two markers, e.g. <c>from 1,7 to 3,11</c>.</summary>
    private static string Markers(string drawing)
    {
        static string Marker(string drawing, string name)
        {
            var m = Regex.Match(drawing,
                $@"<xdr:{name}><xdr:col>(\d+)</xdr:col><xdr:colOff>-?\d+</xdr:colOff><xdr:row>(\d+)</xdr:row>");
            return m.Success ? $"{m.Groups[1].Value},{m.Groups[2].Value}" : "none";
        }

        return $"from {Marker(drawing, "from")} to {Marker(drawing, "to")}";
    }

    private static string Xfrm(string drawing) => Regex.Match(drawing, "<a:xfrm>.*?</a:xfrm>").Value;
}
