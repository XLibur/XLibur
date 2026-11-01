using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// The styling properties of <see cref="IXLPicture"/> as a model: what they read, what they refuse,
/// and what a copy of a picture takes (spec 17).
/// </summary>
public class PictureStylingModelTests
{
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    [Test]
    public async Task A_new_picture_has_no_styling()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));

        await Assert.That(Describe(picture)).IsEqualTo(
            "border - - - no=False fill - no=False rot - flip -/- transparency - recolor - lum -/- crop - shadow -");
    }

    [Test]
    public async Task Each_styling_property_reads_back_what_was_assigned()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.BorderColor = XLColor.Red;
        picture.BorderWidthPt = 1.5;
        picture.BorderDash = XLLineDashStyle.Dash;
        picture.FillColor = XLColor.Yellow;
        picture.Rotation = 20;
        picture.FlipHorizontal = true;
        picture.FlipVertical = false;
        picture.Transparency = 35;
        picture.Recolor = XLPictureRecolor.Grayscale;
        picture.Brightness = 20;
        picture.Contrast = -10;
        picture.Crop = new XLCrop(10, 0, 5, 2.5);
        picture.Shadow = new XLPictureShadow(XLColor.Black, 60, 4, 3, 45);

        await Assert.That(Describe(picture)).IsEqualTo(
            "border FFFF0000 1.5 Dash no=False fill FFFFFF00 no=False rot 20 flip True/False " +
            "transparency 35 recolor Grayscale lum 20/-10 crop 10,0,5,2.5 shadow FF000000 60 4 3 45");
    }

    [Test]
    [Arguments(-30.0, 330.0)]
    [Arguments(400.0, 40.0)]
    [Arguments(360.0, 0.0)]
    [Arguments(-1e-14, 0.0)]
    public async Task A_rotation_is_brought_into_0_up_to_360_degrees(double assigned, double read)
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.Rotation = assigned;

        await Assert.That(picture.Rotation).IsEqualTo(read);
    }

    [Test]
    public async Task A_custom_recolour_cannot_be_assigned()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));

        await Assert.That(() => picture.Recolor = XLPictureRecolor.Custom).Throws<ArgumentException>();
    }

    [Test]
    public async Task Values_outside_their_range_are_refused()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));

        await Assert.That(() => picture.Transparency = 101).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Brightness = -100.5).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Contrast = double.NaN).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.BorderWidthPt = -1).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Crop = new XLCrop(60, 0, 40, 0)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Rotation = double.PositiveInfinity).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Shadow = new XLPictureShadow(XLColor.Black, 0, -1, 0, 0))
            .Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Shadow = new XLPictureShadow(XLColor.Black, 0, 1e6, 0, 0))
            .Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => picture.Crop = new XLCrop(-1e7, 0, 0, 0)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task No_border_and_a_border_colour_exclude_each_other()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.BorderColor = XLColor.Red;
        picture.NoBorder = true;

        await Assert.That(picture.BorderColor).IsNull();

        picture.BorderColor = XLColor.Blue;
        await Assert.That(picture.NoBorder).IsFalse();
    }

    [Test]
    public async Task No_fill_and_a_fill_colour_exclude_each_other()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.FillColor = XLColor.Red;
        picture.NoFill = true;

        await Assert.That(picture.FillColor).IsNull();

        picture.FillColor = XLColor.Blue;
        await Assert.That(picture.NoFill).IsFalse();
    }

    [Test]
    public async Task Washout_is_a_brightness_and_a_contrast()
    {
        // Excel's Washout preset is brightness 70 % and contrast -70 % (an Excel-authored fixture).
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.Recolor = XLPictureRecolor.Washout;

        await Assert.That($"{picture.Brightness}/{picture.Contrast}").IsEqualTo("70/-70");

        picture.Brightness = 10;
        await Assert.That(picture.Recolor).IsEqualTo(XLPictureRecolor.None);
    }

    [Test]
    public async Task Washout_brightness_and_contrast_read_as_washout()
    {
        // As the loader reads them, so the model and a reloaded file agree.
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.Brightness = 70;
        picture.Contrast = -70;

        await Assert.That(picture.Recolor).IsEqualTo(XLPictureRecolor.Washout);
    }

    [Test]
    public async Task Leaving_washout_takes_the_brightness_and_contrast_back()
    {
        using var wb = new XLWorkbook();
        var picture = wb.AddWorksheet().AddPicture(new MemoryStream(Png));
        picture.Recolor = XLPictureRecolor.Washout;
        picture.Recolor = XLPictureRecolor.Grayscale;

        await Assert.That($"{picture.Brightness?.ToString() ?? "-"}/{picture.Contrast?.ToString() ?? "-"}")
            .IsEqualTo("-/-");
    }

    [Test]
    public async Task A_copy_of_a_picture_has_its_styling()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var picture = ws.AddPicture(new MemoryStream(Png)).MoveTo(ws.Cell("B2"));
        picture.BorderColor = XLColor.Red;
        picture.Rotation = 90;
        picture.Recolor = XLPictureRecolor.Sepia;
        picture.Crop = new XLCrop(1, 2, 3, 4);

        var duplicate = picture.Duplicate();
        var copy = picture.CopyTo(wb.AddWorksheet());

        await Assert.That(Describe(duplicate)).IsEqualTo(Describe(picture));
        await Assert.That(Describe(copy)).IsEqualTo(Describe(picture));
    }

    [Test]
    public async Task A_copy_of_a_free_floating_picture_has_its_styling()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var picture = ws.AddPicture(new MemoryStream(Png)).MoveTo(30, 40);
        picture.BorderColor = XLColor.Red;
        picture.Recolor = XLPictureRecolor.Grayscale;

        var duplicate = picture.Duplicate();

        await Assert.That(Describe(duplicate)).IsEqualTo(Describe(picture));
        await Assert.That($"{duplicate.Left},{duplicate.Top}").IsEqualTo("30,40");
    }

    /// <summary>Every styling property of <paramref name="picture"/>, in one line; <c>-</c> for none.</summary>
    internal static string Describe(IXLPicture picture)
    {
        static string V(object? value) => value switch
        {
            null => "-",
            XLColor color => color.ToString(),
            XLCrop crop => $"{crop.Left},{crop.Top},{crop.Right},{crop.Bottom}",
            XLPictureShadow shadow =>
                $"{shadow.Color} {shadow.TransparencyPct} {shadow.BlurPt} {shadow.DistancePt} {shadow.AngleDeg}",
            IFormattable formattable => formattable.ToString(null, System.Globalization.CultureInfo.InvariantCulture),
            _ => value.ToString() ?? "-",
        };

        return $"border {V(picture.BorderColor)} {V(picture.BorderWidthPt)} {V(picture.BorderDash)} no={picture.NoBorder} " +
               $"fill {V(picture.FillColor)} no={picture.NoFill} rot {V(picture.Rotation)} " +
               $"flip {V(picture.FlipHorizontal)}/{V(picture.FlipVertical)} transparency {V(picture.Transparency)} " +
               $"recolor {V(picture.Recolor)} lum {V(picture.Brightness)}/{V(picture.Contrast)} crop {V(picture.Crop)} " +
               $"shadow {V(picture.Shadow)}";
    }
}
