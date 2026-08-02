using System.Globalization;
using System.Threading;
using XLibur.Excel;
using XLibur.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class ColorTests
{
    [Test]
    public async Task ColorEqualOperatorInPlace()
    {
        await Assert.That(XLColor.Black == XLColor.Black).IsTrue();
    }

    [Test]
    public async Task ColorNotEqualOperatorInPlace()
    {
        await Assert.That(XLColor.Black != XLColor.Black).IsFalse();
    }

    [Test]
    public async Task ColorNamedVsHTML()
    {
        await Assert.That(XLColor.Black == XLColor.FromHtml("#000000")).IsTrue();
    }

    [Test]
    public async Task DefaultColorIndex64isTransparentWhite()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var color = ws.FirstCell().Style.Fill.BackgroundColor;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(color.Indexed).IsEqualTo(64);
        await Assert.That(color.Color).IsEqualTo(Color.Transparent);
    }

    [Test]
    public async Task CanConvertXLColorToColorType()
    {
        var xlColor1 = XLColor.Red;
        var xlColor2 = XLColor.FromIndex(20);
        var xlColor3 = XLColor.FromTheme(XLThemeColor.Accent1);
        var xlColor4 = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

        var color1 = new ForegroundColor().FromXLiburColor<ForegroundColor>(xlColor1);
        var color2 = new ForegroundColor().FromXLiburColor<ForegroundColor>(xlColor2);
        var color3 = new BackgroundColor().FromXLiburColor<BackgroundColor>(xlColor3);
        var color4 = new BackgroundColor().FromXLiburColor<BackgroundColor>(xlColor4);

        await Assert.That(color1.Rgb!.Value).IsEqualTo("FFFF0000");
        await Assert.That(color1.Indexed).IsNull();
        await Assert.That(color1.Theme).IsNull();
        await Assert.That(color1.Tint).IsNull();

        await Assert.That(color2.Rgb).IsNull();
        await Assert.That(color2.Indexed!.Value).IsEqualTo(20u);
        await Assert.That(color2.Theme).IsNull();
        await Assert.That(color2.Tint).IsNull();

        await Assert.That(color3.Rgb).IsNull();
        await Assert.That(color3.Indexed).IsNull();
        await Assert.That(color3.Theme!.Value).IsEqualTo(4u);
        await Assert.That(color3.Tint).IsNull();

        await Assert.That(color4.Rgb).IsNull();
        await Assert.That(color4.Indexed).IsNull();
        await Assert.That(color4.Theme!.Value).IsEqualTo(5u);
        await Assert.That(color4.Tint!.Value).IsEqualTo(0.4);
    }

    [Test]
    public async Task CanConvertXlColorToX14ColorType()
    {
        var xlColor1 = XLColor.Red;
        var xlColor2 = XLColor.FromIndex(20);
        var xlColor3 = XLColor.FromTheme(XLThemeColor.Accent1);
        var xlColor4 = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

        var color1 = new X14.AxisColor().FromXLiburColor<X14.AxisColor>(xlColor1);
        var color2 = new X14.BorderColor().FromXLiburColor<X14.BorderColor>(xlColor2);
        var color3 = new X14.FillColor().FromXLiburColor<X14.FillColor>(xlColor3);
        var color4 = new X14.HighMarkerColor().FromXLiburColor<X14.HighMarkerColor>(xlColor4);

        await Assert.That(color1.Rgb!.Value).IsEqualTo("FFFF0000");
        await Assert.That(color1.Indexed).IsNull();
        await Assert.That(color1.Theme).IsNull();
        await Assert.That(color1.Tint).IsNull();

        await Assert.That(color2.Rgb).IsNull();
        await Assert.That(color2.Indexed!.Value).IsEqualTo(20u);
        await Assert.That(color2.Theme).IsNull();
        await Assert.That(color2.Tint).IsNull();

        await Assert.That(color3.Rgb).IsNull();
        await Assert.That(color3.Indexed).IsNull();
        await Assert.That(color3.Theme!.Value).IsEqualTo(4u);
        await Assert.That(color3.Tint).IsNull();

        await Assert.That(color4.Rgb).IsNull();
        await Assert.That(color4.Indexed).IsNull();
        await Assert.That(color4.Theme!.Value).IsEqualTo(5u);
        await Assert.That(color4.Tint!.Value).IsEqualTo(0.4);
    }

    [Test]
    public async Task CanConvertColorTypeToXlColor()
    {
        var color1 = new ForegroundColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000") };
        var color2 = new ForegroundColor { Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20) };
        var color3 = new BackgroundColor { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
        var color4 = new BackgroundColor
        {
            Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
            Tint = new DocumentFormat.OpenXml.DoubleValue(0.4)
        };

        var xlColor1 = color1.ToXLiburColor();
        var xlColor2 = color2.ToXLiburColor();
        var xlColor3 = color3.ToXLiburColor();
        var xlColor4 = color4.ToXLiburColor();

        await Assert.That(xlColor1.ColorType).IsEqualTo(XLColorType.Color);
        await Assert.That(xlColor1.Color).IsEqualTo(XLColor.Red.Color);

        await Assert.That(xlColor2.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(xlColor2.Indexed).IsEqualTo(20);

        await Assert.That(xlColor3.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(xlColor3.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(xlColor3.ThemeTint).IsEqualTo(0).Within(XLHelper.Epsilon);

        await Assert.That(xlColor4.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(xlColor4.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(xlColor4.ThemeTint).IsEqualTo(0.4).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task CanConvertX14ColorTypeToXlColor()
    {
        var color1 = new X14.AxisColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000") };
        var color2 = new X14.BorderColor { Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20) };
        var color3 = new X14.FillColor { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
        var color4 = new X14.HighMarkerColor
        {
            Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
            Tint = new DocumentFormat.OpenXml.DoubleValue(0.4)
        };

        var xlColor1 = color1.ToXLiburColor();
        var xlColor2 = color2.ToXLiburColor();
        var xlColor3 = color3.ToXLiburColor();
        var xlColor4 = color4.ToXLiburColor();

        await Assert.That(xlColor1.ColorType).IsEqualTo(XLColorType.Color);
        await Assert.That(xlColor1.Color).IsEqualTo(XLColor.Red.Color);

        await Assert.That(xlColor2.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(xlColor2.Indexed).IsEqualTo(20);

        await Assert.That(xlColor3.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(xlColor3.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(xlColor3.ThemeTint).IsEqualTo(0).Within(XLHelper.Epsilon);

        await Assert.That(xlColor4.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(xlColor4.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(xlColor4.ThemeTint).IsEqualTo(0.4).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task CanParseColorWithHashAsCultureLineSeparator()
    {
        // https://github.com/XLibur/XLibur/issues/675
        var culture = CultureInfo.CreateSpecificCulture("en-US");
        culture.TextInfo.ListSeparator = "#";
        Thread.CurrentThread.CurrentCulture = culture;
        var color = XLColor.FromHtml("#FF008000");
        await Assert.That(color).IsEqualTo(XLColor.Green);
    }
}
