using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// <see cref="XLBorderStyleValues"/>, <see cref="XLColorType"/> and <see cref="XLThemeColor"/> are
/// public and <c>int</c>-backed, but <c>XLBorderKey</c> and <c>XLColorKey</c> store each of them in
/// a byte to keep the style key small. A public enum can be cast from an arbitrary <c>int</c> - the
/// C# compiler does not reject <c>(XLBorderStyleValues)999</c> - so the byte-narrowing checks its
/// input rather than wrapping an out-of-range value into whichever defined member it happens to
/// reduce to modulo 256.
/// </summary>
public class XLKeyEnumNarrowingTests
{
    [Test]
    public async Task Out_of_range_border_style_throws_rather_than_aliasing_a_defined_member()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var border = ws.Cell("A1").Style.Border;

        // 256 + 2 would wrap to 2, which is XLBorderStyleValues.Dashed - a real, different, defined
        // member. An unchecked cast would apply that style instead of failing.
        var outOfRange = (XLBorderStyleValues)258;

        await Assert.That(() => border.LeftBorder = outOfRange).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Highest_defined_border_style_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var border = ws.Cell("A1").Style.Border;

        // XLBorderStyleValues has 14 members, 0-13; Thin is the highest ordinal and is comfortably
        // inside a byte, and the check must not reject values it should accept.
        border.LeftBorder = XLBorderStyleValues.Thin;

        await Assert.That(border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thin);
    }

    [Test]
    public async Task Out_of_range_theme_colour_throws_rather_than_aliasing_a_defined_member()
    {
        // 256 + 1 would wrap to 1, which is XLThemeColor.Text1 - a real, different, defined member.
        var outOfRange = (XLThemeColor)257;

        await Assert.That(() => XLColor.FromTheme(outOfRange)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Highest_defined_theme_colour_does_not_throw()
    {
        // XLThemeColor has 12 members, 0-11.
        var color = XLColor.FromTheme(XLThemeColor.FollowedHyperlink);

        await Assert.That(color.ThemeColor).IsEqualTo(XLThemeColor.FollowedHyperlink);
    }
}
