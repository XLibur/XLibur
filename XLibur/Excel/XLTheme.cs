using System;

namespace XLibur.Excel;

internal sealed class XLTheme : IXLTheme
{
    public required XLColor Background1 { get; set; }

    public required XLColor Text1 { get; set; }

    public required XLColor Background2 { get; set; }

    public required XLColor Text2 { get; set; }

    public required XLColor Accent1 { get; set; }

    public required XLColor Accent2 { get; set; }

    public required XLColor Accent3 { get; set; }

    public required XLColor Accent4 { get; set; }

    public required XLColor Accent5 { get; set; }

    public required XLColor Accent6 { get; set; }

    public required XLColor Hyperlink { get; set; }

    public required XLColor FollowedHyperlink { get; set; }

    public XLColor ResolveThemeColor(XLThemeColor themeColor)
    {
        return themeColor switch
        {
            XLThemeColor.Background1 => Background1,
            XLThemeColor.Text1 => Text1,
            XLThemeColor.Background2 => Background2,
            XLThemeColor.Text2 => Text2,
            XLThemeColor.Accent1 => Accent1,
            XLThemeColor.Accent2 => Accent2,
            XLThemeColor.Accent3 => Accent3,
            XLThemeColor.Accent4 => Accent4,
            XLThemeColor.Accent5 => Accent5,
            XLThemeColor.Accent6 => Accent6,
            XLThemeColor.Hyperlink => Hyperlink,
            XLThemeColor.FollowedHyperlink => FollowedHyperlink,
            _ => throw new ArgumentOutOfRangeException(nameof(themeColor), themeColor, "Invalid theme color")
        };
    }
}
