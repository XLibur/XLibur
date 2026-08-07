using System.Drawing;

namespace XLibur.Excel;

public sealed partial class XLColor
{
    internal XLColorKey Key { get; private set; }

    private XLColor(XLColor defaultColor) : this(defaultColor.Key)
    {
    }

    /// <summary>
    /// The automatic color. <see cref="XLColorType.Automatic"/> is the first enum member, so a
    /// default <see cref="XLColorKey"/> already describes itself as automatic - no sentinel value
    /// smuggled into an RGB component.
    /// </summary>
    private XLColor() : this(XLColorKey.Automatic)
    {
        HasValue = false;
    }

    private XLColor(Color color) : this(XLColorKey.FromColor(color))
    {
    }

    private XLColor(int index) : this(XLColorKey.FromIndex(index))
    {
    }

    private XLColor(XLThemeColor themeColor) : this(XLColorKey.FromTheme(themeColor))
    {
    }

    private XLColor(XLThemeColor themeColor, double themeTint) : this(XLColorKey.FromTheme(themeColor, themeTint))
    {
    }

    internal XLColor(XLColorKey key)
    {
        Key = key;
        HasValue = true;
    }
}
