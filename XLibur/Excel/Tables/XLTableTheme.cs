using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace XLibur.Excel;

public sealed class XLTableTheme
{
    public static readonly XLTableTheme None = new("None");
    public static readonly XLTableTheme TableStyleMedium28 = new("TableStyleMedium28");
    public static readonly XLTableTheme TableStyleMedium27 = new("TableStyleMedium27");
    public static readonly XLTableTheme TableStyleMedium26 = new("TableStyleMedium26");
    public static readonly XLTableTheme TableStyleMedium25 = new("TableStyleMedium25");
    public static readonly XLTableTheme TableStyleMedium24 = new("TableStyleMedium24");
    public static readonly XLTableTheme TableStyleMedium23 = new("TableStyleMedium23");
    public static readonly XLTableTheme TableStyleMedium22 = new("TableStyleMedium22");
    public static readonly XLTableTheme TableStyleMedium21 = new("TableStyleMedium21");
    public static readonly XLTableTheme TableStyleMedium20 = new("TableStyleMedium20");
    public static readonly XLTableTheme TableStyleMedium19 = new("TableStyleMedium19");
    public static readonly XLTableTheme TableStyleMedium18 = new("TableStyleMedium18");
    public static readonly XLTableTheme TableStyleMedium17 = new("TableStyleMedium17");
    public static readonly XLTableTheme TableStyleMedium16 = new("TableStyleMedium16");
    public static readonly XLTableTheme TableStyleMedium15 = new("TableStyleMedium15");
    public static readonly XLTableTheme TableStyleMedium14 = new("TableStyleMedium14");
    public static readonly XLTableTheme TableStyleMedium13 = new("TableStyleMedium13");
    public static readonly XLTableTheme TableStyleMedium12 = new("TableStyleMedium12");
    public static readonly XLTableTheme TableStyleMedium11 = new("TableStyleMedium11");
    public static readonly XLTableTheme TableStyleMedium10 = new("TableStyleMedium10");
    public static readonly XLTableTheme TableStyleMedium9 = new("TableStyleMedium9");
    public static readonly XLTableTheme TableStyleMedium8 = new("TableStyleMedium8");
    public static readonly XLTableTheme TableStyleMedium7 = new("TableStyleMedium7");
    public static readonly XLTableTheme TableStyleMedium6 = new("TableStyleMedium6");
    public static readonly XLTableTheme TableStyleMedium5 = new("TableStyleMedium5");
    public static readonly XLTableTheme TableStyleMedium4 = new("TableStyleMedium4");
    public static readonly XLTableTheme TableStyleMedium3 = new("TableStyleMedium3");
    public static readonly XLTableTheme TableStyleMedium2 = new("TableStyleMedium2");
    public static readonly XLTableTheme TableStyleMedium1 = new("TableStyleMedium1");
    public static readonly XLTableTheme TableStyleLight21 = new("TableStyleLight21");
    public static readonly XLTableTheme TableStyleLight20 = new("TableStyleLight20");
    public static readonly XLTableTheme TableStyleLight19 = new("TableStyleLight19");
    public static readonly XLTableTheme TableStyleLight18 = new("TableStyleLight18");
    public static readonly XLTableTheme TableStyleLight17 = new("TableStyleLight17");
    public static readonly XLTableTheme TableStyleLight16 = new("TableStyleLight16");
    public static readonly XLTableTheme TableStyleLight15 = new("TableStyleLight15");
    public static readonly XLTableTheme TableStyleLight14 = new("TableStyleLight14");
    public static readonly XLTableTheme TableStyleLight13 = new("TableStyleLight13");
    public static readonly XLTableTheme TableStyleLight12 = new("TableStyleLight12");
    public static readonly XLTableTheme TableStyleLight11 = new("TableStyleLight11");
    public static readonly XLTableTheme TableStyleLight10 = new("TableStyleLight10");
    public static readonly XLTableTheme TableStyleLight9 = new("TableStyleLight9");
    public static readonly XLTableTheme TableStyleLight8 = new("TableStyleLight8");
    public static readonly XLTableTheme TableStyleLight7 = new("TableStyleLight7");
    public static readonly XLTableTheme TableStyleLight6 = new("TableStyleLight6");
    public static readonly XLTableTheme TableStyleLight5 = new("TableStyleLight5");
    public static readonly XLTableTheme TableStyleLight4 = new("TableStyleLight4");
    public static readonly XLTableTheme TableStyleLight3 = new("TableStyleLight3");
    public static readonly XLTableTheme TableStyleLight2 = new("TableStyleLight2");
    public static readonly XLTableTheme TableStyleLight1 = new("TableStyleLight1");
    public static readonly XLTableTheme TableStyleDark11 = new("TableStyleDark11");
    public static readonly XLTableTheme TableStyleDark10 = new("TableStyleDark10");
    public static readonly XLTableTheme TableStyleDark9 = new("TableStyleDark9");
    public static readonly XLTableTheme TableStyleDark8 = new("TableStyleDark8");
    public static readonly XLTableTheme TableStyleDark7 = new("TableStyleDark7");
    public static readonly XLTableTheme TableStyleDark6 = new("TableStyleDark6");
    public static readonly XLTableTheme TableStyleDark5 = new("TableStyleDark5");
    public static readonly XLTableTheme TableStyleDark4 = new("TableStyleDark4");
    public static readonly XLTableTheme TableStyleDark3 = new("TableStyleDark3");
    public static readonly XLTableTheme TableStyleDark2 = new("TableStyleDark2");
    public static readonly XLTableTheme TableStyleDark1 = new("TableStyleDark1");

    public string Name { get; }

    public XLTableTheme(string name)
    {
        Name = name;
    }

    private static IEnumerable<XLTableTheme>? allThemes;

    public static IEnumerable<XLTableTheme> GetAllThemes()
    {
        return allThemes ??= typeof(XLTableTheme).GetFields(BindingFlags.Static | BindingFlags.Public)
            .Where(fi => fi.FieldType == typeof(XLTableTheme))
            .Select(fi => (XLTableTheme)fi.GetValue(null)!)
            .ToArray();
    }

    public static XLTableTheme? FromName(string name)
    {
        return GetAllThemes().FirstOrDefault(s => s.Name == name);
    }

    #region Overrides

    public override bool Equals(object? obj)
    {
        var theme = obj as XLTableTheme;
        return theme is not null && Name.Equals(theme.Name);
    }

    public override int GetHashCode()
    {
        return Name.GetHashCode();
    }

    public override string ToString()
    {
        return Name;
    }

    #endregion Overrides
}
