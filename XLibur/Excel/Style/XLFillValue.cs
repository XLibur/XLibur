
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLFillValue
{
    private static readonly XLFillRepository Repository = new(key => new XLFillValue(key));

    public static XLFillValue FromKey(ref XLFillKey key)
    {
        return Repository.GetOrCreate(ref key);
    }

    private static readonly XLFillKey DefaultKey = new()
    {
        BackgroundColor = XLColor.FromIndex(64).Key,
        PatternType = XLFillPatternValues.None,
        PatternColor = XLColor.FromIndex(64).Key
    };

    internal static readonly XLFillValue Default = FromKey(ref DefaultKey);

    public XLFillKey Key { get; }

    public XLColor BackgroundColor { get; private set; }

    public XLColor PatternColor { get; private set; }

    public XLFillPatternValues PatternType => Key.PatternType;

    private XLFillValue(XLFillKey key)
    {
        Key = key;
        var backgroundColorKey = Key.BackgroundColor;
        var patternColorKey = Key.PatternColor;
        BackgroundColor = XLColor.FromKey(ref backgroundColorKey);
        PatternColor = XLColor.FromKey(ref patternColorKey);
    }

    public override bool Equals(object? obj)
    {
        var cached = obj as XLFillValue;
        return cached != null &&
               Key.Equals(cached.Key);
    }

    public override int GetHashCode()
    {
        return -280332839 + Key.GetHashCode();
    }
}
