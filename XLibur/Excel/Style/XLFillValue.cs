
using System;
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLFillValue : IEquatable<XLFillValue?>
{
    private static readonly XLRepositoryBase<XLFillKey, XLFillValue> Repository = new(key => new XLFillValue(key));

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

    /// <inheritdoc cref="XLBorderValue._hashCode"/>
    private readonly int _hashCode;

    private XLFillValue(XLFillKey key)
    {
        Key = key;
        _hashCode = -280332839 + key.GetHashCode();
        var backgroundColorKey = Key.BackgroundColor;
        var patternColorKey = Key.PatternColor;
        BackgroundColor = XLColor.FromKey(ref backgroundColorKey);
        PatternColor = XLColor.FromKey(ref patternColorKey);
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLFillValue);
    }

    /// <inheritdoc cref="XLBorderValue.Equals(XLBorderValue)"/>
    public bool Equals(XLFillValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;
}
