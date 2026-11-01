using System;
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLAlignmentValue : IEquatable<XLAlignmentValue?>
{
    private static readonly XLRepositoryBase<XLAlignmentKey, XLAlignmentValue> Repository = new(key => new XLAlignmentValue(key));

    public static XLAlignmentValue FromKey(ref XLAlignmentKey key)
    {
        return Repository.GetOrCreate(ref key);
    }

    private static readonly XLAlignmentKey DefaultKey = new()
    {
        Indent = 0,
        Horizontal = XLAlignmentHorizontalValues.General,
        JustifyLastLine = false,
        ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent,
        RelativeIndent = 0,
        ShrinkToFit = false,
        TextRotation = 0,
        Vertical = XLAlignmentVerticalValues.Bottom,
        WrapText = false
    };

    internal static readonly XLAlignmentValue Default = FromKey(ref DefaultKey);

    public XLAlignmentKey Key { get; }

    public XLAlignmentHorizontalValues Horizontal => Key.Horizontal;

    public XLAlignmentVerticalValues Vertical => Key.Vertical;

    public int Indent => Key.Indent;

    public bool JustifyLastLine => Key.JustifyLastLine;

    public XLAlignmentReadingOrderValues ReadingOrder => Key.ReadingOrder;

    public int RelativeIndent => Key.RelativeIndent;

    public bool ShrinkToFit => Key.ShrinkToFit;

    public int TextRotation => Key.TextRotation;

    public bool WrapText => Key.WrapText;

    /// <inheritdoc cref="XLBorderValue._hashCode"/>
    private readonly int _hashCode;

    private XLAlignmentValue(XLAlignmentKey key)
    {
        Key = key;
        _hashCode = 990326508 + key.GetHashCode();
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLAlignmentValue);
    }

    /// <inheritdoc cref="XLBorderValue.Equals(XLBorderValue)"/>
    public bool Equals(XLAlignmentValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;

    internal XLAlignmentValue WithWrapText(bool wrapText)
    {
        var keyCopy = Key with { WrapText = wrapText };
        return FromKey(ref keyCopy);
    }
}
