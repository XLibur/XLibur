using System;
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLBorderValue : IEquatable<XLBorderValue?>
{
    private static readonly XLRepositoryBase<XLBorderKey, XLBorderValue> Repository = new(key => new XLBorderValue(key));

    public static XLBorderValue FromKey(ref XLBorderKey key)
    {
        // Normalizing here rather than at the call sites is what lets XLBorderKey.Equals compare
        // fields directly: every key the repository ever sees has been through it.
        var normalized = key.Normalize();
        return Repository.GetOrCreate(ref normalized);
    }

    private static readonly XLBorderKey DefaultKey = new()
    {
        BottomBorder = XLBorderStyleValues.None,
        DiagonalBorder = XLBorderStyleValues.None,
        DiagonalDown = false,
        DiagonalUp = false,
        LeftBorder = XLBorderStyleValues.None,
        RightBorder = XLBorderStyleValues.None,
        TopBorder = XLBorderStyleValues.None,
        BottomBorderColor = XLColor.Black.Key,
        DiagonalBorderColor = XLColor.Black.Key,
        LeftBorderColor = XLColor.Black.Key,
        RightBorderColor = XLColor.Black.Key,
        TopBorderColor = XLColor.Black.Key
    };

    internal static readonly XLBorderValue Default = FromKey(ref DefaultKey);

    public XLBorderKey Key { get; }

    public XLBorderStyleValues LeftBorder => Key.LeftBorder;

    public XLColor LeftBorderColor { get; private set; }

    public XLBorderStyleValues RightBorder => Key.RightBorder;

    public XLColor RightBorderColor { get; private set; }

    public XLBorderStyleValues TopBorder => Key.TopBorder;

    public XLColor TopBorderColor { get; private set; }

    public XLBorderStyleValues BottomBorder => Key.BottomBorder;

    public XLColor BottomBorderColor { get; private set; }

    public XLBorderStyleValues DiagonalBorder => Key.DiagonalBorder;

    public XLColor DiagonalBorderColor { get; private set; }

    public bool DiagonalUp => Key.DiagonalUp;

    public bool DiagonalDown => Key.DiagonalDown;

    /// <summary>
    /// The hash of <see cref="Key"/>, computed once here rather than on every lookup.
    /// </summary>
    /// <remarks>
    /// Values are used directly as dictionary keys when the styles part is written
    /// (<c>Dictionary&lt;XLBorderValue, BorderInfo&gt;</c>), and the border key's hash folds five
    /// colour hashes. It cannot change: the key is immutable and interned against it.
    /// </remarks>
    private readonly int _hashCode;

    private XLBorderValue(XLBorderKey key)
    {
        Key = key;
        _hashCode = -280332839 + key.GetHashCode();
        var leftBorderColor = Key.LeftBorderColor;
        var rightBorderColor = Key.RightBorderColor;
        var topBorderColor = Key.TopBorderColor;
        var bottomBorderColor = Key.BottomBorderColor;
        var diagonalBorderColor = Key.DiagonalBorderColor;
        LeftBorderColor = XLColor.FromKey(ref leftBorderColor);
        RightBorderColor = XLColor.FromKey(ref rightBorderColor);
        TopBorderColor = XLColor.FromKey(ref topBorderColor);
        BottomBorderColor = XLColor.FromKey(ref bottomBorderColor);
        DiagonalBorderColor = XLColor.FromKey(ref diagonalBorderColor);
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLBorderValue);
    }

    /// <summary>
    /// Reference equality first, because values are interned one per key: two references to the same
    /// border are overwhelmingly the same instance, and the comparison ends without touching the key.
    /// Structural equality is still needed for the rest - the repository holds weak references, so a
    /// collected entry can be rebuilt as a second instance for a key still in use elsewhere.
    /// </summary>
    public bool Equals(XLBorderValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;
}
