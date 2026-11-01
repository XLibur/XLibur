using System;
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLNumberFormatValue : IEquatable<XLNumberFormatValue?>
{
    private static readonly XLRepositoryBase<XLNumberFormatKey, XLNumberFormatValue> Repository = new(key => new XLNumberFormatValue(key));

    public static XLNumberFormatValue FromKey(ref XLNumberFormatKey key)
    {
        return Repository.GetOrCreate(ref key);
    }

    /// <summary>
    /// <em>General</em> number format.
    /// </summary>
    private static readonly XLNumberFormatKey DefaultKey = new XLNumberFormatKey
    {
        NumberFormatId = 0,
        Format = string.Empty
    };

    internal static readonly XLNumberFormatValue Default = FromKey(ref DefaultKey);

    /// <remarks>
    /// Get-only: <see cref="_hashCode"/> is derived from it once, so it must not be reassigned. The
    /// setter this replaces was never used outside the constructor.
    /// </remarks>
    public XLNumberFormatKey Key { get; }

    /// <summary>
    /// Id of the number format. Every workbook has <see cref="XLConstants.NumberOfBuiltInStyles"/>
    /// built-int formats that start at 0 (<em>General</em> format). The built-int formats are
    /// not explicitly written and might differ depending on culture. Custom number formats
    /// have a valid <see cref="Format"/> and the id is <c>-1</c>.
    /// </summary>
    public int NumberFormatId => Key.NumberFormatId;

    public string Format => Key.Format;

    /// <inheritdoc cref="XLBorderValue._hashCode"/>
    /// <remarks>
    /// A custom format pins <see cref="NumberFormatId"/> to <c>-1</c>, so this key's hash reduces to
    /// the hash of the format string.
    /// </remarks>
    private readonly int _hashCode;

    private XLNumberFormatValue(XLNumberFormatKey key)
    {
        Key = key;
        _hashCode = 1507230172 + key.GetHashCode();
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLNumberFormatValue);
    }

    /// <inheritdoc cref="XLBorderValue.Equals(XLBorderValue)"/>
    public bool Equals(XLNumberFormatValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;

    internal XLNumberFormatValue WithNumberFormatId(int numberFormatId)
    {
        var keyCopy = Key with { NumberFormatId = numberFormatId };
        return FromKey(ref keyCopy);
    }
}
