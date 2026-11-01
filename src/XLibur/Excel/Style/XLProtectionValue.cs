using System;
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLProtectionValue : IEquatable<XLProtectionValue?>
{
    private static readonly XLRepositoryBase<XLProtectionKey, XLProtectionValue> Repository = new(key => new XLProtectionValue(key));

    public static XLProtectionValue FromKey(ref XLProtectionKey key)
    {
        return Repository.GetOrCreate(ref key);
    }

    private static readonly XLProtectionKey DefaultKey = new()
    {
        Locked = true,
        Hidden = false
    };

    internal static readonly XLProtectionValue Default = FromKey(ref DefaultKey);

    public XLProtectionKey Key { get; }

    public bool Locked => Key.Locked;

    public bool Hidden => Key.Hidden;

    /// <inheritdoc cref="XLBorderValue._hashCode"/>
    private readonly int _hashCode;

    private XLProtectionValue(XLProtectionKey key)
    {
        Key = key;
        _hashCode = 909014992 + key.GetHashCode();
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLProtectionValue);
    }

    /// <inheritdoc cref="XLBorderValue.Equals(XLBorderValue)"/>
    public bool Equals(XLProtectionValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;
}
