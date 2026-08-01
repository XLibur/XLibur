namespace XLibur.Excel;

/// <summary>
/// Lightweight <see cref="IXLProtection"/> that accumulates property changes into an <see cref="XLProtectionKey"/>
/// without triggering repository lookups or style-slice writes. Used by <see cref="XLDeferredStyle"/>
/// to support batch style mutations.
/// </summary>
internal sealed class XLDeferredProtection : IXLProtection
{
    private readonly XLDeferredStyle _style;
    internal XLProtectionKey Key;

    internal XLDeferredProtection(XLDeferredStyle style, XLProtectionKey key)
    {
        _style = style;
        Key = key;
    }

    public bool Locked
    {
        get => Key.Locked;
        set => Key = Key with { Locked = value };
    }

    public bool Hidden
    {
        get => Key.Hidden;
        set => Key = Key with { Hidden = value };
    }

    public IXLStyle SetLocked() { Locked = true; return _style; }
    public IXLStyle SetLocked(bool value) { Locked = value; return _style; }
    public IXLStyle SetHidden() { Hidden = true; return _style; }
    public IXLStyle SetHidden(bool value) { Hidden = value; return _style; }

    public bool Equals(IXLProtection? other) => other is XLDeferredProtection dp && Key == dp.Key;
}
