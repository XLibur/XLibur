using System;

namespace XLibur.Excel;

internal sealed class XLProtection : IXLProtection
{
    #region Static members

    internal static XLProtectionKey GenerateKey(IXLProtection? defaultProtection) => defaultProtection switch
    {
        null => XLProtectionValue.Default.Key,
        XLProtection protection => protection.Key,
        _ => new XLProtectionKey
        {
            Locked = defaultProtection.Locked,
            Hidden = defaultProtection.Hidden
        },
    };

    #endregion Static members

    #region Properties

    private readonly XLStyle _style;

    private XLProtectionValue _value;

    internal XLProtectionKey Key
    {
        get => _value.Key;
        private set => _value = XLProtectionValue.FromKey(ref value);
    }

    #endregion Properties

    #region Constructors

    /// <summary>
    /// Create an instance of XLProtection initializing it with the specified value.
    /// </summary>
    /// <param name="style">Style to attach the new instance to.</param>
    /// <param name="value">Style value to use.</param>
    public XLProtection(XLStyle? style, XLProtectionValue value)
    {
        _style = style ?? XLStyle.CreateEmptyStyle();
        _value = value;
    }

    public XLProtection(XLStyle? style, XLProtectionKey key) : this(style, XLProtectionValue.FromKey(ref key))
    {
    }

    public XLProtection(XLStyle? style = null, IXLProtection? d = null) : this(style, GenerateKey(d))
    {
    }

    #endregion Constructors

    internal void SyncValue(XLProtectionValue value) { _value = value; }

    #region IXLProtection Members

    public bool Locked
    {
        get => Key.Locked;
        set
        {
            if (Key.Locked == value) return;
            if (_style.IsCellContainer)
                SetKey(Key with { Locked = value });
            else
                Modify(k => k with { Locked = value });
        }
    }

    public bool Hidden
    {
        get => Key.Hidden;
        set
        {
            if (Key.Hidden == value) return;
            if (_style.IsCellContainer)
                SetKey(Key with { Hidden = value });
            else
                Modify(k => k with { Hidden = value });
        }
    }

    public IXLStyle SetLocked()
    {
        Locked = true;
        return _style;
    }

    public IXLStyle SetLocked(bool value)
    {
        Locked = value;
        return _style;
    }

    public IXLStyle SetHidden()
    {
        Hidden = true;
        return _style;
    }

    public IXLStyle SetHidden(bool value)
    {
        Hidden = value;
        return _style;
    }

    #endregion IXLProtection Members

    /// <summary>
    /// Apply a new component key to the cell this facade is attached to.
    /// </summary>
    /// <remarks>
    /// The new key is deliberately <em>not</em> interned before being applied. Assigning it to
    /// <c>Key</c> first would run a repository lookup -- hashing the key and probing a dictionary --
    /// whose result is then thrown away: on a transition-cache hit <c>ModifyProtection</c> never needs the
    /// component value at all, and on a miss it interns the component anyway inside
    /// <c>XLStyleValue.FromKey</c>. Taking the interned value back off the resulting style instead
    /// leaves the facade just as correct for later reads, at no lookup. Measured over 20,000 cells
    /// setting one property each, this was the single largest cost on the per-cell styling path.
    /// </remarks>
    private void SetKey(XLProtectionKey newKey)
    {
        _style.ModifyProtection(newKey);
        _value = _style.Value.Protection;
    }

    private void Modify(Func<XLProtectionKey, XLProtectionKey> modification)
    {
        Key = modification(Key);
        _style.Modify(styleKey => styleKey with { Protection = modification(styleKey.Protection) });
    }

    #region Overridden

    public override bool Equals(object? obj)
    {
        return obj is IXLProtection protection && Equals(protection);
    }

    public bool Equals(IXLProtection? other)
    {
        var otherP = other as XLProtection;
        if (otherP == null)
            return false;

        return Key == otherP.Key;
    }

    public override string ToString()
    {
        if (Locked)
            return Hidden ? "Locked-Hidden" : "Locked";

        return Hidden ? "Hidden" : "None";
    }

    public override int GetHashCode()
    {
        var hashCode = 416600561;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden
}
