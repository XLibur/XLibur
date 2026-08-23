using System;
using System.Text;

namespace XLibur.Excel;

internal sealed class XLStyle : IXLStyle
{
    #region Static members

    public static XLStyle Default => new(XLStyleValue.Default);

    internal static XLStyleKey GenerateKey(IXLStyle? initialStyle) => initialStyle switch
    {
        null => Default.Key,
        XLStyle style => style.Key,
        _ => new XLStyleKey
        {
            Alignment = XLAlignment.GenerateKey(initialStyle.Alignment),
            Border = XLBorder.GenerateKey(initialStyle.Border),
            Fill = XLFill.GenerateKey(initialStyle.Fill),
            Font = XLFont.GenerateKey(initialStyle.Font),
            IncludeQuotePrefix = initialStyle.IncludeQuotePrefix,
            NumberFormat = XLNumberFormat.GenerateKey(initialStyle.NumberFormat),
            Protection = XLProtection.GenerateKey(initialStyle.Protection)
        },
    };

    internal static XLStyle CreateEmptyStyle()
    {
        return new XLStyle(new XLStylizedEmpty(null));
    }

    #endregion Static members

    #region properties

    private readonly IXLStylized? _container;

    internal XLStyleValue Value { get; private set; }

    /// <remarks>
    /// Reads through <see cref="_pending"/> while a <see cref="Batch"/> is accumulating, so every
    /// reader of the whole key - <c>RestoreOutsideBorder</c>, <see cref="Equals(IXLStyle)"/> - sees
    /// what has been assigned so far in the batch rather than the pre-batch value. Assembling the
    /// key costs six component hashes, so the component facades read their own slice through
    /// <see cref="CurrentFontKey"/> and its siblings instead of coming through here. The setter is
    /// only reachable outside a batch: <see cref="Modify"/> writes the pending components directly.
    /// </remarks>
    internal XLStyleKey Key
    {
        get => _pending is null ? Value.Key : _pending.ToKey();
        private set => Value = XLStyleValue.FromKey(ref value);
    }

    /// <summary>
    /// Non-null while <see cref="Batch"/> is accumulating. The component fast paths write their new
    /// component key into it instead of resolving a style value and pushing it to the cell, so a
    /// batch of N property assignments costs one resolution rather than N.
    /// </summary>
    private PendingKey? _pending;

    /// <summary>True while a batch is accumulating.</summary>
    internal bool IsBatching => _pending is not null;

    internal XLBorderKey CurrentBorderKey => _pending is null ? Value.Border.Key : _pending.Border;
    internal XLFontKey CurrentFontKey => _pending is null ? Value.Font.Key : _pending.Font;
    internal XLFillKey CurrentFillKey => _pending is null ? Value.Fill.Key : _pending.Fill;
    internal XLAlignmentKey CurrentAlignmentKey => _pending is null ? Value.Alignment.Key : _pending.Alignment;
    internal XLNumberFormatKey CurrentNumberFormatKey => _pending is null ? Value.NumberFormat.Key : _pending.NumberFormat;
    internal XLProtectionKey CurrentProtectionKey => _pending is null ? Value.Protection.Key : _pending.Protection;

    #endregion properties

    #region constructors

    public XLStyle(IXLStylized container, IXLStyle? initialStyle = null, bool useDefaultModify = true) : this(container,
        GenerateKey(initialStyle))
    {
    }

    public XLStyle(IXLStylized container, XLStyleKey key) : this(container, XLStyleValue.FromKey(ref key))
    {
    }

    internal XLStyle(IXLStylized container, XLStyleValue value)
    {
        _container = container ?? new XLStylizedEmpty(Default);
        Value = value;
    }

    /// <summary>
    /// To initialize XLStyle.Default only
    /// </summary>
    private XLStyle(XLStyleValue value)
    {
        Value = value;
    }

    #endregion constructors

    internal void Modify(Func<XLStyleKey, XLStyleKey> modification)
    {
        if (_pending is not null)
        {
            _pending.SetFrom(modification(_pending.ToKey()));
            return;
        }

        Key = modification(Key);

        if (_container != null)
        {
            _container.ModifyStyle(modification);
        }
    }

    /// <summary>
    /// Fast-path style modification for XLCell containers. Only called when <see cref="IsCellContainer"/> is true.
    /// Bypasses closure allocation by directly computing the new style key.
    /// Uses per-base-style transition cache to skip full key hash + repository lookup on repeat transitions.
    /// <para>
    /// While a <see cref="Batch"/> is accumulating the new component key goes into the pending key
    /// instead: no repository lookup, no transition-cache probe, no style-slice write. The batch
    /// resolves once, at flush.
    /// </para>
    /// </summary>
    internal void ModifyFont(XLFontKey newFontKey)
    {
        if (_pending is not null)
        {
            _pending.Font = newFontKey;
            return;
        }

        var transitionHash = (newFontKey.GetHashCode() * 397) ^ 0;
        Value = Value.GetTransition(transitionHash, in newFontKey)
                ?? Value.StoreTransition(transitionHash, in newFontKey, ResolveFont(newFontKey));
        ((XLCell)_container!).SetStyleValue(Value);

        XLStyleValue ResolveFont(XLFontKey key)
        {
            var styleKey = Key with { Font = key };
            return XLStyleValue.FromKey(ref styleKey);
        }
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyBorder(XLBorderKey newBorderKey)
    {
        // Normalized up front so the transition cache keys on the same form the repositories hold;
        // otherwise two keys differing only in a styleless edge's colour would occupy two slots and
        // resolve to the same style anyway.
        newBorderKey = newBorderKey.Normalize();

        if (_pending is not null)
        {
            _pending.Border = newBorderKey;
            return;
        }

        // Tag the hash so the same component key applied to different components lands in a
        // different slot. This only spreads the entries out; correctness comes from the key
        // comparison inside GetTransition, which also rejects a cross-component hash collision.
        var transitionHash = (newBorderKey.GetHashCode() * 397) ^ 1;
        Value = Value.GetTransition(transitionHash, in newBorderKey)
                ?? Value.StoreTransition(transitionHash, in newBorderKey, ResolveBorder(newBorderKey));
        ((XLCell)_container!).SetStyleValue(Value);
        return;

        XLStyleValue ResolveBorder(XLBorderKey key)
        {
            var styleKey = Key with { Border = key };
            return XLStyleValue.FromKey(ref styleKey);
        }
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyFill(XLFillKey newFillKey)
    {
        if (_pending is not null)
        {
            _pending.Fill = newFillKey;
            return;
        }

        var transitionHash = (newFillKey.GetHashCode() * 397) ^ 2;
        Value = Value.GetTransition(transitionHash, in newFillKey)
                ?? Value.StoreTransition(transitionHash, in newFillKey, ResolveFill(newFillKey));
        ((XLCell)_container!).SetStyleValue(Value);
        return;

        XLStyleValue ResolveFill(XLFillKey key)
        {
            var styleKey = Key with { Fill = key };
            return XLStyleValue.FromKey(ref styleKey);
        }
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyAlignment(XLAlignmentKey newAlignmentKey)
    {
        if (_pending is not null)
        {
            _pending.Alignment = newAlignmentKey;
            return;
        }

        var transitionHash = (newAlignmentKey.GetHashCode() * 397) ^ 3;
        Value = Value.GetTransition(transitionHash, in newAlignmentKey)
                ?? Value.StoreTransition(transitionHash, in newAlignmentKey, ResolveAlignment(newAlignmentKey));
        ((XLCell)_container!).SetStyleValue(Value);
        return;

        XLStyleValue ResolveAlignment(XLAlignmentKey key)
        {
            var styleKey = Key with { Alignment = key };
            return XLStyleValue.FromKey(ref styleKey);
        }
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyNumberFormat(XLNumberFormatKey newNumberFormatKey)
    {
        if (_pending is not null)
        {
            _pending.NumberFormat = newNumberFormatKey;
            return;
        }

        var transitionHash = (newNumberFormatKey.GetHashCode() * 397) ^ 4;
        Value = Value.GetTransition(transitionHash, in newNumberFormatKey)
                ?? Value.StoreTransition(transitionHash, in newNumberFormatKey, ResolveNumberFormat(newNumberFormatKey));
        ((XLCell)_container!).SetStyleValue(Value);
        return;

        XLStyleValue ResolveNumberFormat(XLNumberFormatKey key)
        {
            var styleKey = Key with { NumberFormat = key };
            return XLStyleValue.FromKey(ref styleKey);
        }
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyProtection(XLProtectionKey newProtectionKey)
    {
        if (_pending is not null)
        {
            _pending.Protection = newProtectionKey;
            return;
        }

        var transitionHash = (newProtectionKey.GetHashCode() * 397) ^ 5;
        Value = Value.GetTransition(transitionHash, in newProtectionKey)
                ?? Value.StoreTransition(transitionHash, in newProtectionKey, ResolveProtection(newProtectionKey));
        ((XLCell)_container!).SetStyleValue(Value);
        return;

        XLStyleValue ResolveProtection(XLProtectionKey key)
        {
            var styleKey = Key with { Protection = key };
            return XLStyleValue.FromKey(ref styleKey);
        }
    }

    /// <summary>
    /// Apply multiple style changes as a single operation. For cell containers, only one repository
    /// lookup and one style-slice write occurs regardless of how many key fields change.
    /// </summary>
    internal void BatchModify(Func<XLStyleKey, XLStyleKey> modify)
    {
        var currentKey = Value.Key;
        var newKey = modify(currentKey);
        if (currentKey.Equals(newKey))
            return;

        Value = XLStyleValue.FromKey(ref newKey);
        if (_container is XLCell cell)
            cell.SetStyleValue(Value);
        else
            _container?.ModifyStyle(_ => newKey);
    }

    /// <inheritdoc/>
    public IXLStyle Batch(Action<IXLStyle> modifications)
    {
        if (!IsCellContainer || _pending is not null)
        {
            // For ranges: fall back to normal behavior (each property triggers ModifyStyle).
            // For a batch nested inside a batch: the outer one is already accumulating, and
            // restarting it here would discard whatever it holds and flush at the inner close.
            modifications(this);
            return this;
        }

        // For cells: accumulate into a pending key and resolve once. The facades are the ordinary
        // ones, so container-aware operations - a compound border edit, say - behave exactly as
        // they do outside a batch.
        var pending = PendingKey.Rent(Value);
        _pending = pending;
        XLStyleKey newKey;
        try
        {
            modifications(this);
        }
        finally
        {
            newKey = pending.ToKey();
            _pending = null;
            PendingKey.Return(pending);
        }

        if (!Value.Key.Equals(newKey))
        {
            Value = XLStyleValue.FromKey(ref newKey);
            ((XLCell)_container!).SetStyleValue(Value);
        }

        return this;
    }

    internal void SyncValue(XLStyleValue value)
    {
        Value = value;
    }

    /// <summary>
    /// True when the container is an XLCell, allowing fast-path style modifications without closures.
    /// </summary>
    internal bool IsCellContainer => _container is XLCell;

    #region Cached sub-wrappers

    private XLFont? _cachedFont;
    private XLAlignment? _cachedAlignment;
    private XLBorder? _cachedBorder;
    private XLFill? _cachedFill;
    private XLNumberFormat? _cachedNumberFormat;
    private XLProtection? _cachedProtection;

    #endregion Cached sub-wrappers

    #region IXLStyle members

    public IXLFont Font
    {
        get
        {
            if (_cachedFont == null)
                _cachedFont = new XLFont(this, Value.Font);
            else
                _cachedFont.SyncValue(Value.Font);
            return _cachedFont;
        }
        set { Modify(k => k with { Font = XLFont.GenerateKey(value) }); }
    }

    public IXLAlignment Alignment
    {
        get
        {
            if (_cachedAlignment == null)
                _cachedAlignment = new XLAlignment(this, Value.Alignment);
            else
                _cachedAlignment.SyncValue(Value.Alignment);
            return _cachedAlignment;
        }
        set { Modify(k => k with { Alignment = XLAlignment.GenerateKey(value) }); }
    }

    public IXLBorder Border
    {
        get
        {
            if (_cachedBorder == null)
                _cachedBorder = new XLBorder(_container!, this, Value.Border);
            else
                _cachedBorder.SyncValue(Value.Border);
            return _cachedBorder;
        }
        set { Modify(k => k with { Border = XLBorder.GenerateKey(value) }); }
    }

    public IXLFill Fill
    {
        get
        {
            if (_cachedFill == null)
                _cachedFill = new XLFill(this, Value.Fill);
            else
                _cachedFill.SyncValue(Value.Fill);
            return _cachedFill;
        }
        set { Modify(k => k with { Fill = XLFill.GenerateKey(value) }); }
    }

    public bool IncludeQuotePrefix
    {
        get => _pending is null ? Value.IncludeQuotePrefix : _pending.IncludeQuotePrefix;
        set { Modify(k => k with { IncludeQuotePrefix = value }); }
    }

    public IXLStyle SetIncludeQuotePrefix(bool includeQuotePrefix = true)
    {
        IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    public IXLNumberFormat NumberFormat
    {
        get
        {
            if (_cachedNumberFormat == null)
                _cachedNumberFormat = new XLNumberFormat(this, Value.NumberFormat);
            else
                _cachedNumberFormat.SyncValue(Value.NumberFormat);
            return _cachedNumberFormat;
        }
        set { Modify(k => k with { NumberFormat = XLNumberFormat.GenerateKey(value) }); }
    }

    public IXLProtection Protection
    {
        get
        {
            if (_cachedProtection == null)
                _cachedProtection = new XLProtection(this, Value.Protection);
            else
                _cachedProtection.SyncValue(Value.Protection);
            return _cachedProtection;
        }
        set { Modify(k => k with { Protection = XLProtection.GenerateKey(value) }); }
    }

    public IXLNumberFormat DateFormat => NumberFormat;

    #endregion IXLStyle members

    #region Overridden

    public override string ToString()
    {
        var sb = new StringBuilder();
        sb.Append("Font:");
        sb.Append(Font);
        sb.Append(" Fill:");
        sb.Append(Fill);
        sb.Append(" Border:");
        sb.Append(Border);
        sb.Append(" NumberFormat: ");
        sb.Append(NumberFormat);
        sb.Append(" Alignment: ");
        sb.Append(Alignment);
        sb.Append(" Protection: ");
        sb.Append(Protection);
        return sb.ToString();
    }

    public bool Equals(IXLStyle? other)
    {
        var otherS = other as XLStyle;

        if (otherS == null)
            return false;

        return Key == otherS.Key;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as XLStyle);
    }

    public override int GetHashCode()
    {
        var hashCode = 416600561;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden

    #region Nested classes

    /// <summary>
    /// The pending state of an accumulating <see cref="Batch"/>, held one component key at a time.
    /// </summary>
    /// <remarks>
    /// Deliberately not an <c>XLStyleKey?</c> field. That key is a large struct whose every
    /// <c>with</c> expression copies the whole thing and re-runs the assigned component's <c>init</c>
    /// accessor, which normalizes and re-hashes it - so a six-property batch paid six copies and six
    /// component hashes it would pay again when the key was finally resolved. Held inline it also
    /// grew every <c>XLStyle</c> by the size of the key, and an <c>XLStyle</c> is allocated per cell
    /// on the ordinary styling path, which made <em>unbatched</em> styling measurably slower and
    /// allocate ~13 MB more over 50,000 cells. Loose component fields behind one reference cost the
    /// style 8 bytes and assemble the key exactly once, at flush.
    /// <para>
    /// Rented from a one-deep per-thread cache rather than allocated per <c>Batch</c>: it is ~250
    /// bytes, it never outlives the call that rents it, and styling a sheet opens one batch per
    /// cell - allocating it outright put 12.6 MB and a run of gen1 collections on a 50,000-cell
    /// batch that the object graph it replaced did not pay.
    /// </para>
    /// </remarks>
    private sealed class PendingKey
    {
        /// <remarks>
        /// One deep, and empty while its instance is out on loan. A batch nested inside another
        /// batch on the same thread therefore finds it empty and allocates, which is correct: the
        /// outer batch still owns the cached instance and is still writing to it.
        /// </remarks>
        [ThreadStatic]
        private static PendingKey? _cached;

        internal XLFontKey Font;
        internal XLFillKey Fill;
        internal XLBorderKey Border;
        internal XLAlignmentKey Alignment;
        internal XLNumberFormatKey NumberFormat;
        internal XLProtectionKey Protection;
        internal bool IncludeQuotePrefix;

        internal static PendingKey Rent(XLStyleValue value)
        {
            var pending = _cached;
            if (pending is null)
            {
                pending = new PendingKey();
            }
            else
            {
                _cached = null;
            }

            pending.SeedFrom(value);
            return pending;
        }

        internal static void Return(PendingKey pending) => _cached = pending;

        /// <summary>
        /// Seeded from the resolved value's interned components rather than from its
        /// <see cref="XLStyleValue.Key"/>, so opening a batch copies no large struct.
        /// </summary>
        private void SeedFrom(XLStyleValue value)
        {
            Font = value.Font.Key;
            Fill = value.Fill.Key;
            Border = value.Border.Key;
            Alignment = value.Alignment.Key;
            NumberFormat = value.NumberFormat.Key;
            Protection = value.Protection.Key;
            IncludeQuotePrefix = value.IncludeQuotePrefix;
        }

        internal XLStyleKey ToKey() => new()
        {
            Font = Font,
            Fill = Fill,
            Border = Border,
            Alignment = Alignment,
            NumberFormat = NumberFormat,
            Protection = Protection,
            IncludeQuotePrefix = IncludeQuotePrefix,
        };

        internal void SetFrom(XLStyleKey key)
        {
            Font = key.Font;
            Fill = key.Fill;
            Border = key.Border;
            Alignment = key.Alignment;
            NumberFormat = key.NumberFormat;
            Protection = key.Protection;
            IncludeQuotePrefix = key.IncludeQuotePrefix;
        }
    }

    #endregion Nested classes
}
