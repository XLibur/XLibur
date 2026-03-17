using XLibur.Excel.Caching;
using System;

namespace XLibur.Excel;

/// <summary>
/// An immutable style value.
/// </summary>
internal sealed class XLStyleValue : IEquatable<XLStyleValue?>
{
    private static readonly XLRepositoryBase<XLStyleKey, XLStyleValue> Repository = new(key => new XLStyleValue(key));
    private readonly int _hashCode;

    public static XLStyleValue FromKey(ref XLStyleKey key)
    {
        return Repository.GetOrCreate(ref key);
    }

    private static readonly XLStyleKey DefaultKey = new()
    {
        Alignment = XLAlignmentValue.Default.Key,
        Border = XLBorderValue.Default.Key,
        Fill = XLFillValue.Default.Key,
        Font = XLFontValue.Default.Key,
        IncludeQuotePrefix = false,
        NumberFormat = XLNumberFormatValue.Default.Key,
        Protection = XLProtectionValue.Default.Key
    };

    internal static readonly XLStyleValue Default = FromKey(ref DefaultKey);

    private XLStyleValue(XLStyleKey key)
    {
        Key = key;
        var (alignment, border, fill, font, _, numberFormat, protection) = Key;
        Alignment = XLAlignmentValue.FromKey(ref alignment);
        Border = XLBorderValue.FromKey(ref border);
        Fill = XLFillValue.FromKey(ref fill);
        Font = XLFontValue.FromKey(ref font);
        IncludeQuotePrefix = key.IncludeQuotePrefix;
        NumberFormat = XLNumberFormatValue.FromKey(ref numberFormat);
        Protection = XLProtectionValue.FromKey(ref protection);
        _hashCode = -280332839 + Key.GetHashCode();
    }

    internal XLStyleKey Key { get; }

    #region Transition cache

    /// <summary>
    /// Direct-mapped transition cache stored per base style. When many cells undergo the same
    /// style transition (e.g., Default → Bold=true), this cache returns the result immediately
    /// without computing XLStyleKey hash or hitting the ConcurrentDictionary repository.
    /// Each entry is a (hash, result) pair; collisions simply evict.
    /// Thread-safety: benign races at most cause a cache miss, never incorrect results.
    /// </summary>
    private const int TransitionCacheSize = 8;
    private const int TransitionCacheMask = TransitionCacheSize - 1;

    private int[]? _transitionHashes;
    private XLStyleValue?[]? _transitionResults;

    /// <summary>
    /// Look up a cached transition from this style. Returns null on miss.
    /// </summary>
    internal XLStyleValue? GetTransition(int transitionHash)
    {
        var hashes = _transitionHashes;
        if (hashes is null)
            return null;

        var slot = transitionHash & TransitionCacheMask;
        if (hashes[slot] == transitionHash)
            return _transitionResults![slot];

        return null;
    }

    /// <summary>
    /// Store a transition result and return it (for fluent use: value ?? StoreTransition(...)).
    /// </summary>
    internal XLStyleValue StoreTransition(int transitionHash, XLStyleValue result)
    {
        var hashes = _transitionHashes;
        if (hashes is null)
        {
            hashes = new int[TransitionCacheSize];
            _transitionResults = new XLStyleValue?[TransitionCacheSize];
            _transitionHashes = hashes;
        }

        var slot = transitionHash & TransitionCacheMask;
        hashes[slot] = transitionHash;
        _transitionResults![slot] = result;
        return result;
    }

    #endregion Transition cache

    internal XLAlignmentValue Alignment { get; }

    internal XLBorderValue Border { get; }

    internal XLFillValue Fill { get; }

    internal XLFontValue Font { get; }

    internal bool IncludeQuotePrefix { get; }

    internal XLNumberFormatValue NumberFormat { get; }

    internal XLProtectionValue Protection { get; }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLStyleValue);
    }

    public bool Equals(XLStyleValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;

    public static bool operator ==(XLStyleValue? left, XLStyleValue? right)
    {
        if (left is null)
            return right is null;

        return left.Equals(right);
    }

    public static bool operator !=(XLStyleValue? left, XLStyleValue? right)
    {
        return !(left == right);
    }

    /// <summary>
    /// Combine row and column styles into a combined style. This style is used by non-pinged
    /// cells of a worksheet.
    /// </summary>
    internal static XLStyleValue Combine(XLStyleValue sheetStyle, XLStyleValue rowStyle, XLStyleValue colStyle)
    {
        var isRowSame = ReferenceEquals(sheetStyle, rowStyle);
        var isColSame = ReferenceEquals(sheetStyle, colStyle);

        if (isRowSame && isColSame)
            return sheetStyle;

        // At least one style is different, maybe both.
        if (isRowSame)
            return colStyle;

        if (isColSame)
            return rowStyle;

        // Both styles are different from sheet one, merge. If both style components differ,
        // row has a preference because Excel gives it preference. Generally, if there is
        // a row / col style conflict, all cells affected by conflict should be materialized (aka
        // 'pinged') during row/col style modification and have their own style explicitly
        // specified to avoid ambiguity, so we shouldn't really need to rely on this resolution
        var alignment = GetExplicitlySet(sheetStyle.Alignment, rowStyle.Alignment, colStyle.Alignment);
        var border = GetExplicitlySet(sheetStyle.Border, rowStyle.Border, colStyle.Border);
        var fill = GetExplicitlySet(sheetStyle.Fill, rowStyle.Fill, colStyle.Fill);
        var font = GetExplicitlySet(sheetStyle.Font, rowStyle.Font, colStyle.Font);
        var includeQuotePrefix = GetExplicitlySet(sheetStyle.IncludeQuotePrefix, rowStyle.IncludeQuotePrefix,
            colStyle.IncludeQuotePrefix);
        var numberFormat = GetExplicitlySet(sheetStyle.NumberFormat, rowStyle.NumberFormat, colStyle.NumberFormat);
        var protection = GetExplicitlySet(sheetStyle.Protection, rowStyle.Protection, colStyle.Protection);

        var combinedStyleKey = new XLStyleKey
        {
            Alignment = alignment.Key,
            Border = border.Key,
            Fill = fill.Key,
            Font = font.Key,
            IncludeQuotePrefix = includeQuotePrefix,
            NumberFormat = numberFormat.Key,
            Protection = protection.Key,
        };
        return Repository.GetOrCreate(ref combinedStyleKey);

        static T GetExplicitlySet<T>(T sheetComponent, T rowComponent, T colComponent)
            where T : notnull
        {
            // Use reference equal to speed up the process instead of standard equals.
            var rowHasSameComponent = typeof(T).IsClass
                ? ReferenceEquals(sheetComponent, rowComponent)
                : sheetComponent.Equals(rowComponent);
            var colHasSameComponent = typeof(T).IsClass
                ? ReferenceEquals(sheetComponent, colComponent)
                : sheetComponent.Equals(colComponent);

            return rowHasSameComponent switch
            {
                true when colHasSameComponent => sheetComponent, // At least one style is different, maybe both.
                true => colComponent, // If col has the same component as a sheet, we should return row.
                _ => rowComponent // If both are different, the row component should have precedence.
            };
        }
    }

    internal XLStyleValue WithAlignment(Func<XLAlignmentValue, XLAlignmentValue> modify)
    {
        return WithAlignment(modify(Alignment));
    }

    private XLStyleValue WithAlignment(XLAlignmentValue alignment)
    {
        var keyCopy = Key with { Alignment = alignment.Key };
        return FromKey(ref keyCopy);
    }

    internal XLStyleValue WithIncludeQuotePrefix(bool includeQuotePrefix)
    {
        var keyCopy = Key with { IncludeQuotePrefix = includeQuotePrefix };
        return FromKey(ref keyCopy);
    }

    internal XLStyleValue WithNumberFormat(XLNumberFormatValue numberFormat)
    {
        var keyCopy = Key with { NumberFormat = numberFormat.Key };
        return FromKey(ref keyCopy);
    }
}
