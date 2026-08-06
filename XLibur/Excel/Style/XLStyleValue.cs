using System;
using System.Threading;
using XLibur.Excel.Caching;

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
    /// Each entry is a (hash, component key, result) triple; slot collisions simply evict.
    /// Thread-safety: benign races at most cause a cache miss, never incorrect results.
    /// </summary>
    /// <remarks>
    /// The component key is held alongside the hash and compared on every hit. Matching on the hash
    /// alone is not sufficient: two distinct component keys sharing a 32-bit hash would return each
    /// other's style, silently giving a cell a format, font or border it was never assigned.
    /// <para>
    /// That is not the remote possibility it looks like. A custom number format pins
    /// <c>NumberFormatId</c> to a constant −1, so its key hash reduces to the hash of the format
    /// string, and a birthday search finds a colliding pair of ordinary-looking format strings among
    /// roughly 2^16 candidates — milliseconds of work. For a workbook applying k distinct custom
    /// formats to default-styled cells the odds run near k²/2^33. <c>XLFontKey</c> carries a font
    /// name and is exposed the same way.
    /// </para>
    /// <para>
    /// Because <see cref="string.GetHashCode()"/> is randomised per process, such a defect does not
    /// reproduce: the same workbook round-trips correctly on most runs and corrupts on others, with
    /// nothing in the file to explain it. The hash stays as the cheap reject; the key comparison is
    /// what makes a hit sound.
    /// </para>
    /// </remarks>
    private const int TransitionCacheSize = 8;
    private const int TransitionCacheMask = TransitionCacheSize - 1;

    private int[]? _transitionHashes;
    private object?[]? _transitionKeys;
    private XLStyleValue?[]? _transitionResults;

    /// <summary>
    /// Look up a cached transition from this style. Returns null on miss.
    /// </summary>
    /// <param name="transitionHash">Hash of <paramref name="componentKey"/>, mixed with a tag identifying its component.</param>
    /// <param name="componentKey">The component key being applied, compared against the cached entry.</param>
    internal XLStyleValue? GetTransition<TKey>(int transitionHash, in TKey componentKey)
        where TKey : struct, IEquatable<TKey>
    {
        var hashes = _transitionHashes;
        if (hashes is null)
            return null;

        var slot = transitionHash & TransitionCacheMask;
        if (hashes[slot] != transitionHash)
            return null;

        // The type test also rejects a hash collision between two different component types, which
        // the tag mixed into the hash makes unlikely but does not prevent.
        if (_transitionKeys![slot] is not TKey cachedKey || !cachedKey.Equals(componentKey))
            return null;

        return _transitionResults![slot];
    }

    /// <summary>
    /// Store a transition result and return it (for fluent use: value ?? StoreTransition(...)).
    /// </summary>
    /// <remarks>
    /// Boxes <paramref name="componentKey"/>. That cost falls only on a miss — one box per distinct
    /// transition per base style, not per cell — so the repeat transitions this cache exists to
    /// serve do not pay it.
    /// </remarks>
    internal XLStyleValue StoreTransition<TKey>(int transitionHash, in TKey componentKey, XLStyleValue result)
        where TKey : struct, IEquatable<TKey>
    {
        var hashes = _transitionHashes;
        if (hashes is null)
        {
            hashes = new int[TransitionCacheSize];
            _transitionKeys = new object?[TransitionCacheSize];
            _transitionResults = new XLStyleValue?[TransitionCacheSize];
            _transitionHashes = hashes;
        }

        var slot = transitionHash & TransitionCacheMask;

        // The result is published before the hash, and the hash is what a lookup gates on, so a
        // concurrent reader either misses or sees a fully written entry.
        _transitionKeys![slot] = componentKey;
        _transitionResults![slot] = result;
        Volatile.Write(ref hashes[slot], transitionHash);
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
