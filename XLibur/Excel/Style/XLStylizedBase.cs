using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Base class for any workbook element that has or may have a style.
/// </summary>
internal abstract class XLStylizedBase : IXLStylized
{
    #region Properties

    /// <summary>
    /// Read-only style property.
    /// </summary>
    internal virtual XLStyleValue StyleValue { get; private protected set; } = null!;

    /// <inheritdoc cref="IXLStylized.StyleValue"/>
    XLStyleValue IXLStylized.StyleValue => StyleValue;

    /// <inheritdoc cref="IXLStylized.Style"/>
    public IXLStyle Style
    {
        get => InnerStyle;
        set => SetStyle(value, true);
    }

    private XLStyle? _cachedStyle;

    /// <inheritdoc cref="IXLStylized.InnerStyle"/>
    public IXLStyle InnerStyle
    {
        get
        {
            if (_cachedStyle == null)
                _cachedStyle = new XLStyle(this, StyleValue);
            else
                _cachedStyle.SyncValue(StyleValue);
            return _cachedStyle;
        }
        set => SetStyle(value);
    }

    /// <summary>
    /// Get a collection of stylized entities which current entity's style changes should be propagated to.
    /// </summary>
    protected abstract IEnumerable<XLStylizedBase> Children { get; }

    public abstract IXLRanges RangesUsed { get; }

    #endregion Properties

    protected XLStylizedBase(XLStyleValue? styleValue)
    {
        StyleValue = styleValue ?? XLWorkbook.DefaultStyleValue;
    }

    /// <summary>
    /// Ctor only for XLCell that stores <see cref="StyleValue"/> in a slice.
    /// Do not set StyleValue here — XLCell overrides the property with a virtual
    /// setter that requires fields not yet initialized by the derived constructor.
    /// The backing field is initialized to <c>null!</c> via the property initializer.
    /// </summary>
    protected XLStylizedBase()
    {
    }

    #region Private methods

    private void SetStyle(IXLStyle style, bool propagate = false)
    {
        if (style is XLStyle xlStyle)
            SetStyle(xlStyle.Value, propagate);
        else
        {
            var styleKey = XLStyle.GenerateKey(style);
            SetStyle(XLStyleValue.FromKey(ref styleKey), propagate);
        }
    }

    /// <summary>
    /// Apply specified style to the container.
    /// </summary>
    /// <param name="value">Style to apply.</param>
    /// <param name="propagate">Whether or not propagate the style to inner ranges.</param>
    private void SetStyle(XLStyleValue value, bool propagate = false)
    {
        StyleValue = value;
        if (!propagate)
            return;

        // The value is absolute, so the container's own style having already been written cannot
        // affect what the cells resolve to — the fast path is free to run afterwards.
        if (TryApplyToCellStyles(_ => value))
            return;

        Children.ForEach(child => child.SetStyle(StyleValue, true));
    }

    private static readonly ReferenceEqualityComparer<XLStyleValue> Comparer = new();

    void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
    {
        // Snapshot before touching anything. The slow path collects every child up front and
        // groups by the *original* style, so no child may observe another's new style; a row or
        // column writing its own style early would change what its unstyled cells inherit.
        var ownSource = StyleValue;

        if (TryApplyToCellStyles(source => Modify(source, modification)))
        {
            StyleValue = Modify(ownSource, modification);
            return;
        }

        var children = GetChildrenRecursively(this)
            .GroupBy(child => child.StyleValue, Comparer);

        foreach (var group in children)
        {
            var styleValue = Modify(group.Key, modification);
            foreach (var child in group)
            {
                child.StyleValue = styleValue;
            }
        }
    }

    private static XLStyleValue Modify(XLStyleValue source, Func<XLStyleKey, XLStyleKey> modification)
    {
        var styleKey = modification(source.Key);
        return XLStyleValue.FromKey(ref styleKey);
    }

    /// <summary>
    /// Apply <paramref name="transform"/> to the style of every cell this container styles, writing
    /// the style slice directly instead of materialising an <see cref="XLCell"/> per cell.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the container handled its cells; <c>false</c> to fall back to walking
    /// <see cref="Children"/>. The default is <c>false</c> — a container only opts in when it can
    /// enumerate exactly the same cells its <see cref="Children"/> would have yielded.
    /// </returns>
    /// <remarks>
    /// Implementations must not write the container's own style: <c>ModifyStyle</c> applies that
    /// afterwards, from a value snapshotted before any cell was touched.
    /// </remarks>
    private protected virtual bool TryApplyToCellStyles(Func<XLStyleValue, XLStyleValue> transform)
    {
        return false;
    }

    /// <summary>
    /// Shared body of the fast path: walk points, resolve each cell's current effective style, and
    /// write the transformed value back. A last-value memo keeps the transform off runs of cells
    /// that share a style, which is the normal case for bulk styling.
    /// </summary>
    private protected static void ApplyToCellStyles(
        XLWorksheet worksheet,
        IEnumerable<XLSheetPoint> points,
        Func<XLStyleValue, XLStyleValue> transform)
    {
        var styleSlice = worksheet.Internals.CellsCollection.StyleSlice;
        XLStyleValue? memoSource = null;
        XLStyleValue? memoResult = null;

        foreach (var point in points)
        {
            var source = worksheet.GetStyleValue(point);
            if (!ReferenceEquals(source, memoSource))
            {
                memoSource = source;
                memoResult = transform(source);
            }

            styleSlice.Set(point, memoResult!);
        }
    }

    /// <summary>
    /// The points of cells that are actually used within <paramref name="range"/>, matching what
    /// <c>XLCellsCollection.GetCells</c> would have yielded — without the wrapper per cell.
    /// </summary>
    /// <remarks>
    /// Materialised into a list rather than streamed: the caller writes the style slice as it goes,
    /// and the slice enumerator must not be walked while the slice it reads is being mutated.
    /// </remarks>
    private protected static List<XLSheetPoint> UsedPoints(XLWorksheet worksheet, XLSheetRange range)
    {
        var points = new List<XLSheetPoint>();
        var enumerator = new XLCellsCollection.SlicesEnumerator(range, worksheet.Internals.CellsCollection);
        while (enumerator.MoveNext())
            points.Add(enumerator.Current);

        return points;
    }

    private static HashSet<XLStylizedBase> GetChildrenRecursively(XLStylizedBase parent)
    {
        var results = new HashSet<XLStylizedBase>();
        Collect(parent, results);

        return results;

        void Collect(XLStylizedBase root, HashSet<XLStylizedBase> collector)
        {
            collector.Add(root);
            foreach (var child in root.Children)
            {
                Collect(child, collector);
            }
        }
    }

    #endregion Private methods

    #region Nested classes

    public sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
    {
        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);

        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }

    #endregion Nested classes
}
