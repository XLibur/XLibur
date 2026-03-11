using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

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

    /// <inheritdoc cref="IXLStylized.InnerStyle"/>
    public IXLStyle InnerStyle
    {
        get => new XLStyle(this, StyleValue.Key);
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
        if (propagate)
        {
            Children.ForEach(child => child.SetStyle(StyleValue, true));
        }
    }

    private static readonly ReferenceEqualityComparer<XLStyleValue> _comparer = new();

    void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
    {
        var children = GetChildrenRecursively(this)
            .GroupBy(child => child.StyleValue, _comparer);

        foreach (var group in children)
        {
            var styleKey = modification(group.Key.Key);
            var styleValue = XLStyleValue.FromKey(ref styleKey);
            foreach (var child in group)
            {
                child.StyleValue = styleValue;
            }
        }
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

        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj!);
    }

    #endregion Nested classes
}
