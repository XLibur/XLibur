using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A base class for pivot styling API. It has takes a selected <see cref="XLPivotArea"/>
/// and applies the style using <c>.Style*</c> API. The derived classes are responsible for
/// exposing API so user can define an area and then create the desired area (from what user
/// specified) through <see cref="GetCurrentArea"/> method.
/// </summary>
internal abstract class XLPivotStyleFormatBase : IXLPivotStyleFormat, IXLStylized
{
    protected readonly XLPivotTable PivotTable;
    private XLStyleValue _styleValue;

    protected XLPivotStyleFormatBase(XLPivotTable pivotTable)
    {
        PivotTable = pivotTable;

        // Only used until a matching format exists, see the StyleValue getter.
        _styleValue = XLStyle.Default.Value;
    }

    #region IXLPivotStyleFormat members

    public XLPivotStyleFormatElement AppliesTo { get; init; } = XLPivotStyleFormatElement.Data;

    public IXLStyle Style
    {
        get => InnerStyle;
        set => InnerStyle = value;
    }

    #endregion IXLPivotStyleFormat members

    #region IXLStylized

    public IXLStyle InnerStyle
    {
        get => new XLStyle(this, StyleValue);
        set
        {
            var styleKey = XLStyle.GenerateKey(value);
            StyleValue = XLStyleValue.FromKey(ref styleKey);
        }
    }
    public IXLRanges RangesUsed { get; } = new XLRanges();

    public XLStyleValue StyleValue
    {
        // Read through to the pivot table's formats, which is where the style actually lives and
        // where loading a file puts it. These style format objects are created fresh on every
        // property access and the area they match can still be narrowed after construction
        // (IXLPivotValueStyleFormat.AndWith), so nothing is cached. Unlike GetFormats, this does
        // not create a format when none matches -- reading a style must not write one.
        get => TryGetFormatStyleValue(out var formatStyleValue) ? formatStyleValue : _styleValue;
        set
        {
            // This sets the style of everything to the passed style, while ModifyStyle
            // is for fluent API that can modify format styles individually. Because initial
            // value of _styleValue is Default, this setter shouldn't be used as a basis
            // for modifying the DxStyleValue.
            _styleValue = value;
            foreach (var format in GetFormats())
                format.DxfStyleValue = value;
        }
    }

    public void ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
    {
        // Seeded from _styleValue, not StyleValue: this is the style a format gets if it has to be
        // created, and a new format starts from Default. Formats that already exist are modified
        // from their own value in the loop below.
        var styleKey = modification(_styleValue.Key);
        _styleValue = XLStyleValue.FromKey(ref styleKey);

        // Do not use StyleValue setter, because some formats might have different formats and
        // we should only modify them, not replace other potentially different style props of formats.
        foreach (var format in GetFormats())
        {
            var formatStyleValue = modification(format.DxfStyleValue.Key);
            format.DxfStyleValue = XLStyleValue.FromKey(ref formatStyleValue);
        }
    }

    #endregion IXLStylized

    internal abstract XLPivotArea GetCurrentArea();

    internal abstract bool Filter(XLPivotArea area);

    /// <summary>
    /// Whether a format's area styles the cells this instance stands for. Defaults to
    /// <see cref="Filter"/>, the exact area used when writing, and is loosened where a file can
    /// legitimately hold a wider area that still styles those cells.
    /// </summary>
    internal virtual bool Covers(XLPivotArea area) => Filter(area);

    /// <summary>
    /// The style of the first format whose area this instance selects, if the pivot table has one.
    /// </summary>
    private bool TryGetFormatStyleValue(out XLStyleValue styleValue)
    {
        foreach (var format in PivotTable.Formats)
        {
            if (format.Action == XLPivotFormatAction.Formatting && Covers(format.PivotArea))
            {
                styleValue = format.DxfStyleValue;
                return true;
            }
        }

        styleValue = XLStyle.Default.Value;
        return false;
    }

    private IEnumerable<XLPivotFormat> GetFormats()
    {
        var exists = false;
        foreach (var format in PivotTable.Formats)
        {
            if (format.Action == XLPivotFormatAction.Formatting && Filter(format.PivotArea))
            {
                exists = true;
                yield return format;
            }
        }

        if (!exists)
        {
            var format = new XLPivotFormat(GetCurrentArea())
            {
                DxfStyleValue = _styleValue
            };
            PivotTable.AddFormat(format);
            yield return format;
        }
    }
}
