using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLibur.Excel.ConditionalFormats;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLBorder : IXLBorder
{
    private const string ColorCannotBeNull = "Color cannot be null";

    #region Static members

    internal static XLBorderKey GenerateKey(IXLBorder? defaultBorder) => defaultBorder switch
    {
        null => XLBorderValue.Default.Key,
        XLBorder border => border.Key,
        _ => new XLBorderKey
        {
            LeftBorder = defaultBorder.LeftBorder,
            LeftBorderColor = defaultBorder.LeftBorderColor.Key,
            RightBorder = defaultBorder.RightBorder,
            RightBorderColor = defaultBorder.RightBorderColor.Key,
            TopBorder = defaultBorder.TopBorder,
            TopBorderColor = defaultBorder.TopBorderColor.Key,
            BottomBorder = defaultBorder.BottomBorder,
            BottomBorderColor = defaultBorder.BottomBorderColor.Key,
            DiagonalBorder = defaultBorder.DiagonalBorder,
            DiagonalBorderColor = defaultBorder.DiagonalBorderColor.Key,
            DiagonalUp = defaultBorder.DiagonalUp,
            DiagonalDown = defaultBorder.DiagonalDown,
        },
    };

    #endregion Static members

    private readonly XLStyle _style;

    private readonly IXLStylized _container;

    private XLBorderValue _value;

    /// <summary>
    /// A colour assigned to an edge while that edge has no style, held here instead of being
    /// written through - and one field per edge, since each edge's style and colour move
    /// independently of the others.
    /// </summary>
    /// <remarks>
    /// An edge with no style has no colour to draw with, and <see cref="XLBorderKey.Normalize"/>
    /// replaces such a colour with black the moment the key reaches a repository or an
    /// <c>XLStyleKey</c>. Without this, the ordinary call order
    /// <c>border.TopBorderColor = red; border.TopBorder = XLBorderStyleValues.Thin;</c> would set
    /// the colour, have normalization erase it straight back to black, and only then apply the
    /// style - the edge ends up black rather than red, and neither statement looks responsible.
    /// <see cref="ApplyEdgeStyle"/> re-applies whichever colour is still pending here when the edge
    /// is given a style, so the two statements read in either order give the same result. A colour
    /// assigned to an edge that already has a style needs none of this and is written through at
    /// once, as before.
    /// <para>
    /// Where the key is everything styled, meaningful only while the corresponding edge has no
    /// style. <see cref="ApplyEdgeStyle"/> clears it once it has been read, whether or not the new
    /// style made it applicable, and it counts only against the ground truth it was recorded
    /// against: <see cref="SyncValue"/> clears every pending colour whenever the interned border
    /// changes for a reason other than this facade's own writes, so a stale colour can never be
    /// applied to a style the caller does not know it still holds.
    /// </para>
    /// <para>
    /// "This edge has no style" describes the thing styled only where the key is everything
    /// styled: a cell, or a style with no cells under it - see <see cref="XLStyle.IsWholeStyle"/>.
    /// For a range, <c>IXLCells</c>, a worksheet, a row or a column the key is only the container's
    /// own record of its style, which its cells need not share. A range is held weakly by its
    /// worksheet, and one rebuilt after a collection starts from its parent's style rather than its
    /// cells', so it can see an edge as unstyled that its cells have styled - which is how
    /// <see cref="OutsideBorderColor"/>, working through <c>LastColumn()</c> and the like, once lost
    /// its colour at random (#505) - or as styled where a cell has none. So there the colour is
    /// written through at once, for each cell whose edge has a style, and held here whatever the
    /// record says, for the style this facade may be given next. <see cref="ApplyEdgeStyle"/> then
    /// writes that style and the colour to each cell in the same write, so a cell whose edge had no
    /// style takes both together instead of having the colour normalized away first. An unchanged
    /// style, or an unchanged colour on an edge that has one, is skipped wherever
    /// <see cref="XLStyle.SkipsUnchangedValues"/> allows, and written through anywhere else.
    /// </para>
    /// <para>
    /// The pending colour lives only on this facade, and so only on the style object it was set
    /// through. A range is held weakly by its worksheet: two <c>ws.Range(...)</c> calls usually
    /// return one cached object, but after a collection the second builds a new one, whose facade
    /// holds no pending colour, and a cell whose edge had no style then takes the next style alone,
    /// in black. So colour then style should go through one object the caller keeps. Making this
    /// independent of the collector would mean storing a colour on an unstyled edge in the cell
    /// itself, which would change the saved output and how styles are shared - out of scope for
    /// #505.
    /// </para>
    /// <para>
    /// Facade-local, and only for the five single-edge properties: <see cref="OutsideBorder"/>,
    /// <see cref="OutsideBorderColor"/>, <see cref="InsideBorder"/> and
    /// <see cref="InsideBorderColor"/> apply their four edges as one combined write and do not use
    /// it, so a colour assigned there before the matching style keeps the behaviour these remarks
    /// describe as broken - callers of the compound setters still set style before colour.
    /// </para>
    /// </remarks>
    private XLColorKey? _pendingLeftBorderColor;
    private XLColorKey? _pendingRightBorderColor;
    private XLColorKey? _pendingTopBorderColor;
    private XLColorKey? _pendingBottomBorderColor;
    private XLColorKey? _pendingDiagonalBorderColor;

    /// <remarks>
    /// Read-only. The setter this replaces interned the assigned key in the border repository, and
    /// every caller then went on to resolve a style that interned it again - see
    /// <see cref="SetKey"/> and <see cref="Modify"/>, which both take the interned value off the
    /// resulting style instead.
    /// <para>
    /// While the style is batching, the key comes from the style's pending key rather than from
    /// <see cref="_value"/>: a batch resolves nothing until it flushes, so the cached value would
    /// report pre-batch borders to every getter and to <see cref="ApplyEdgeStyle"/>'s own
    /// current-style test. Outside a batch <see cref="_value"/> stays the source of truth - a facade
    /// can be constructed over a key that its style does not hold (see the constructors that pass a
    /// null style), and those must keep reading the key they were given.
    /// </para>
    /// </remarks>
    internal XLBorderKey Key
    {
        get
        {
            var pending = _style.Pending;
            return pending is null ? _value.Key : pending.Border;
        }
    }

    #region Constructors

    /// <summary>
    /// Create an instance of XLBorder initializing it with the specified value.
    /// </summary>
    /// <param name="container">Container the border is applied to.</param>
    /// <param name="style">Style to attach the new instance to.</param>
    /// <param name="value">Style value to use.</param>
    public XLBorder(IXLStylized container, XLStyle style, XLBorderValue value)
    {
        _container = container;
        _style = style ?? (_container.Style as XLStyle) ?? XLStyle.CreateEmptyStyle();
        _value = value;
    }

    public XLBorder(IXLStylized container, XLStyle style, XLBorderKey key) : this(container, style, XLBorderValue.FromKey(ref key))
    {
    }

    public XLBorder(IXLStylized container, XLStyle? style = null, IXLBorder? d = null) : this(container, style!, GenerateKey(d))
    {
    }

    #endregion Constructors

    /// <remarks>
    /// Clears every pending colour whenever the incoming value's key differs from the one already
    /// held - see the remarks on <see cref="_pendingLeftBorderColor"/>. Compared by key rather than
    /// by reference: values are interned, so identical keys are ordinarily the same instance, but
    /// the repository holds only weak references and can hand back a second instance for a key
    /// still in use once the first is collected. Comparing keys makes that indistinguishable from
    /// "nothing changed" rather than a false trigger that would silently drop a pending colour.
    /// </remarks>
    /// <summary>
    /// Take the interned value that this facade's own writes produced, leaving the pending colours
    /// alone - the bookkeeping <see cref="SetKey"/> does on the direct path, for a caller that
    /// applied the writes some other way.
    /// </summary>
    /// <remarks>
    /// Distinct from <see cref="SyncValue"/>, which exists for the opposite case: the style moved
    /// for a reason this facade did not cause, so a colour pending against the old key must not
    /// survive onto the new one.
    /// </remarks>
    internal void RefreshValue(XLBorderValue value) => _value = value;

    internal void SyncValue(XLBorderValue value)
    {
        if (!value.Key.Equals(_value.Key))
        {
            _pendingLeftBorderColor = null;
            _pendingRightBorderColor = null;
            _pendingTopBorderColor = null;
            _pendingBottomBorderColor = null;
            _pendingDiagonalBorderColor = null;
        }

        _value = value;
    }

    #region IXLBorder Members

#pragma warning disable S2376
    public XLBorderStyleValues OutsideBorder
    {
        set
        {
            if (_container is null or XLWorksheet or XLConditionalFormat or XLCell)
            {
                Modify(k => k with
                {
                    TopBorder = value,
                    BottomBorder = value,
                    LeftBorder = value,
                    RightBorder = value,
                });
            }
            else
            {
                foreach (var r in _container.RangesUsed)
                {
                    r.FirstColumn()!.Style.Border.LeftBorder = value;
                    r.LastColumn()!.Style.Border.RightBorder = value;
                    r.FirstRow()!.Style.Border.TopBorder = value;
                    r.LastRow()!.Style.Border.BottomBorder = value;
                }
            }
        }
    }

    public XLColor OutsideBorderColor
    {
        set
        {
            if (_container is null or XLWorksheet or XLConditionalFormat or XLCell)
            {
                Modify(k => k with
                {
                    TopBorderColor = value.Key,
                    BottomBorderColor = value.Key,
                    LeftBorderColor = value.Key,
                    RightBorderColor = value.Key,
                });
            }
            else
            {
                foreach (var r in _container.RangesUsed)
                {
                    r.FirstColumn()!.Style.Border.LeftBorderColor = value;
                    r.LastColumn()!.Style.Border.RightBorderColor = value;
                    r.FirstRow()!.Style.Border.TopBorderColor = value;
                    r.LastRow()!.Style.Border.BottomBorderColor = value;
                }
            }
        }
    }

    public XLBorderStyleValues InsideBorder
    {
        set
        {
            if (_container is null or XLWorksheet)
            {
                Modify(k => k with
                {
                    TopBorder = value,
                    BottomBorder = value,
                    LeftBorder = value,
                    RightBorder = value,
                });
            }
            else
            {
                foreach (var r in _container.RangesUsed)
                {
                    using (new RestoreOutsideBorder(r))
                    {
                        foreach (var cell in r.Cells())
                        {
                            ((XLBorder)cell.Style.Border)
                                .Modify(k => k with
                                {
                                    TopBorder = value,
                                    BottomBorder = value,
                                    LeftBorder = value,
                                    RightBorder = value,
                                });
                        }
                    }
                }
            }
        }
    }

    public XLColor InsideBorderColor
    {
        set
        {
            if (_container is null or XLWorksheet)
            {
                Modify(k => k with
                {
                    TopBorderColor = value.Key,
                    BottomBorderColor = value.Key,
                    LeftBorderColor = value.Key,
                    RightBorderColor = value.Key,
                });
            }
            else
            {
                foreach (var r in _container.RangesUsed)
                {
                    using (new RestoreOutsideBorder(r))
                    {
                        foreach (var cell in r.Cells())
                        {
                            ((XLBorder)cell.Style.Border)
                                .Modify(k => k with
                                {
                                    TopBorderColor = value.Key,
                                    BottomBorderColor = value.Key,
                                    LeftBorderColor = value.Key,
                                    RightBorderColor = value.Key,
                                });
                        }
                    }
                }
            }
        }
    }
#pragma warning restore S2376

    public XLBorderStyleValues LeftBorder
    {
        get => Key.LeftBorder;
        set => ApplyEdgeStyle(value, Key.LeftBorder, ref _pendingLeftBorderColor,
            static (k, s) => k with { LeftBorder = s },
            static (k, c) => k with { LeftBorderColor = c });
    }

    public XLColor LeftBorderColor
    {
        get
        {
            var colorKey = Key.LeftBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set => ApplyEdgeColor(value, Key.LeftBorder, Key.LeftBorderColor, ref _pendingLeftBorderColor,
            static k => k.LeftBorder, static (k, c) => k with { LeftBorderColor = c });
    }

    public XLBorderStyleValues RightBorder
    {
        get => Key.RightBorder;
        set => ApplyEdgeStyle(value, Key.RightBorder, ref _pendingRightBorderColor,
            static (k, s) => k with { RightBorder = s },
            static (k, c) => k with { RightBorderColor = c });
    }

    public XLColor RightBorderColor
    {
        get
        {
            var colorKey = Key.RightBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set => ApplyEdgeColor(value, Key.RightBorder, Key.RightBorderColor, ref _pendingRightBorderColor,
            static k => k.RightBorder, static (k, c) => k with { RightBorderColor = c });
    }

    public XLBorderStyleValues TopBorder
    {
        get => Key.TopBorder;
        set => ApplyEdgeStyle(value, Key.TopBorder, ref _pendingTopBorderColor,
            static (k, s) => k with { TopBorder = s },
            static (k, c) => k with { TopBorderColor = c });
    }

    public XLColor TopBorderColor
    {
        get
        {
            var colorKey = Key.TopBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set => ApplyEdgeColor(value, Key.TopBorder, Key.TopBorderColor, ref _pendingTopBorderColor,
            static k => k.TopBorder, static (k, c) => k with { TopBorderColor = c });
    }

    public XLBorderStyleValues BottomBorder
    {
        get => Key.BottomBorder;
        set => ApplyEdgeStyle(value, Key.BottomBorder, ref _pendingBottomBorderColor,
            static (k, s) => k with { BottomBorder = s },
            static (k, c) => k with { BottomBorderColor = c });
    }

    public XLColor BottomBorderColor
    {
        get
        {
            var colorKey = Key.BottomBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set => ApplyEdgeColor(value, Key.BottomBorder, Key.BottomBorderColor, ref _pendingBottomBorderColor,
            static k => k.BottomBorder, static (k, c) => k with { BottomBorderColor = c });
    }

    public XLBorderStyleValues DiagonalBorder
    {
        get => Key.DiagonalBorder;
        set => ApplyEdgeStyle(value, Key.DiagonalBorder, ref _pendingDiagonalBorderColor,
            static (k, s) => k with { DiagonalBorder = s },
            static (k, c) => k with { DiagonalBorderColor = c });
    }

    public XLColor DiagonalBorderColor
    {
        get
        {
            var colorKey = Key.DiagonalBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set => ApplyEdgeColor(value, Key.DiagonalBorder, Key.DiagonalBorderColor, ref _pendingDiagonalBorderColor,
            static k => k.DiagonalBorder, static (k, c) => k with { DiagonalBorderColor = c });
    }

    public bool DiagonalUp
    {
        get => Key.DiagonalUp;
        set
        {
            var key = Key;
            if (key.DiagonalUp == value && _style.SkipsUnchangedValues) return;
            if (_style.IsCellContainer)
                SetKey(key with { DiagonalUp = value });
            else
                Modify(k => k with { DiagonalUp = value });
        }
    }

    public bool DiagonalDown
    {
        get => Key.DiagonalDown;
        set
        {
            var key = Key;
            if (key.DiagonalDown == value && _style.SkipsUnchangedValues) return;
            if (_style.IsCellContainer)
                SetKey(key with { DiagonalDown = value });
            else
                Modify(k => k with { DiagonalDown = value });
        }
    }

    public IXLStyle SetOutsideBorder(XLBorderStyleValues value)
    {
        OutsideBorder = value;
        return _style;
    }

    public IXLStyle SetOutsideBorderColor(XLColor value)
    {
        OutsideBorderColor = value;
        return _style;
    }

    public IXLStyle SetInsideBorder(XLBorderStyleValues value)
    {
        InsideBorder = value;
        return _style;
    }

    public IXLStyle SetInsideBorderColor(XLColor value)
    {
        InsideBorderColor = value;
        return _style;
    }

    public IXLStyle SetLeftBorder(XLBorderStyleValues value)
    {
        LeftBorder = value;
        return _style;
    }

    public IXLStyle SetLeftBorderColor(XLColor value)
    {
        LeftBorderColor = value;
        return _style;
    }

    public IXLStyle SetRightBorder(XLBorderStyleValues value)
    {
        RightBorder = value;
        return _style;
    }

    public IXLStyle SetRightBorderColor(XLColor value)
    {
        RightBorderColor = value;
        return _style;
    }

    public IXLStyle SetTopBorder(XLBorderStyleValues value)
    {
        TopBorder = value;
        return _style;
    }

    public IXLStyle SetTopBorderColor(XLColor value)
    {
        TopBorderColor = value;
        return _style;
    }

    public IXLStyle SetBottomBorder(XLBorderStyleValues value)
    {
        BottomBorder = value;
        return _style;
    }

    public IXLStyle SetBottomBorderColor(XLColor value)
    {
        BottomBorderColor = value;
        return _style;
    }

    public IXLStyle SetDiagonalUp()
    {
        DiagonalUp = true;
        return _style;
    }

    public IXLStyle SetDiagonalUp(bool value)
    {
        DiagonalUp = value;
        return _style;
    }

    public IXLStyle SetDiagonalDown()
    {
        DiagonalDown = true;
        return _style;
    }

    public IXLStyle SetDiagonalDown(bool value)
    {
        DiagonalDown = value;
        return _style;
    }

    public IXLStyle SetDiagonalBorder(XLBorderStyleValues value)
    {
        DiagonalBorder = value;
        return _style;
    }

    public IXLStyle SetDiagonalBorderColor(XLColor value)
    {
        DiagonalBorderColor = value;
        return _style;
    }

    #endregion IXLBorder Members

    /// <summary>
    /// Set one edge's style, applying whichever colour is pending for it - see
    /// <see cref="_pendingLeftBorderColor"/> - if the new style makes one applicable.
    /// </summary>
    /// <param name="newStyle">Style to give the edge.</param>
    /// <param name="currentStyle">The edge's style before this call.</param>
    /// <param name="pendingColor">The pending-colour field for this edge specifically.</param>
    /// <param name="withStyle">Rewrites a border key's style for this edge.</param>
    /// <param name="withColor">Rewrites a border key's colour for this edge.</param>
    private void ApplyEdgeStyle(
        XLBorderStyleValues newStyle,
        XLBorderStyleValues currentStyle,
        ref XLColorKey? pendingColor,
        Func<XLBorderKey, XLBorderStyleValues, XLBorderKey> withStyle,
        Func<XLBorderKey, XLColorKey, XLBorderKey> withColor)
    {
        // An unchanged style leaves nothing to do wherever XLStyle.SkipsUnchangedValues allows the
        // skip, and a pending colour stays for the next style, as it always has. A selection's key
        // is only its own record of its style, which its cells need not share, so there the style
        // is written whatever the record says - with any pending colour, in the same write, so a
        // cell whose edge had no style takes both together (#505).
        if (currentStyle == newStyle && _style.SkipsUnchangedValues) return;

        // A colour pending against the old style is either about to become applicable (None to a
        // style) or moot (any style to None, since the edge is about to have nothing to draw
        // regardless of what colour it was given). Either way it does not survive this call.
        var colorToApply = newStyle != XLBorderStyleValues.None ? pendingColor : null;
        pendingColor = null;

        if (_style.IsCellContainer)
        {
            var newKey = withStyle(Key, newStyle);
            if (colorToApply is { } color)
                newKey = withColor(newKey, color);
            SetKey(newKey);
        }
        else
        {
            Modify(k =>
            {
                var newKey = withStyle(k, newStyle);
                return colorToApply is { } color ? withColor(newKey, color) : newKey;
            });
        }
    }

    /// <summary>
    /// Set one edge's colour, or - if the edge currently has no style - record it as pending
    /// instead of writing it through. See <see cref="_pendingLeftBorderColor"/> for why.
    /// </summary>
    /// <remarks>
    /// Held back only where the key is everything styled - see <see cref="XLStyle.IsWholeStyle"/>.
    /// Anywhere else the colour is written through, and held pending, whatever the edge's style
    /// here, and each cell decides for itself - see the remarks on
    /// <see cref="_pendingLeftBorderColor"/>. An unchanged colour on an edge that has a style is
    /// still skipped wherever <see cref="XLStyle.SkipsUnchangedValues"/> allows.
    /// </remarks>
    /// <param name="value">Colour to give the edge.</param>
    /// <param name="currentStyle">The edge's style as this facade sees it, which decides whether
    /// the colour is held pending.</param>
    /// <param name="currentColor">The edge's colour before this call.</param>
    /// <param name="pendingColor">The pending-colour field for this edge specifically.</param>
    /// <param name="styleOf">Reads a border key's style for this edge, to tell which cells can take
    /// the colour.</param>
    /// <param name="withColor">Rewrites a border key's colour for this edge.</param>
    private void ApplyEdgeColor(
        XLColor value,
        XLBorderStyleValues currentStyle,
        XLColorKey currentColor,
        ref XLColorKey? pendingColor,
        Func<XLBorderKey, XLBorderStyleValues> styleOf,
        Func<XLBorderKey, XLColorKey, XLBorderKey> withColor)
    {
        if (value == null)
            throw new ArgumentNullException(nameof(value), ColorCannotBeNull);

        var edgeHasNoStyle = currentStyle == XLBorderStyleValues.None;

        if (_style.IsWholeStyle)
        {
            ApplyEdgeColorToWholeStyle(value, edgeHasNoStyle, currentColor, ref pendingColor, withColor);
            return;
        }

        ApplyEdgeColorThroughCells(value, edgeHasNoStyle, currentColor, ref pendingColor, styleOf, withColor);
    }

    private void ApplyEdgeColorToWholeStyle(
        XLColor value,
        bool edgeHasNoStyle,
        XLColorKey currentColor,
        ref XLColorKey? pendingColor,
        Func<XLBorderKey, XLColorKey, XLBorderKey> withColor)
    {
        // The key is everything styled, so an edge it gives no style has nothing to draw with:
        // the colour is held for the style this facade may be given next, and nothing is
        // written - which also spares a pivot area a format that would change nothing.
        pendingColor = edgeHasNoStyle ? value.Key : null;
        if (edgeHasNoStyle || currentColor == value.Key) return;

        if (_style.IsCellContainer)
            SetKey(withColor(Key, value.Key));
        else
            Modify(k => withColor(k, value.Key));
    }

    private void ApplyEdgeColorThroughCells(
        XLColor value,
        bool edgeHasNoStyle,
        XLColorKey currentColor,
        ref XLColorKey? pendingColor,
        Func<XLBorderKey, XLBorderStyleValues> styleOf,
        Func<XLBorderKey, XLColorKey, XLBorderKey> withColor)
    {
        // Which cells have no style on this edge is not something the container's record can say.
        // So the colour is written through now, to each cell whose edge has a style, and kept
        // pending for the style this facade may be given next only if something could not take
        // it: a cell whose edge has no style, seen as the write passes over its key, or the
        // container's own record, whose unstyled edge is the colour-then-style case the pending
        // colour exists for. A colour every cell took is not kept, because it would paint over a
        // colour a cell was given directly in between (#505 review).
        //
        // Known limit, accepted: once some cell could not take the colour, the next style paints
        // it on every cell - including one given its own colour directly in between. Telling the
        // two apart would need a record of each cell's colour, which the container does not keep.
        if (!edgeHasNoStyle && currentColor == value.Key && _style.SkipsUnchangedValues)
        {
            pendingColor = null;
            return;
        }

        var somethingCouldNotTakeIt = edgeHasNoStyle;
        Modify(k =>
        {
            if (styleOf(k) == XLBorderStyleValues.None)
                somethingCouldNotTakeIt = true;
            return withColor(k, value.Key);
        });
        pendingColor = somethingCouldNotTakeIt ? value.Key : null;
    }

    /// <summary>
    /// Apply a new component key to the cell this facade is attached to.
    /// </summary>
    /// <remarks>
    /// The new key is deliberately <em>not</em> interned before being applied. Assigning it to
    /// <c>Key</c> first would run a repository lookup -- hashing the key and probing a dictionary --
    /// whose result is then thrown away: on a transition-cache hit <c>ModifyBorder</c> never needs the
    /// component value at all, and on a miss it interns the component anyway inside
    /// <c>XLStyleValue.FromKey</c>. Taking the interned value back off the resulting style instead
    /// leaves the facade just as correct for later reads, at no lookup. Measured over 20,000 cells
    /// setting one property each, this was the single largest cost on the per-cell styling path.
    /// </remarks>
    private void SetKey(XLBorderKey newKey)
    {
        _style.ModifyBorder(newKey);
        _value = _style.Value.Border;
    }

    /// <summary>
    /// Kept for <see cref="RestoreOutsideBorder"/>, compound inside-border operations,
    /// and non-cell-container paths that need per-property delta applied to each cell.
    /// </summary>
    /// <remarks>
    /// Neither branch assigns <c>Key</c>, which would intern the new border in its repository only
    /// for the result to be discarded - the same wasted lookup <see cref="SetKey"/> documents. A cell
    /// container never needs the component value, and a non-cell container interns the border again
    /// inside <c>XLStyleValue</c> when the style key is resolved. Taking the interned value back off
    /// the resulting style leaves the facade just as correct for later reads, at no lookup.
    /// <para>
    /// Assigning it also ran <paramref name="modification"/> an extra time, on the facade's own key,
    /// before the style ran it again on its own.
    /// </para>
    /// </remarks>
    private void Modify(Func<XLBorderKey, XLBorderKey> modification)
    {
        if (_style.IsCellContainer)
        {
            SetKey(modification(Key));
            return;
        }

        _style.Modify(styleKey => styleKey with { Border = modification(styleKey.Border) });
        _value = _style.Value.Border;
    }

    #region Overridden

    public override string ToString()
    {
        var sb = new StringBuilder();
        sb.Append(LeftBorder.ToString());
        sb.Append('-');
        sb.Append(LeftBorderColor);
        sb.Append('-');
        sb.Append(RightBorder.ToString());
        sb.Append('-');
        sb.Append(RightBorderColor);
        sb.Append('-');
        sb.Append(TopBorder.ToString());
        sb.Append('-');
        sb.Append(TopBorderColor);
        sb.Append('-');
        sb.Append(BottomBorder.ToString());
        sb.Append('-');
        sb.Append(BottomBorderColor);
        sb.Append('-');
        sb.Append(DiagonalBorder.ToString());
        sb.Append('-');
        sb.Append(DiagonalBorderColor);
        sb.Append('-');
        sb.Append(DiagonalUp);
        sb.Append('-');
        sb.Append(DiagonalDown);
        return sb.ToString();
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as XLBorder);
    }

    public bool Equals(IXLBorder? other)
    {
        var otherB = other as XLBorder;
        if (otherB == null)
            return false;

        return Key == otherB.Key;
    }

    public override int GetHashCode()
    {
        var hashCode = 416600561;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden

    /// <summary>
    /// Helper class that remembers outside border state before editing (in constructor) and restore afterwards (on disposing).
    /// It presumes that size of the range does not change during the editing, else it will fail.
    /// </summary>
    private sealed class RestoreOutsideBorder : IDisposable
    {
        private readonly IXLRange _range;
        private readonly Dictionary<int, XLBorderKey> _topBorders;
        private readonly Dictionary<int, XLBorderKey> _bottomBorders;
        private readonly Dictionary<int, XLBorderKey> _leftBorders;
        private readonly Dictionary<int, XLBorderKey> _rightBorders;

        public RestoreOutsideBorder(IXLRange range)
        {
            _range = range ?? throw new ArgumentNullException(nameof(range));

            _topBorders = range.FirstRow()!.Cells().ToDictionary(
                c => c.Address.ColumnNumber - range.RangeAddress.FirstAddress.ColumnNumber + 1,
                c => ((XLStyle)c.Style).Key.Border);

            _bottomBorders = range.LastRow()!.Cells().ToDictionary(
                c => c.Address.ColumnNumber - range.RangeAddress.FirstAddress.ColumnNumber + 1,
                c => ((XLStyle)c.Style).Key.Border);

            _leftBorders = range.FirstColumn()!.Cells().ToDictionary(
                c => c.Address.RowNumber - range.RangeAddress.FirstAddress.RowNumber + 1,
                c => ((XLStyle)c.Style).Key.Border);

            _rightBorders = range.LastColumn()!.Cells().ToDictionary(
                c => c.Address.RowNumber - range.RangeAddress.FirstAddress.RowNumber + 1,
                c => ((XLStyle)c.Style).Key.Border);
        }

        public void Dispose()
        {
            _topBorders.ForEach(kp => ((XLBorder)_range.FirstRow()!.Cell(kp.Key).Style
                .Border).Modify(k => k with
                {
                    TopBorder = kp.Value.TopBorder,
                    TopBorderColor = kp.Value.TopBorderColor,
                }));
            _bottomBorders.ForEach(kp => ((XLBorder)_range.LastRow()!.Cell(kp.Key).Style
                .Border).Modify(k => k with
                {
                    BottomBorder = kp.Value.BottomBorder,
                    BottomBorderColor = kp.Value.BottomBorderColor,
                }));
            _leftBorders.ForEach(kp => ((XLBorder)_range.FirstColumn()!.Cell(kp.Key).Style
                .Border).Modify(k => k with
                {
                    LeftBorder = kp.Value.LeftBorder,
                    LeftBorderColor = kp.Value.LeftBorderColor,
                }));
            _rightBorders.ForEach(kp => ((XLBorder)_range.LastColumn()!.Cell(kp.Key).Style
                .Border).Modify(k => k with
                {
                    RightBorder = kp.Value.RightBorder,
                    RightBorderColor = kp.Value.RightBorderColor,
                }));
            GC.SuppressFinalize(this);
        }
    }
}
