using System;
using System.Text;

namespace XLibur.Excel;

internal sealed class XLAlignment : IXLAlignment
{
    #region Static members

    internal static XLAlignmentKey GenerateKey(IXLAlignment? d) => d switch
    {
        null => XLAlignmentValue.Default.Key,
        XLAlignment alignment => alignment.Key,
        _ => new XLAlignmentKey
        {
            Horizontal = d.Horizontal,
            Vertical = d.Vertical,
            Indent = d.Indent,
            JustifyLastLine = d.JustifyLastLine,
            ReadingOrder = d.ReadingOrder,
            RelativeIndent = d.RelativeIndent,
            ShrinkToFit = d.ShrinkToFit,
            TextRotation = d.TextRotation,
            WrapText = d.WrapText
        },
    };

    #endregion Static members

    #region Properties
    private readonly XLStyle _style;

    private XLAlignmentValue _value;

    /// <inheritdoc cref="XLFont.Key"/>
    private XLAlignmentKey Key
    {
        get
        {
            var pending = _style.Pending;
            return pending is null ? _value.Key : pending.Alignment;
        }
        set => _value = XLAlignmentValue.FromKey(ref value);
    }

    #endregion Properties

    #region Constructors

    /// <summary>
    /// Create an instance of XLAlignment initializing it with the specified value.
    /// </summary>
    /// <param name="style">Style to attach the new instance to.</param>
    /// <param name="value">Style value to use.</param>
    public XLAlignment(XLStyle? style, XLAlignmentValue value)
    {
        _style = style ?? XLStyle.CreateEmptyStyle();
        _value = value;
    }

    public XLAlignment(XLStyle? style, XLAlignmentKey key) : this(style, XLAlignmentValue.FromKey(ref key))
    {
    }

    public XLAlignment(XLStyle? style = null, IXLAlignment? d = null) : this(style, GenerateKey(d))
    {
    }

    #endregion Constructors

    internal void SyncValue(XLAlignmentValue value) { _value = value; }

    #region IXLAlignment Members

    public XLAlignmentHorizontalValues Horizontal
    {
        get => Key.Horizontal;
        set
        {
            bool updateIndent = !(
                value == XLAlignmentHorizontalValues.Left
                || value == XLAlignmentHorizontalValues.Right
                || value == XLAlignmentHorizontalValues.Distributed
            );

            Apply(false, value, static (k, v) => k with { Horizontal = v });
            if (updateIndent)
                Indent = 0;
        }
    }

    public XLAlignmentVerticalValues Vertical
    {
        get => Key.Vertical;
        set => Apply(Key.Vertical == value, value, static (k, v) => k with { Vertical = v });
    }

    public int Indent
    {
        get => Key.Indent;
        set
        {
            // An indent the style already holds is skipped wherever XLStyle.SkipsUnchangedValues
            // allows, as every other setter's unchanged value is. On a cell it was a no-op anyway;
            // a worksheet, row or column is spared a walk of its cells, and a pivot area a format
            // that changes nothing.
            if (Key.Indent == value && _style.SkipsUnchangedValues) return;

            if (!_style.IsWholeStyle)
            {
                // A range, a worksheet, a row or a column styles cells that need not share this
                // key, so each cell is asked for itself, inside the one modification - see
                // WithIndent. Nothing is validated against the key, so the write cannot stop
                // part-way through the cells (#505).
                Modify(value, static (k, v) => WithIndent(k, v));
                return;
            }

            if (Indent != value)
                PrepareHorizontalForIndent(value);

            Apply(false, value, static (k, v) => k with { Indent = v });
        }
    }

    /// <summary>
    /// Before the indent of a whole style changes: a general alignment becomes left, and an
    /// alignment that cannot take a positive indent is refused.
    /// </summary>
    private void PrepareHorizontalForIndent(int value)
    {
        if (Horizontal == XLAlignmentHorizontalValues.General)
            Horizontal = XLAlignmentHorizontalValues.Left;

        if (value > 0 && !(
                Horizontal == XLAlignmentHorizontalValues.Left
                || Horizontal == XLAlignmentHorizontalValues.Right
                || Horizontal == XLAlignmentHorizontalValues.Distributed
            ))
        {
            throw new ArgumentException(
                "For indents, only left, right, and distributed horizontal alignments are supported.");
        }
    }

    /// <summary>
    /// <paramref name="key"/> with <paramref name="value"/> as its indent, and left alignment where
    /// its own horizontal alignment cannot take one.
    /// </summary>
    /// <remarks>
    /// The <see cref="Indent"/> setter's rules, asked of one cell. A general alignment becomes left
    /// whenever the indent changes, as on a cell. Where a single cell refuses an indent its
    /// alignment cannot take, a cell styled through a range, worksheet, row or column is given left
    /// alignment instead: throwing there would leave the cells already written behind.
    /// </remarks>
    private static XLAlignmentKey WithIndent(XLAlignmentKey key, int value)
    {
        if (key.Indent == value)
            return key;

        var horizontal = key.Horizontal;
        if (horizontal == XLAlignmentHorizontalValues.General
            || (value > 0 && horizontal is not (XLAlignmentHorizontalValues.Left
                or XLAlignmentHorizontalValues.Right
                or XLAlignmentHorizontalValues.Distributed)))
        {
            horizontal = XLAlignmentHorizontalValues.Left;
        }

        return key with { Horizontal = horizontal, Indent = value };
    }

    public bool JustifyLastLine
    {
        get => Key.JustifyLastLine;
        set => Apply(Key.JustifyLastLine == value, value, static (k, v) => k with { JustifyLastLine = v });
    }

    public XLAlignmentReadingOrderValues ReadingOrder
    {
        get => Key.ReadingOrder;
        set => Apply(Key.ReadingOrder == value, value, static (k, v) => k with { ReadingOrder = v });
    }

    public int RelativeIndent
    {
        get => Key.RelativeIndent;
        set => Apply(Key.RelativeIndent == value, value, static (k, v) => k with { RelativeIndent = v });
    }

    public bool ShrinkToFit
    {
        get => Key.ShrinkToFit;
        set => Apply(Key.ShrinkToFit == value, value, static (k, v) => k with { ShrinkToFit = v });
    }

    public int TextRotation
    {
        get => Key.TextRotation;
        set
        {
            int rotation = value;

            if (rotation != 255 && (rotation < -90 || rotation > 90))
                throw new ArgumentException("TextRotation must be between -90 and 90 degrees, or 255.");

            Apply(Key.TextRotation == rotation, rotation, static (k, v) => k with { TextRotation = v });
        }
    }

    public bool WrapText
    {
        get => Key.WrapText;
        set => Apply(Key.WrapText == value, value, static (k, v) => k with { WrapText = v });
    }

    public bool TopToBottom
    {
        get => TextRotation == 255;
        set => TextRotation = value ? 255 : 0;
    }

    public IXLStyle SetHorizontal(XLAlignmentHorizontalValues value)
    {
        Horizontal = value;
        return _style;
    }

    public IXLStyle SetVertical(XLAlignmentVerticalValues value)
    {
        Vertical = value;
        return _style;
    }

    public IXLStyle SetIndent(int value)
    {
        Indent = value;
        return _style;
    }

    public IXLStyle SetJustifyLastLine()
    {
        JustifyLastLine = true;
        return _style;
    }

    public IXLStyle SetJustifyLastLine(bool value)
    {
        JustifyLastLine = value;
        return _style;
    }

    public IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value)
    {
        ReadingOrder = value;
        return _style;
    }

    public IXLStyle SetRelativeIndent(int value)
    {
        RelativeIndent = value;
        return _style;
    }

    public IXLStyle SetShrinkToFit()
    {
        ShrinkToFit = true;
        return _style;
    }

    public IXLStyle SetShrinkToFit(bool value)
    {
        ShrinkToFit = value;
        return _style;
    }

    public IXLStyle SetTextRotation(int value)
    {
        TextRotation = value;
        return _style;
    }

    public IXLStyle SetWrapText()
    {
        WrapText = true;
        return _style;
    }

    public IXLStyle SetWrapText(bool value)
    {
        WrapText = value;
        return _style;
    }

    public IXLStyle SetTopToBottom()
    {
        TopToBottom = true;
        return _style;
    }

    public IXLStyle SetTopToBottom(bool value)
    {
        TopToBottom = value;
        return _style;
    }

    #endregion

    /// <summary>
    /// Apply a new component key to the cell this facade is attached to.
    /// </summary>
    /// <remarks>
    /// The new key is deliberately <em>not</em> interned before being applied. Assigning it to
    /// <c>Key</c> first would run a repository lookup -- hashing the key and probing a dictionary --
    /// whose result is then thrown away: on a transition-cache hit <c>ModifyAlignment</c> never needs the
    /// component value at all, and on a miss it interns the component anyway inside
    /// <c>XLStyleValue.FromKey</c>. Taking the interned value back off the resulting style instead
    /// leaves the facade just as correct for later reads, at no lookup. Measured over 20,000 cells
    /// setting one property each, this was the single largest cost on the per-cell styling path.
    /// </remarks>
    private void SetKey(XLAlignmentKey newKey)
    {
        _style.ModifyAlignment(newKey);
        _value = _style.Value.Alignment;
    }

    /// <remarks>
    /// A setter skips a value equal to <see cref="Key"/> only where
    /// <see cref="XLStyle.SkipsUnchangedValues"/> allows it: on a cell or a worksheet. On a range or
    /// <c>IXLCells</c> the key is only that container's record of its style, which its cells need
    /// not share (#505), so there the setter always writes.
    /// <para>
    /// <c>Key</c> is assigned only where the facade does not hold the very value its style does - see
    /// <see cref="XLFont"/>'s <c>Modify</c>.
    /// </para>
    /// </remarks>
    private void Modify<T>(T value, Func<XLAlignmentKey, T, XLAlignmentKey> with)
    {
        if (!ReferenceEquals(_value, _style.Value.Alignment))
        {
            Key = with(Key, value);
            _style.Modify(styleKey => styleKey with { Alignment = with(styleKey.Alignment, value) });
            return;
        }

        _style.Modify(styleKey => styleKey with { Alignment = with(styleKey.Alignment, value) });
        _value = _style.Value.Alignment;
    }

    /// <inheritdoc cref="XLFont.Apply{T}"/>
    private void Apply<T>(bool unchanged, T value, Func<XLAlignmentKey, T, XLAlignmentKey> with)
    {
        if (unchanged && _style.SkipsUnchangedValues) return;
        if (_style.IsCellContainer)
            SetKey(with(Key, value));
        else
            Modify(value, with);
    }

    #region Overridden

    public override string ToString()
    {
        var sb = new StringBuilder();
        sb.Append(Horizontal);
        sb.Append('-');
        sb.Append(Vertical);
        sb.Append('-');
        sb.Append(Indent);
        sb.Append('-');
        sb.Append(JustifyLastLine);
        sb.Append('-');
        sb.Append(ReadingOrder);
        sb.Append('-');
        sb.Append(RelativeIndent);
        sb.Append('-');
        sb.Append(ShrinkToFit);
        sb.Append('-');
        sb.Append(TextRotation);
        sb.Append('-');
        sb.Append(WrapText);
        sb.Append('-');
        return sb.ToString();
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as XLAlignment);
    }

    public bool Equals(IXLAlignment? other)
    {
        var otherA = other as XLAlignment;
        if (otherA == null)
            return false;

        return Key == otherA.Key;
    }

    public override int GetHashCode()
    {
        var hashCode = 1214962009;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden
}
