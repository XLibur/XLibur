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

            if (_style.IsCellContainer)
                SetKey(Key with { Horizontal = value });
            else
                Modify(k => k with { Horizontal = value });
            if (updateIndent)
                Indent = 0;
        }
    }

    public XLAlignmentVerticalValues Vertical
    {
        get => Key.Vertical;
        set
        {
            var key = Key;
            if (key.Vertical == value) return;
            if (_style.IsCellContainer)
                SetKey(key with { Vertical = value });
            else
                Modify(k => k with { Vertical = value });
        }
    }

    public int Indent
    {
        get => Key.Indent;
        set
        {
            if (Indent != value)
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
            if (_style.IsCellContainer)
                SetKey(Key with { Indent = value });
            else
                Modify(k => k with { Indent = value });
        }
    }

    public bool JustifyLastLine
    {
        get => Key.JustifyLastLine;
        set
        {
            var key = Key;
            if (key.JustifyLastLine == value) return;
            if (_style.IsCellContainer)
                SetKey(key with { JustifyLastLine = value });
            else
                Modify(k => k with { JustifyLastLine = value });
        }
    }

    public XLAlignmentReadingOrderValues ReadingOrder
    {
        get => Key.ReadingOrder;
        set
        {
            var key = Key;
            if (key.ReadingOrder == value) return;
            if (_style.IsCellContainer)
                SetKey(key with { ReadingOrder = value });
            else
                Modify(k => k with { ReadingOrder = value });
        }
    }

    public int RelativeIndent
    {
        get => Key.RelativeIndent;
        set
        {
            var key = Key;
            if (key.RelativeIndent == value) return;
            if (_style.IsCellContainer)
                SetKey(key with { RelativeIndent = value });
            else
                Modify(k => k with { RelativeIndent = value });
        }
    }

    public bool ShrinkToFit
    {
        get => Key.ShrinkToFit;
        set
        {
            var key = Key;
            if (key.ShrinkToFit == value) return;
            if (_style.IsCellContainer)
                SetKey(key with { ShrinkToFit = value });
            else
                Modify(k => k with { ShrinkToFit = value });
        }
    }

    public int TextRotation
    {
        get => Key.TextRotation;
        set
        {
            int rotation = value;

            if (rotation != 255 && (rotation < -90 || rotation > 90))
                throw new ArgumentException("TextRotation must be between -90 and 90 degrees, or 255.");

            var key = Key;
            if (key.TextRotation == rotation) return;
            if (_style.IsCellContainer)
                SetKey(key with { TextRotation = rotation });
            else
                Modify(k => k with { TextRotation = rotation });
        }
    }

    public bool WrapText
    {
        get => Key.WrapText;
        set
        {
            var key = Key;
            if (key.WrapText == value) return;
            if (_style.IsCellContainer)
                SetKey(key with { WrapText = value });
            else
                Modify(k => k with { WrapText = value });
        }
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

    private void Modify(Func<XLAlignmentKey, XLAlignmentKey> modification)
    {
        Key = modification(Key);
        _style.Modify(styleKey => styleKey with { Alignment = modification(styleKey.Alignment) });
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
