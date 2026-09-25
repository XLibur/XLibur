using System;
using System.Globalization;
using System.Text;

namespace XLibur.Excel;

internal sealed class XLFont : IXLFont
{
    #region Static members

    public static IXLFontBase DefaultCommentFont
    {
        get
        {
            // MS Excel uses Tahoma 9 Swiss no matter what current style font
            var defaultCommentFont = new XLFont
            {
                FontName = "Tahoma",
                FontSize = 9,
                FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss
            };

            return defaultCommentFont;
        }
    }

    internal static XLFontKey GenerateKey(IXLFontBase? defaultFont)
    {
        return defaultFont switch
        {
            null => XLFontValue.Default.Key,
            XLFont font => font.Key,
            _ => new XLFontKey
            {
                Bold = defaultFont.Bold,
                Italic = defaultFont.Italic,
                Underline = defaultFont.Underline,
                Strikethrough = defaultFont.Strikethrough,
                VerticalAlignment = defaultFont.VerticalAlignment,
                Shadow = defaultFont.Shadow,
                FontSize = defaultFont.FontSize,
                FontColor = defaultFont.FontColor.Key,
                FontName = defaultFont.FontName,
                FontFamilyNumbering = defaultFont.FontFamilyNumbering,
                FontCharSet = defaultFont.FontCharSet,
                FontScheme = defaultFont.FontScheme
            }
        };
    }

    #endregion Static members

    private readonly XLStyle _style;

    private XLFontValue _value;

    /// <remarks>
    /// While the style is batching, the key comes from the style's pending key rather than from
    /// <see cref="_value"/>: a batch resolves nothing until it flushes, so the cached value would
    /// report pre-batch values to every getter. Outside a batch <see cref="_value"/> stays the
    /// source of truth - a facade can be constructed over a key its style does not hold (see the
    /// constructors that pass a null style), and those must keep reading the key they were given.
    /// </remarks>
    internal XLFontKey Key
    {
        get
        {
            var pending = _style.Pending;
            return pending is null ? _value.Key : pending.Font;
        }
        private set => _value = XLFontValue.FromKey(ref value);
    }

    #region Constructors

    /// <summary>
    /// Create an instance of XLFont initializing it with the specified value.
    /// </summary>
    /// <param name="style">Style to attach the new instance to.</param>
    /// <param name="value">Style value to use.</param>
    public XLFont(XLStyle? style, XLFontValue value)
    {
        _style = style ?? XLStyle.CreateEmptyStyle();
        _value = value;
    }

    private XLFont(XLStyle? style, XLFontKey key) : this(style, XLFontValue.FromKey(ref key))
    {
    }

    /// <summary>
    /// Create a new font that is attached to a style, and the changes to the font object are propagated to the style.
    /// </summary>
    /// <param name="style">The container style that will be modified by changes of created <c>XLFont</c>.</param>
    public XLFont(XLStyle style) : this(style, GenerateKey(style.Font))
    {
    }

    /// <summary>
    /// Create a new font. The changes to the object are not propagated to a style.
    /// </summary>
    public XLFont(IXLFontBase font) : this(null, GenerateKey(font))
    {
    }

    public XLFont(XLFontKey key) : this(null, XLFontValue.FromKey(ref key))
    {
    }

    private XLFont() : this(null, GenerateKey(null))
    {
    }

    #endregion Constructors

    internal void SyncValue(XLFontValue value)
    {
        _value = value;
    }

    /// <summary>
    /// Apply a new component key to the cell this facade is attached to.
    /// </summary>
    /// <remarks>
    /// The cell-container fast path, so it takes the key directly rather than a
    /// <see cref="Func{T,TResult}"/> delta and allocates no closure. Contrast <see cref="Modify{T}"/>,
    /// which ranges and worksheets need.
    /// <para>
    /// The new key is deliberately <em>not</em> interned before being applied. Assigning it to
    /// <c>Key</c> first would run a repository lookup -- hashing the key and probing a dictionary --
    /// whose result is then thrown away: on a transition-cache hit <c>ModifyFont</c> never needs the
    /// component value at all, and on a miss it interns the component anyway inside
    /// <c>XLStyleValue.FromKey</c>. Taking the interned value back off the resulting style instead
    /// leaves the facade just as correct for later reads, at no lookup. Measured over 20,000 cells
    /// setting one property each, this was the single largest cost on the per-cell styling path.
    /// </para>
    /// </remarks>
    private void SetKey(XLFontKey newKey)
    {
        _style.ModifyFont(newKey);
        _value = _style.Value.Font;
    }

    /// <summary>
    /// The body every setter shares: skip an unchanged value where
    /// <see cref="XLStyle.SkipsUnchangedValues"/> allows it, then take the cell fast path or the
    /// <see cref="Modify{T}"/> path.
    /// </summary>
    /// <param name="unchanged">Whether <paramref name="value"/> equals the one in <see cref="Key"/>.</param>
    /// <param name="value">The value to set.</param>
    /// <param name="with">Rewrites a font key with the value. Pass a static lambda, which is
    /// cached, so the cell fast path allocates nothing.</param>
    /// <remarks>
    /// The value travels as state beside a static lambda rather than captured by a closure. C#
    /// allocates a closure over captured parameters on entry to the method that declares the
    /// lambda, whichever branch then runs, so a setter written as
    /// <c>Modify(k =&gt; k with { Bold = value })</c> paid for it on the cell path as well.
    /// </remarks>
    private void Apply<T>(bool unchanged, T value, Func<XLFontKey, T, XLFontKey> with)
    {
        if (unchanged && _style.SkipsUnchangedValues) return;
        if (_style.IsCellContainer)
            SetKey(with(Key, value));
        else
            Modify(value, with);
    }

    /// <summary>
    /// Non-cell path (ranges, worksheets, conditional formats): apply the delta to each cell's own
    /// font key.
    /// </summary>
    /// <remarks>
    /// A setter skips a value equal to <see cref="Key"/> only where
    /// <see cref="XLStyle.SkipsUnchangedValues"/> allows it: on a cell or a worksheet. On a range or
    /// <c>IXLCells</c> the key is only that container's record of its style, which its cells need
    /// not share: a range is held weakly by its worksheet, and one rebuilt after a collection starts
    /// from its parent's style rather than its cells' (#505). A value equal to that record can still
    /// change a cell, so there the setter always writes.
    /// <para>
    /// Where the facade holds the very value its style does, it does not assign <c>Key</c>, which
    /// would intern the new font in its repository only for the result to be discarded - the same
    /// wasted lookup <see cref="SetKey"/> documents - and would run <paramref name="with"/> an extra
    /// time. The style interns the font again when it resolves its key, so the facade takes the
    /// interned value back off the resulting style instead.
    /// </para>
    /// <para>
    /// Anywhere else the facade's own key stays the source of truth, as it always was. That is a
    /// font built without a style - a rich text run's, which gets an empty style holding the default
    /// font rather than its own - and, rarely, a facade whose style was reset under it, or whose value
    /// the repository handed back as a second instance after a collection. Values are interned, so a
    /// reference test tells the two apart at the cost of one field read and no state of its own: the
    /// facade is allocated per cell, and a flag would grow every one of them.
    /// </para>
    /// <para>
    /// Both lambdas capture the same two parameters, so they share one closure, and only the branch
    /// taken allocates its delegate.
    /// </para>
    /// </remarks>
    private void Modify<T>(T value, Func<XLFontKey, T, XLFontKey> with)
    {
        if (!ReferenceEquals(_value, _style.Value.Font))
        {
            Key = with(Key, value);
            _style.Modify(styleKey => styleKey with { Font = with(styleKey.Font, value) });
            return;
        }

        _style.Modify(styleKey => styleKey with { Font = with(styleKey.Font, value) });
        _value = _style.Value.Font;
    }

    #region IXLFont Members

    public bool Bold
    {
        get => Key.Bold;
        set => Apply(Key.Bold == value, value, static (k, v) => k with { Bold = v });
    }

    public bool Italic
    {
        get => Key.Italic;
        set => Apply(Key.Italic == value, value, static (k, v) => k with { Italic = v });
    }

    public XLFontUnderlineValues Underline
    {
        get => Key.Underline;
        set => Apply(Key.Underline == value, value, static (k, v) => k with { Underline = v });
    }

    public bool Strikethrough
    {
        get => Key.Strikethrough;
        set => Apply(Key.Strikethrough == value, value, static (k, v) => k with { Strikethrough = v });
    }

    public XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => Key.VerticalAlignment;
        set => Apply(Key.VerticalAlignment == value, value, static (k, v) => k with { VerticalAlignment = v });
    }

    public bool Shadow
    {
        get => Key.Shadow;
        set => Apply(Key.Shadow == value, value, static (k, v) => k with { Shadow = v });
    }

    public double FontSize
    {
        get => Key.FontSize;
        set => Apply(XLHelper.AreEqual(Key.FontSize, value), value, static (k, v) => k with { FontSize = v });
    }

    public XLColor FontColor
    {
        get
        {
            var fontColorKey = Key.FontColor;
            return XLColor.FromKey(ref fontColorKey);
        }
        set
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Apply(Key.FontColor == value.Key, value.Key, static (k, v) => k with { FontColor = v });
        }
    }

    public string FontName
    {
        get => Key.FontName;
        set => Apply(Key.FontName == value, value, static (k, v) => k with { FontName = v });
    }

    public XLFontFamilyNumberingValues FontFamilyNumbering
    {
        get => Key.FontFamilyNumbering;
        set => Apply(Key.FontFamilyNumbering == value, value, static (k, v) => k with { FontFamilyNumbering = v });
    }

    public XLFontCharSet FontCharSet
    {
        get => Key.FontCharSet;
        set => Apply(Key.FontCharSet == value, value, static (k, v) => k with { FontCharSet = v });
    }

    public XLFontScheme FontScheme
    {
        get => Key.FontScheme;
        set => Apply(Key.FontScheme == value, value, static (k, v) => k with { FontScheme = v });
    }

    public IXLStyle SetBold()
    {
        Bold = true;
        return _style;
    }

    public IXLStyle SetBold(bool value)
    {
        Bold = value;
        return _style;
    }

    public IXLStyle SetItalic()
    {
        Italic = true;
        return _style;
    }

    public IXLStyle SetItalic(bool value)
    {
        Italic = value;
        return _style;
    }

    public IXLStyle SetUnderline()
    {
        Underline = XLFontUnderlineValues.Single;
        return _style;
    }

    public IXLStyle SetUnderline(XLFontUnderlineValues value)
    {
        Underline = value;
        return _style;
    }

    public IXLStyle SetStrikethrough()
    {
        Strikethrough = true;
        return _style;
    }

    public IXLStyle SetStrikethrough(bool value)
    {
        Strikethrough = value;
        return _style;
    }

    public IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        VerticalAlignment = value;
        return _style;
    }

    public IXLStyle SetShadow()
    {
        Shadow = true;
        return _style;
    }

    public IXLStyle SetShadow(bool value)
    {
        Shadow = value;
        return _style;
    }

    public IXLStyle SetFontSize(double value)
    {
        FontSize = value;
        return _style;
    }

    public IXLStyle SetFontColor(XLColor value)
    {
        FontColor = value;
        return _style;
    }

    public IXLStyle SetFontName(string value)
    {
        FontName = value;
        return _style;
    }

    public IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        FontFamilyNumbering = value;
        return _style;
    }

    public IXLStyle SetFontCharSet(XLFontCharSet value)
    {
        FontCharSet = value;
        return _style;
    }

    public IXLStyle SetFontScheme(XLFontScheme value)
    {
        FontScheme = value;
        return _style;
    }

    #endregion IXLFont Members

    #region Overridden

    public override string ToString()
    {
        var sb = new StringBuilder();
        sb.Append(Bold);
        sb.Append('-');
        sb.Append(Italic);
        sb.Append('-');
        sb.Append(Underline.ToString());
        sb.Append('-');
        sb.Append(Strikethrough);
        sb.Append('-');
        sb.Append(VerticalAlignment.ToString());
        sb.Append('-');
        sb.Append(Shadow);
        sb.Append('-');
        sb.Append(FontSize.ToString(CultureInfo.InvariantCulture));
        sb.Append('-');
        sb.Append(FontColor);
        sb.Append('-');
        sb.Append(FontName);
        sb.Append('-');
        sb.Append(FontFamilyNumbering.ToString());
        sb.Append('-');
        sb.Append(FontCharSet.ToString());
        sb.Append('-');
        sb.Append(FontScheme.ToString());
        return sb.ToString();
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as XLFont);
    }

    public bool Equals(IXLFont? other)
    {
        if (other is not XLFont otherF)
            return false;

        return Key == otherF.Key;
    }

    public override int GetHashCode()
    {
        var hashCode = 416600561;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden
}
