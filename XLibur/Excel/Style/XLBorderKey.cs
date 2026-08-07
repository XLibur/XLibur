namespace XLibur.Excel;

/// <summary>
/// The composite key of a border: a style and a colour for each of the five edges, plus the two
/// diagonal direction flags.
/// </summary>
/// <remarks>
/// The five colours dominate the size of this struct, and this struct dominates the size of
/// <c>XLStyleKey</c>, which is copied on every style mutation and again on every repository probe.
/// The edge styles are therefore stored as bytes - <see cref="XLBorderStyleValues"/> has fourteen
/// members but is public, so it stays <c>int</c>-backed and widens at the property boundary.
/// <para>
/// Keys reaching a repository or an <c>XLStyleKey</c> are <see cref="Normalize">normalized</see>
/// first, so <see cref="Equals(XLBorderKey)"/> is a plain field comparison rather than the
/// edge-by-edge equivalence test it used to be.
/// </para>
/// </remarks>
internal readonly record struct XLBorderKey
{
    /// <summary>
    /// The colour every edge without a style is given by <see cref="Normalize"/>. Black, because
    /// that is what a border key with no explicit colour has always carried, and what reading the
    /// colour of a styleless edge has always returned.
    /// </summary>
    /// <remarks>
    /// Built from its ARGB value rather than <c>XLColor.Black.Key</c> so that normalizing does not
    /// drag the <see cref="XLColor"/> static tables into a static initialisation order it does not
    /// otherwise participate in. The two are the same colour.
    /// </remarks>
    private static readonly XLColorKey StylelessEdgeColor = XLColorKey.FromArgb(0xFF000000);

    // Every backing field is declared here rather than left implicit on the properties, because the
    // runtime lays a struct's fields out in declaration order where alignment allows and does not
    // reorder across the eight-byte-aligned colours. Interleaved with the properties - a style byte,
    // then a colour, then the next style byte - the five bytes land before the colours and the two
    // bools after them, and the tail pads the struct out to 96. Grouped, all seven scalars share the
    // colours' leading pad word and the struct is 88.
    private readonly byte _leftBorder;
    private readonly byte _rightBorder;
    private readonly byte _topBorder;
    private readonly byte _bottomBorder;
    private readonly byte _diagonalBorder;
    private readonly bool _diagonalUp;
    private readonly bool _diagonalDown;

    private readonly XLColorKey _leftBorderColor;
    private readonly XLColorKey _rightBorderColor;
    private readonly XLColorKey _topBorderColor;
    private readonly XLColorKey _bottomBorderColor;
    private readonly XLColorKey _diagonalBorderColor;

    public required XLBorderStyleValues LeftBorder
    {
        get => (XLBorderStyleValues)_leftBorder;
        init => _leftBorder = (byte)value;
    }

    public required XLColorKey LeftBorderColor
    {
        get => _leftBorderColor;
        init => _leftBorderColor = value;
    }

    public required XLBorderStyleValues RightBorder
    {
        get => (XLBorderStyleValues)_rightBorder;
        init => _rightBorder = (byte)value;
    }

    public required XLColorKey RightBorderColor
    {
        get => _rightBorderColor;
        init => _rightBorderColor = value;
    }

    public required XLBorderStyleValues TopBorder
    {
        get => (XLBorderStyleValues)_topBorder;
        init => _topBorder = (byte)value;
    }

    public required XLColorKey TopBorderColor
    {
        get => _topBorderColor;
        init => _topBorderColor = value;
    }

    public required XLBorderStyleValues BottomBorder
    {
        get => (XLBorderStyleValues)_bottomBorder;
        init => _bottomBorder = (byte)value;
    }

    public required XLColorKey BottomBorderColor
    {
        get => _bottomBorderColor;
        init => _bottomBorderColor = value;
    }

    public required XLBorderStyleValues DiagonalBorder
    {
        get => (XLBorderStyleValues)_diagonalBorder;
        init => _diagonalBorder = (byte)value;
    }

    public required XLColorKey DiagonalBorderColor
    {
        get => _diagonalBorderColor;
        init => _diagonalBorderColor = value;
    }

    public required bool DiagonalUp
    {
        get => _diagonalUp;
        init => _diagonalUp = value;
    }

    public required bool DiagonalDown
    {
        get => _diagonalDown;
        init => _diagonalDown = value;
    }

    /// <summary>
    /// Replace the colour of every edge whose style is <see cref="XLBorderStyleValues.None"/> with
    /// <see cref="StylelessEdgeColor"/>, returning the key unchanged if there is nothing to replace.
    /// </summary>
    /// <remarks>
    /// An edge with no style has no colour to draw with, and Excel writes none. Until this ran, the
    /// colour set on such an edge was still carried in the key, and which of several equivalent keys
    /// a repository ended up interning - and so which colour reading the edge back reported - was
    /// decided by whichever was stored first. Collapsing them to one key makes that read
    /// deterministic and lets equality compare fields directly.
    /// <para>
    /// Applied where a key enters a repository or an <c>XLStyleKey</c>, not in the <c>init</c>
    /// accessors: an object initialiser may set an edge's colour before its style, and an accessor
    /// clearing the colour against a style that has not been assigned yet would discard a colour the
    /// caller did set.
    /// </para>
    /// </remarks>
    internal XLBorderKey Normalize()
    {
        if (!NeedsNormalizing())
            return this;

        return new XLBorderKey
        {
            LeftBorder = LeftBorder,
            LeftBorderColor = ColorOf(LeftBorder, LeftBorderColor),
            RightBorder = RightBorder,
            RightBorderColor = ColorOf(RightBorder, RightBorderColor),
            TopBorder = TopBorder,
            TopBorderColor = ColorOf(TopBorder, TopBorderColor),
            BottomBorder = BottomBorder,
            BottomBorderColor = ColorOf(BottomBorder, BottomBorderColor),
            DiagonalBorder = DiagonalBorder,
            DiagonalBorderColor = ColorOf(DiagonalBorder, DiagonalBorderColor),
            DiagonalUp = DiagonalUp,
            DiagonalDown = DiagonalDown,
        };

        static XLColorKey ColorOf(XLBorderStyleValues style, XLColorKey color) =>
            style == XLBorderStyleValues.None ? StylelessEdgeColor : color;
    }

    /// <summary>
    /// Whether any edge carries a colour <see cref="Normalize"/> would drop. Checked first so the
    /// overwhelmingly common already-normal key is returned without being rebuilt.
    /// </summary>
    private bool NeedsNormalizing()
    {
        return IsStylelessWithColor(_leftBorder, LeftBorderColor)
               || IsStylelessWithColor(_rightBorder, RightBorderColor)
               || IsStylelessWithColor(_topBorder, TopBorderColor)
               || IsStylelessWithColor(_bottomBorder, BottomBorderColor)
               || IsStylelessWithColor(_diagonalBorder, DiagonalBorderColor);

        static bool IsStylelessWithColor(byte style, XLColorKey color) =>
            style == (byte)XLBorderStyleValues.None && !color.Equals(StylelessEdgeColor);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = (int)_leftBorder;
            hash = (hash * 397) ^ LeftBorderColor.GetHashCode();

            hash = (hash * 397) ^ _rightBorder;
            hash = (hash * 397) ^ RightBorderColor.GetHashCode();

            hash = (hash * 397) ^ _topBorder;
            hash = (hash * 397) ^ TopBorderColor.GetHashCode();

            hash = (hash * 397) ^ _bottomBorder;
            hash = (hash * 397) ^ BottomBorderColor.GetHashCode();

            hash = (hash * 397) ^ _diagonalBorder;
            hash = (hash * 397) ^ DiagonalBorderColor.GetHashCode();

            hash = (hash * 397) ^ (DiagonalUp ? 1 : 0);
            hash = (hash * 397) ^ (DiagonalDown ? 1 : 0);

            return hash;
        }
    }

    public bool Equals(XLBorderKey other)
    {
        // Scalars first: they reject most non-matching keys for the price of a few byte compares,
        // before any colour is touched.
        return _leftBorder == other._leftBorder
               && _rightBorder == other._rightBorder
               && _topBorder == other._topBorder
               && _bottomBorder == other._bottomBorder
               && _diagonalBorder == other._diagonalBorder
               && _diagonalUp == other._diagonalUp
               && _diagonalDown == other._diagonalDown
               && _leftBorderColor.Equals(other._leftBorderColor)
               && _rightBorderColor.Equals(other._rightBorderColor)
               && _topBorderColor.Equals(other._topBorderColor)
               && _bottomBorderColor.Equals(other._bottomBorderColor)
               && _diagonalBorderColor.Equals(other._diagonalBorderColor);
    }

    public override string ToString()
    {
        return $"{LeftBorder} {LeftBorderColor} {RightBorder} {RightBorderColor} {TopBorder} {TopBorderColor} " +
               $"{BottomBorder} {BottomBorderColor} {DiagonalBorder} {DiagonalBorderColor} " +
               (DiagonalUp ? "DiagonalUp" : "") +
               (DiagonalDown ? "DiagonalDown" : "");
    }
}
