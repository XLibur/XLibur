namespace XLibur.Excel;

/// <summary>
/// The two things every range type needs at construction: where it is, and the style it starts
/// from.
/// </summary>
/// <remarks>
/// A <see langword="struct"/> rather than a class: it is a pure parameter object, it is never
/// mutated after construction, nothing compares or stores it by identity, and it exists only to
/// carry two values from <see cref="XLRangeFactory"/> into a range constructor. As a class it cost
/// one heap allocation per range construction — including on the repository-cache-hit path through
/// <see cref="XLWorksheet.GetOrCreateRange"/>, which builds one before it looks anything up.
/// <para>
/// Pass it by <see langword="in"/> on any new call site. It is 8 bytes of reference plus an
/// <see cref="XLRangeAddress"/>, so a copy is not free.
/// </para>
/// </remarks>
internal readonly struct XLRangeParameters
{
    #region Constructor

    public XLRangeParameters(XLRangeAddress rangeAddress, IXLStyle defaultStyle)
    {
        RangeAddress = rangeAddress;
        DefaultStyle = defaultStyle;
    }

    #endregion

    #region Properties

    public XLRangeAddress RangeAddress { get; }

    public IXLStyle DefaultStyle { get; }
    #endregion
}
