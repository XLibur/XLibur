using XLibur.Excel.Coordinates;

namespace XLibur.Excel.Ranges;

internal static class XLAddressableHelper
{
    /// <summary>
    /// Whether <paramref name="addressable"/> covers <paramref name="address"/>, without boxing.
    /// </summary>
    /// <remarks>
    /// <see cref="IXLAddressable.RangeAddress"/> is typed as the <see cref="IXLRangeAddress"/>
    /// interface, but <see cref="XLRangeAddress"/> is a struct — so reading the property through
    /// the interface boxes it, and passing an <see cref="XLAddress"/> to
    /// <see cref="IXLRangeAddress.Contains(IXLAddress)"/> boxes again. Two allocations per range
    /// per test is invisible in a one-off call and very visible in the merged-range check that
    /// runs on every cell write. Every addressable in the range indexes is an
    /// <see cref="XLRangeBase"/>, whose <c>RangeAddress</c> is the concrete struct; the interface
    /// path is kept only as a fallback for any implementation that is not.
    /// <see cref="IXLAddressable"/> is public, so it cannot simply grow a non-boxing member.
    /// <para>
    /// Reads through <see cref="Area"/> rather than calling <c>RangeAddress.Contains</c>
    /// directly: <c>XLRangeAddress.Contains(in XLAddress)</c> assumes <c>FirstAddress</c> is the
    /// top-left corner, which does not hold for a range whose corners were given in reverse
    /// order (a later <c>AddRange</c> on an existing merge or data validation, for example - its
    /// first area is normalised on entry, but nothing normalises one added afterwards). The
    /// non-generic <see cref="Area.FromRangeAddress(XLRangeAddress)"/> overload keeps the
    /// concrete-struct branch non-boxing; the interface fallback already boxed before this
    /// change and still does.
    /// </para>
    /// </remarks>
    internal static bool Contains(IXLAddressable addressable, in XLAddress address)
    {
        var point = new Point(address.RowNumber, address.ColumnNumber);

        if (addressable is XLRangeBase rangeBase)
            return Area.FromRangeAddress(rangeBase.RangeAddress).Contains(point);

        return Area.FromRangeAddress(addressable.RangeAddress).Contains(point);
    }
}
