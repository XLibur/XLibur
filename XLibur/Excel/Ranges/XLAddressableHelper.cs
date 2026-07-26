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
    /// </remarks>
    internal static bool Contains(IXLAddressable addressable, in XLAddress address)
    {
        if (addressable is XLRangeBase rangeBase)
            return rangeBase.RangeAddress.Contains(in address);

        return addressable.RangeAddress.Contains(address);
    }
}
