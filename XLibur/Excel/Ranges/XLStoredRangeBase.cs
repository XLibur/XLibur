using XLibur.Excel.Rows;

namespace XLibur.Excel;

/// <summary>
/// Base class for range types that store their <see cref="XLRangeAddress"/> directly.
/// <see cref="XLRow"/> and <see cref="XLColumn"/> compute the address from a stored
/// row/column number instead, avoiding the 48-byte <see cref="XLRangeAddress"/> overhead per instance.
/// </summary>
internal abstract class XLStoredRangeBase : XLRangeBase
{
    private XLRangeAddress _rangeAddress;

    protected XLStoredRangeBase(XLRangeAddress rangeAddress, XLStyleValue styleValue)
        : base(styleValue)
    {
        _rangeAddress = rangeAddress;
    }

    public sealed override XLRangeAddress RangeAddress
    {
        get => _rangeAddress;
        protected set
        {
            if (_rangeAddress == value) return;
            var oldAddress = _rangeAddress;
            _rangeAddress = value;
            OnRangeAddressChanged(oldAddress, _rangeAddress);
        }
    }
}
