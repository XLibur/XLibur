using System;
using XLibur.Extensions;

namespace XLibur.Excel;

public partial class XLHyperlink
{
    private Uri? _externalAddress;
    private string? _internalAddress;

    public XLHyperlink(string address)
    {
        SetValues(address, string.Empty);
    }

    public XLHyperlink(string address, string tooltip)
    {
        SetValues(address, tooltip);
    }

    public XLHyperlink(IXLCell cell)
    {
        SetValues(cell, string.Empty);
    }

    public XLHyperlink(IXLCell cell, string tooltip)
    {
        SetValues(cell, tooltip);
    }

    public XLHyperlink(IXLRangeBase range)
    {
        SetValues(range, string.Empty);
    }

    public XLHyperlink(IXLRangeBase range, string tooltip)
    {
        SetValues(range, tooltip);
    }

    public XLHyperlink(Uri uri)
    {
        SetValues(uri, string.Empty);
    }

    public XLHyperlink(Uri uri, string tooltip)
    {
        SetValues(uri, tooltip);
    }

    public bool IsExternal { get; set; }

    public Uri? ExternalAddress
    {
        get
        {
            return IsExternal ? _externalAddress : null;
        }
        set
        {
            _externalAddress = value;
            IsExternal = true;
        }
    }

    /// <summary>
    /// Gets top left cell of a hyperlink range. Return <c>null</c>,
    /// if the hyperlink isn't in a worksheet.
    /// </summary>
    public IXLCell? Cell
    {
        get
        {
            if (Container is null)
                return null;

            return Container.GetCell(this);
        }
    }

    public string? InternalAddress
    {
        get
        {
            if (IsExternal)
                return null;
            if (_internalAddress!.Contains('!'))
            {
                return _internalAddress[0] != '\''
                    ? string.Concat(
                        _internalAddress
                            .Substring(0, _internalAddress.IndexOf('!'))
                            .EscapeSheetName(),
                        '!',
                        _internalAddress.Substring(_internalAddress.IndexOf('!') + 1))
                    : _internalAddress;
            }

            if (Container is null)
                throw new InvalidOperationException("Hyperlink is not attached to a worksheet.");

            var sheetName = Container.WorksheetName;
            return string.Concat(
                sheetName.EscapeSheetName(),
                '!',
                _internalAddress);
        }
        set
        {
            _internalAddress = value;
            IsExternal = false;
        }
    }

    /// <summary>
    /// Tooltip displayed when user hovers over the hyperlink range. If not specified,
    /// the link target is displayed in the tooltip.
    /// </summary>
    public string Tooltip { get; set; } = string.Empty;

    /// <inheritdoc cref="IXLHyperlinks.Delete(XLHyperlink)"/>
    public void Delete()
    {
        Container?.Delete(this);
    }
}
