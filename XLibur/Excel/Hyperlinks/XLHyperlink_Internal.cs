using System;

namespace XLibur.Excel;

public partial class XLHyperlink
{
    internal XLHyperlink()
    {

    }

    internal XLHyperlink(XLHyperlink hyperlink)
    {
        _externalAddress = hyperlink._externalAddress;
        _internalAddress = hyperlink._internalAddress;
        Tooltip = hyperlink.Tooltip;
        IsExternal = hyperlink.IsExternal;
    }

    internal void SetValues(string address, string tooltip)
    {
        Tooltip = tooltip;
        if (address[0] == '.')
        {
            _externalAddress = new Uri(address, UriKind.Relative);
            IsExternal = true;
        }
        else
        {
            if (Uri.TryCreate(address, UriKind.Absolute, out Uri? uri))
            {
                _externalAddress = uri;
                IsExternal = true;
            }
            else
            {
                _internalAddress = address;
                IsExternal = false;
            }
        }
    }

    internal void SetValues(Uri uri, string tooltip)
    {
        Tooltip = tooltip;
        _externalAddress = uri;
        IsExternal = true;
    }

    internal void SetValues(IXLCell cell, string tooltip)
    {
        Tooltip = tooltip;
        _internalAddress = cell.Address.ToString(XLReferenceStyle.A1, true);
        IsExternal = false;
    }

    internal void SetValues(IXLRangeBase range, string tooltip)
    {
        Tooltip = tooltip;
        _internalAddress = range.RangeAddress.ToString(XLReferenceStyle.A1, true);
        IsExternal = false;
    }

    internal XLHyperlinks? Container { get; set; }

    /// <summary>
    /// The address <see cref="_relId"/> was written or loaded for. The relationship ID is only valid
    /// while the hyperlink still points at this exact <see cref="Uri"/>.
    /// </summary>
    private Uri? _relIdAddress;

    private string? _relId;

    /// <summary>
    /// The ID of the worksheet relationship that holds the external address, as loaded or last
    /// saved. <c>null</c> when the hyperlink is internal, new, or its address has changed since.
    /// Reusing it keeps the saved output stable across a load and save.
    /// </summary>
    internal string? RelId
    {
        get => IsExternal && _relId is not null && ReferenceEquals(_externalAddress, _relIdAddress) ? _relId : null;
        set
        {
            _relId = value;
            _relIdAddress = value is null ? null : _externalAddress;
        }
    }
}
