using System;

namespace XLibur.Excel;

/// <summary>
/// Identifies the legend properties a caller has explicitly assigned.
/// </summary>
[Flags]
internal enum XLChartLegendFormat
{
    None = 0,
    Visible = 1 << 0,
    Position = 1 << 1,
    Overlay = 1 << 2
}

internal sealed class XLChartLegend : IXLChartLegend
{
    private bool _visible;
    private XLLegendPosition _position;
    private bool _overlay;

    public bool Visible
    {
        get => _visible;
        set => Assign(ref _visible, value, XLChartLegendFormat.Visible);
    }

    public XLLegendPosition Position
    {
        get => _position;
        set => Assign(ref _position, value, XLChartLegendFormat.Position);
    }

    public bool Overlay
    {
        get => _overlay;
        set => Assign(ref _overlay, value, XLChartLegendFormat.Overlay);
    }

    /// <summary>
    /// The properties that have been explicitly assigned through the public API.
    /// </summary>
    internal XLChartLegendFormat AssignedFormat { get; private set; }

    /// <summary>
    /// Seeds the properties from an existing chart part without marking them as assigned, so that a
    /// legend nobody edited is never written back.
    /// </summary>
    internal void SeedLoaded(bool visible, XLLegendPosition position, bool overlay)
    {
        _visible = visible;
        _position = position;
        _overlay = overlay;
    }

    private void Assign<T>(ref T field, T value, XLChartLegendFormat flag)
    {
        field = value;
        AssignedFormat |= flag;
    }
}
