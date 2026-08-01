using System;

namespace XLibur.Excel;

/// <summary>
/// Which axis of a chart an <see cref="XLChartAxis"/> is, which decides both the properties the file
/// format lets it carry and where the patcher finds it.
/// </summary>
internal enum XLChartAxisRole
{
    /// <summary>The horizontal axis — a category axis, or the X value axis of a scatter chart.</summary>
    Category,

    /// <summary>The vertical value axis.</summary>
    Value,

    /// <summary>The value axis on the right, used by series bound to the secondary axis.</summary>
    SecondaryValue
}

/// <summary>
/// Identifies the axis properties a caller has explicitly assigned, so that the patcher which updates
/// charts loaded from a file can rewrite those and only those.
/// </summary>
[Flags]
internal enum XLChartAxisFormat
{
    None = 0,
    Title = 1 << 0,
    NumberFormat = 1 << 1,
    Min = 1 << 2,
    Max = 1 << 3,
    MajorUnit = 1 << 4,
    MinorUnit = 1 << 5,
    Visible = 1 << 6,
    MajorGridlines = 1 << 7,
    Orientation = 1 << 8,
    LogScale = 1 << 9,
    LogBase = 1 << 10
}

internal sealed class XLChartAxis : IXLChartAxis
{
    private const double MinLogBase = 2;
    private const double MaxLogBase = 1000;

    private string? _title;
    private string? _numberFormat;
    private double? _min;
    private double? _max;
    private double? _majorUnit;
    private double? _minorUnit;
    private bool _visible = true;
    private bool _majorGridlines;
    private XLAxisOrientation _orientation;
    private bool _logScale;
    private double _logBase = 10;

    private readonly XLChart _chart;

    internal XLChartAxis(XLChart chart, XLChartAxisRole role)
    {
        _chart = chart;
        Role = role;
    }

    internal XLChartAxisRole Role { get; }

    /// <summary>
    /// Whether this axis can carry the value-axis-only properties: the unit intervals and the
    /// logarithmic scale. True for the vertical axes always, and for the horizontal axis of the chart
    /// types that plot numbers along it — scatter and bubble, where <c>c:catAx</c> is really a
    /// <c>c:valAx</c>.
    /// </summary>
    internal bool IsValueAxis => Role != XLChartAxisRole.Category || HasNumericHorizontalAxis(_chart.ChartType);

    private static bool HasNumericHorizontalAxis(XLChartType chartType) => chartType
        is XLChartType.XYScatterMarkers
        or XLChartType.XYScatterSmoothLinesNoMarkers or XLChartType.XYScatterSmoothLinesWithMarkers
        or XLChartType.XYScatterStraightLinesNoMarkers or XLChartType.XYScatterStraightLinesWithMarkers
        or XLChartType.Bubble or XLChartType.Bubble3D;

    public string? Title
    {
        get => _title;
        set => Assign(ref _title, value, XLChartAxisFormat.Title);
    }

    public string? NumberFormat
    {
        get => _numberFormat;
        set => Assign(ref _numberFormat, value, XLChartAxisFormat.NumberFormat);
    }

    public double? Min
    {
        get => _min;
        set => Assign(ref _min, value, XLChartAxisFormat.Min);
    }

    public double? Max
    {
        get => _max;
        set => Assign(ref _max, value, XLChartAxisFormat.Max);
    }

    public double? MajorUnit
    {
        get => _majorUnit;
        set
        {
            if (value <= 0)
                throw new ArgumentOutOfRangeException(nameof(MajorUnit), value,
                    "An axis unit must be greater than zero.");

            Assign(ref _majorUnit, value, XLChartAxisFormat.MajorUnit);
        }
    }

    public double? MinorUnit
    {
        get => _minorUnit;
        set
        {
            if (value <= 0)
                throw new ArgumentOutOfRangeException(nameof(MinorUnit), value,
                    "An axis unit must be greater than zero.");

            Assign(ref _minorUnit, value, XLChartAxisFormat.MinorUnit);
        }
    }

    public bool Visible
    {
        get => _visible;
        set => Assign(ref _visible, value, XLChartAxisFormat.Visible);
    }

    public bool MajorGridlines
    {
        get => _majorGridlines;
        set => Assign(ref _majorGridlines, value, XLChartAxisFormat.MajorGridlines);
    }

    public XLAxisOrientation Orientation
    {
        get => _orientation;
        set => Assign(ref _orientation, value, XLChartAxisFormat.Orientation);
    }

    public bool LogScale
    {
        get => _logScale;
        set => Assign(ref _logScale, value, XLChartAxisFormat.LogScale);
    }

    public double LogBase
    {
        get => _logBase;
        set
        {
            if (value is < MinLogBase or > MaxLogBase)
                throw new ArgumentOutOfRangeException(nameof(LogBase), value,
                    $"The base of a logarithmic axis must be between {MinLogBase} and {MaxLogBase}.");

            Assign(ref _logBase, value, XLChartAxisFormat.LogBase);
        }
    }

    /// <summary>
    /// The properties that have been explicitly assigned through the public API.
    /// </summary>
    internal XLChartAxisFormat AssignedFormat { get; private set; }

    /// <summary>
    /// Seeds the properties from an existing chart part without marking them as assigned, so that an
    /// axis nobody edited is never written back.
    /// </summary>
#pragma warning disable S107 // One parameter per axis attribute read from the chart part
    internal void SeedLoaded(
        string? title,
        string? numberFormat,
        double? min,
        double? max,
        double? majorUnit,
        double? minorUnit,
        bool visible,
        bool majorGridlines,
        XLAxisOrientation orientation,
        bool logScale,
        double logBase)
#pragma warning restore S107
    {
        _title = title;
        _numberFormat = numberFormat;
        _min = min;
        _max = max;
        _majorUnit = majorUnit;
        _minorUnit = minorUnit;
        _visible = visible;
        _majorGridlines = majorGridlines;
        _orientation = orientation;
        _logScale = logScale;
        _logBase = logBase;
    }

    private void Assign<T>(ref T field, T value, XLChartAxisFormat flag)
    {
        field = value;
        AssignedFormat |= flag;
    }
}
