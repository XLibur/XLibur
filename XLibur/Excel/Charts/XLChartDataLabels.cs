using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Identifies the data label properties a caller has explicitly assigned, so that the patcher which
/// updates charts loaded from a file can rewrite those and only those.
/// </summary>
[Flags]
internal enum XLDataLabelsFormat
{
    None = 0,
    ShowValue = 1 << 0,
    ShowCategoryName = 1 << 1,
    ShowSeriesName = 1 << 2,
    ShowPercentage = 1 << 3,
    NumberFormat = 1 << 4,
    Position = 1 << 5
}

internal sealed class XLChartDataLabels : IXLDataLabels
{
    private static readonly XLDataLabelPosition[] NoPosition = [];

    private static readonly XLDataLabelPosition[] ClusteredBarPositions =
    [
        XLDataLabelPosition.Center, XLDataLabelPosition.InsideEnd,
        XLDataLabelPosition.InsideBase, XLDataLabelPosition.OutsideEnd
    ];

    private static readonly XLDataLabelPosition[] StackedBarPositions =
    [
        XLDataLabelPosition.Center, XLDataLabelPosition.InsideEnd, XLDataLabelPosition.InsideBase
    ];

    private static readonly XLDataLabelPosition[] MarkerPositions =
    [
        XLDataLabelPosition.Center, XLDataLabelPosition.Left,
        XLDataLabelPosition.Right, XLDataLabelPosition.Above, XLDataLabelPosition.Below
    ];

    private static readonly XLDataLabelPosition[] PiePositions =
    [
        XLDataLabelPosition.BestFit, XLDataLabelPosition.Center,
        XLDataLabelPosition.InsideEnd, XLDataLabelPosition.OutsideEnd
    ];

    private readonly XLChart _chart;

    /// <summary>
    /// Whether these labels belong to a series of the combo chart's secondary type, whose label
    /// positions are governed by <see cref="IXLChart.SecondaryChartType"/> rather than
    /// <see cref="IXLChart.ChartType"/>.
    /// </summary>
    private readonly bool _secondary;

    private bool _showValue;
    private bool _showCategoryName;
    private bool _showSeriesName;
    private bool _showPercentage;
    private string? _numberFormat;
    private XLDataLabelPosition _position;

    internal XLChartDataLabels(XLChart chart, bool secondary = false)
    {
        _chart = chart;
        _secondary = secondary;
    }

    /// <summary>The chart type whose rules apply to these labels.</summary>
    private XLChartType OwningChartType => _secondary
        ? _chart.SecondaryChartType ?? _chart.ChartType
        : _chart.ChartType;

    public bool ShowValue
    {
        get => _showValue;
        set => Assign(ref _showValue, value, XLDataLabelsFormat.ShowValue);
    }

    public bool ShowCategoryName
    {
        get => _showCategoryName;
        set => Assign(ref _showCategoryName, value, XLDataLabelsFormat.ShowCategoryName);
    }

    public bool ShowSeriesName
    {
        get => _showSeriesName;
        set => Assign(ref _showSeriesName, value, XLDataLabelsFormat.ShowSeriesName);
    }

    public bool ShowPercentage
    {
        get => _showPercentage;
        set => Assign(ref _showPercentage, value, XLDataLabelsFormat.ShowPercentage);
    }

    public string? NumberFormat
    {
        get => _numberFormat;
        set => Assign(ref _numberFormat, value, XLDataLabelsFormat.NumberFormat);
    }

    public XLDataLabelPosition Position
    {
        get => _position;
        set
        {
            var chartType = OwningChartType;
            if (value != XLDataLabelPosition.Auto && !IsPositionAllowed(chartType, value))
            {
                var allowed = AllowedPositions(chartType);
                var offered = allowed.Length == 0
                    ? "only Auto"
                    : "Auto, " + string.Join(", ", allowed);
                throw new ArgumentException(
                    $"A {chartType} chart does not offer the {value} data label position; " +
                    $"Excel accepts {offered}.",
                    nameof(Position));
            }

            Assign(ref _position, value, XLDataLabelsFormat.Position);
        }
    }

    /// <summary>
    /// The properties that have been explicitly assigned through the public API. Nothing is written
    /// to the file while this is <see cref="XLDataLabelsFormat.None"/>.
    /// </summary>
    internal XLDataLabelsFormat AssignedFormat { get; private set; }

    /// <summary>
    /// Seeds the properties from an existing chart part without marking them as assigned, so that
    /// labels nobody edited are never written back.
    /// </summary>
    internal void SeedLoaded(
        bool showValue,
        bool showCategoryName,
        bool showSeriesName,
        bool showPercentage,
        string? numberFormat,
        XLDataLabelPosition position)
    {
        _showValue = showValue;
        _showCategoryName = showCategoryName;
        _showSeriesName = showSeriesName;
        _showPercentage = showPercentage;
        _numberFormat = numberFormat;
        _position = position;
    }

    /// <summary>
    /// The position to write for a chart of the given type. A position the type does not offer is
    /// dropped rather than written, because Excel refuses to open a file that uses one — the setter
    /// rejects those, but the chart type can be changed after the position was set.
    /// </summary>
    internal XLDataLabelPosition EffectivePosition(XLChartType chartType) =>
        IsPositionAllowed(chartType, _position) ? _position : XLDataLabelPosition.Auto;

    private static bool IsPositionAllowed(XLChartType chartType, XLDataLabelPosition position) =>
        position == XLDataLabelPosition.Auto || AllowedPositions(chartType).Contains(position);

    /// <summary>
    /// The explicit label positions Excel offers for a chart type, excluding
    /// <see cref="XLDataLabelPosition.Auto"/>, which is always allowed. Area, doughnut, bubble, stock,
    /// surface and every 3D type offer none: Excel places their labels itself and rejects a file that
    /// says otherwise.
    /// </summary>
    private static XLDataLabelPosition[] AllowedPositions(XLChartType chartType) => chartType switch
    {
        XLChartType.BarClustered or XLChartType.ColumnClustered => ClusteredBarPositions,

        XLChartType.BarStacked or XLChartType.BarStacked100Percent
            or XLChartType.ColumnStacked or XLChartType.ColumnStacked100Percent => StackedBarPositions,

        XLChartType.Line or XLChartType.LineStacked or XLChartType.LineStacked100Percent
            or XLChartType.LineWithMarkers or XLChartType.LineWithMarkersStacked
            or XLChartType.LineWithMarkersStacked100Percent
            or XLChartType.Radar or XLChartType.RadarWithMarkers or XLChartType.RadarFilled
            or XLChartType.XYScatterMarkers
            or XLChartType.XYScatterSmoothLinesNoMarkers or XLChartType.XYScatterSmoothLinesWithMarkers
            or XLChartType.XYScatterStraightLinesNoMarkers
            or XLChartType.XYScatterStraightLinesWithMarkers => MarkerPositions,

        XLChartType.Pie or XLChartType.PieExploded
            or XLChartType.PieToPie or XLChartType.PieToBar => PiePositions,

        _ => NoPosition
    };

    private void Assign<T>(ref T field, T value, XLDataLabelsFormat flag)
    {
        field = value;
        AssignedFormat |= flag;
    }
}
