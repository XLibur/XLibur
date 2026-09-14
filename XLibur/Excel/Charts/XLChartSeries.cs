using System;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Identifies the formatting properties a caller has explicitly assigned. The writer ignores this
/// (a <c>null</c> property simply omits its element), but the patcher that updates charts loaded
/// from a file needs to tell "never assigned" from "assigned to <c>null</c>" so that it only
/// rewrites the parts of the chart XML the caller actually asked to change.
/// </summary>
[Flags]
internal enum XLChartSeriesFormat
{
    None = 0,
    Fill = 1 << 0,
    Line = 1 << 1,
    LineWidth = 1 << 2,
    Marker = 1 << 3,
    MarkerSize = 1 << 4,
    MarkerFill = 1 << 5,
    Smooth = 1 << 6,
    ValueReferences = 1 << 7,
    CategoryReferences = 1 << 8,
    NameReference = 1 << 9,

    /// <summary>
    /// The value references were rewritten by a sheet rename or delete, not re-pointed by the caller.
    /// The patcher writes them, and keeps the cached values that go with them.
    /// </summary>
    ValueReferencesRewritten = 1 << 10,

    /// <summary>As <see cref="ValueReferencesRewritten"/>, for the category references.</summary>
    CategoryReferencesRewritten = 1 << 11
}

internal sealed class XLChartSeries : IXLChartSeries
{
    /// <summary>The largest line width, in points, that Excel accepts (1584 pt = 20116800 EMU).</summary>
    private const double MaxLineWidthPt = 1584;

    private const double MinMarkerSize = 2;
    private const double MaxMarkerSize = 72;

    private readonly XLChart _chart;

    private XLColor? _fillColor;
    private XLColor? _lineColor;
    private double? _lineWidthPt;
    private XLMarkerStyle _markerStyle;
    private double? _markerSize;
    private XLColor? _markerFillColor;
    private bool _smooth;
    private bool _useSecondaryAxis;
    private string _valueReferences = string.Empty;
    private string? _categoryReferences;
    private string? _nameReference;

    /// <param name="chart">The chart this series belongs to.</param>
    /// <param name="secondary">
    /// <c>true</c> when the series sits in the combo chart's secondary collection, which decides
    /// which chart type's rules its data labels follow.
    /// </param>
    internal XLChartSeries(XLChart chart, bool secondary)
    {
        _chart = chart;
        DataLabelsInternal = new XLChartDataLabels(chart, secondary);
    }

    public string Name { get; set; } = string.Empty;

    public string? CategoryReferences
    {
        get => _categoryReferences;
        set => Assign(ref _categoryReferences, value, XLChartSeriesFormat.CategoryReferences);
    }

    public string ValueReferences
    {
        get => _valueReferences;
        set => Assign(ref _valueReferences, value ?? string.Empty, XLChartSeriesFormat.ValueReferences);
    }

    /// <summary>
    /// The cell the series takes its name from, <c>Sheet1!$D$1</c>, as Excel writes it in
    /// <c>c:tx/c:strRef/c:f</c>; <c>null</c> when the name is literal text, as it is for every series
    /// XLibur creates. Internal: it is read from a loaded chart, and a sheet rename or delete rewrites
    /// it (see <see cref="XLCharts"/>).
    /// </summary>
    internal string? NameReference
    {
        get => _nameReference;
        set => Assign(ref _nameReference, value, XLChartSeriesFormat.NameReference);
    }

    public uint Index { get; internal set; }
    public uint Order { get; internal set; }

    public XLColor? FillColor
    {
        get => _fillColor;
        set => Assign(ref _fillColor, value, XLChartSeriesFormat.Fill);
    }

    public XLColor? LineColor
    {
        get => _lineColor;
        set => Assign(ref _lineColor, value, XLChartSeriesFormat.Line);
    }

    public double? LineWidthPt
    {
        get => _lineWidthPt;
        set
        {
            if (value is < 0 or > MaxLineWidthPt)
                throw new ArgumentOutOfRangeException(nameof(LineWidthPt), value,
                    $"Line width must be between 0 and {MaxLineWidthPt} points.");

            Assign(ref _lineWidthPt, value, XLChartSeriesFormat.LineWidth);
        }
    }

    public XLMarkerStyle MarkerStyle
    {
        get => _markerStyle;
        set => Assign(ref _markerStyle, value, XLChartSeriesFormat.Marker);
    }

    public double? MarkerSize
    {
        get => _markerSize;
        set
        {
            if (value is < MinMarkerSize or > MaxMarkerSize)
                throw new ArgumentOutOfRangeException(nameof(MarkerSize), value,
                    $"Marker size must be between {MinMarkerSize} and {MaxMarkerSize} points.");

            Assign(ref _markerSize, value, XLChartSeriesFormat.MarkerSize);
        }
    }

    public XLColor? MarkerFillColor
    {
        get => _markerFillColor;
        set => Assign(ref _markerFillColor, value, XLChartSeriesFormat.MarkerFill);
    }

    public bool Smooth
    {
        get => _smooth;
        set => Assign(ref _smooth, value, XLChartSeriesFormat.Smooth);
    }

    public bool UseSecondaryAxis
    {
        get => _useSecondaryAxis;
        set
        {
            if (value != _useSecondaryAxis && _chart.LoadedFromFile)
                throw new NotSupportedException(
                    "UseSecondaryAxis cannot be changed on a chart loaded from a file, because the " +
                    "series would have to be moved into a different plot group. Recreate the chart " +
                    "with IXLCharts.Add instead.");

            _useSecondaryAxis = value;
        }
    }

    public IXLDataLabels DataLabels => DataLabelsInternal;

    /// <summary>
    /// The data labels of this series, typed for internal use.
    /// </summary>
    internal XLChartDataLabels DataLabelsInternal { get; }

    /// <summary>
    /// The formatting properties that have been explicitly assigned through the public API.
    /// </summary>
    internal XLChartSeriesFormat AssignedFormat { get; private set; }

    /// <summary>
    /// Sets the series' references without marking them as assigned by the caller — the path a
    /// series takes when it is being created, from a file or from
    /// <see cref="IXLChartSeriesCollection.Add"/>, rather than edited afterwards.
    /// </summary>
    internal void SeedReferences(string valueReferences, string? categoryReferences)
    {
        _valueReferences = valueReferences ?? string.Empty;
        _categoryReferences = categoryReferences;
    }

    /// <summary>
    /// Sets the cell the series takes its name from, read from a loaded chart, without marking it as
    /// assigned, as <see cref="SeedReferences"/> does for the other references.
    /// </summary>
    internal void SeedNameReference(string? nameReference)
    {
        _nameReference = nameReference;
    }

    /// <summary>
    /// Rewrites the series' references the way <paramref name="rewrite"/> says, for a sheet rename or
    /// delete. A reference that changes is marked as rewritten rather than seeded: a loaded chart is
    /// patched in place (spec 10), and the patcher writes a reference only when it is marked. It is
    /// not marked as assigned by the caller, because the patcher drops the cached values of a
    /// reference the caller re-pointed, and keeps those of one a rename or delete rewrote. A reference
    /// the parser refuses keeps its text (ADR 0002).
    /// </summary>
    /// <param name="formulaSheetName">The sheet the chart is on.</param>
    /// <param name="rewrite">What the rename or the delete does to formula text.</param>
    internal void RewriteSheet(string formulaSheetName, SheetRewrite rewrite)
    {
        if (TryRewrite(_valueReferences, out var values))
        {
            _valueReferences = values;
            AssignedFormat |= XLChartSeriesFormat.ValueReferencesRewritten;
        }

        if (_categoryReferences is { } categories && TryRewrite(categories, out var rewrittenCategories))
        {
            _categoryReferences = rewrittenCategories;
            AssignedFormat |= XLChartSeriesFormat.CategoryReferencesRewritten;
        }

        if (_nameReference is { } name && TryRewrite(name, out var rewrittenName))
            NameReference = rewrittenName;

        // The rewrite does not move a reference, so any origin reads the reference the same way.
        bool TryRewrite(string reference, out string rewritten)
            => rewrite.TryRewrite(reference, formulaSheetName, new Point(1, 1), out rewritten)
               && rewritten != reference;
    }

    /// <summary>
    /// Seeds the formatting properties from the values read out of an existing chart part, without
    /// marking them as assigned by the caller. Values seeded this way are never written back, so a
    /// chart nobody edited keeps its original XML byte for byte.
    /// </summary>
#pragma warning disable S107 // One parameter per series format attribute read from the chart part
    internal void SeedLoadedFormat(
        XLColor? fillColor,
        XLColor? lineColor,
        double? lineWidthPt,
        XLMarkerStyle markerStyle,
        double? markerSize,
        XLColor? markerFillColor,
        bool smooth,
        bool useSecondaryAxis)
#pragma warning restore S107
    {
        _fillColor = fillColor;
        _lineColor = lineColor;
        _lineWidthPt = lineWidthPt;
        _markerStyle = markerStyle;
        _markerSize = markerSize;
        _markerFillColor = markerFillColor;
        _smooth = smooth;
        _useSecondaryAxis = useSecondaryAxis;
    }

    private void Assign<T>(ref T field, T value, XLChartSeriesFormat flag)
    {
        field = value;
        AssignedFormat |= flag;
    }
}
