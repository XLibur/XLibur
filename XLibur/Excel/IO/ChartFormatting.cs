using System;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.IO.Charts;
using XLibur.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO;

/// <summary>
/// Translates the series formatting exposed by <see cref="IXLChartSeries"/> to and from the
/// DrawingML elements Excel understands (<c>c:spPr</c>, <c>c:marker</c>, <c>c:smooth</c>).
/// </summary>
/// <remarks>
/// A <c>null</c> property always means "omit the element" so that Excel falls back to its own
/// automatic formatting. Nothing here ever writes an explicit black or white.
/// </remarks>
internal static class ChartFormatting
{
    /// <summary>EMU per point, the unit DrawingML uses for line widths.</summary>
    private const double EmuPerPoint = 12700;

    // ── Writing ─────────────────────────────────────────────────────────

    /// <summary>
    /// Builds the <c>c:spPr</c> element for a series, or returns <c>null</c> when the series has no
    /// explicit fill or outline.
    /// </summary>
    internal static C.ChartShapeProperties? BuildSeriesShapeProperties(XLChartSeries series)
    {
        var fill = BuildSolidFill(series.FillColor);
        var outline = BuildOutline(series.LineColor, series.LineWidthPt);
        if (fill == null && outline == null)
            return null;

        var shapeProperties = new C.ChartShapeProperties();
        if (fill != null)
            shapeProperties.Append(fill);
        if (outline != null)
            shapeProperties.Append(outline);
        return shapeProperties;
    }

    /// <summary>
    /// Builds the <c>c:marker</c> element for a series, or returns <c>null</c> when the series
    /// leaves every marker property automatic.
    /// </summary>
    /// <param name="series">The series whose marker is built.</param>
    /// <param name="autoSymbol">
    /// <c>true</c> for the chart types that draw markers by default (<c>LineWithMarkers*</c>), which
    /// need an explicit <c>&lt;c:symbol val="auto"/&gt;</c> even when the caller set nothing.
    /// </param>
    internal static C.Marker? BuildMarker(XLChartSeries series, bool autoSymbol)
    {
        var symbol = MapMarkerSymbol(series.MarkerStyle);
        var fill = BuildSolidFill(series.MarkerFillColor);

        if (symbol == null && series.MarkerSize == null && fill == null)
        {
            if (!autoSymbol)
                return null;

            return new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Auto });
        }

        var marker = new C.Marker();
        marker.Append(new C.Symbol { Val = symbol ?? C.MarkerStyleValues.Auto });
        if (series.MarkerSize != null)
            marker.Append(new C.Size { Val = (byte)Math.Round(series.MarkerSize.Value) });
        if (fill != null)
            marker.Append(new C.ChartShapeProperties(fill));
        return marker;
    }

    /// <summary>
    /// Builds the <c>c:smooth</c> element, or returns <c>null</c> to leave the chart type's own
    /// default in place.
    /// </summary>
    /// <param name="series">The series whose smoothing is built.</param>
    /// <param name="smoothByChartType">
    /// <c>true</c> for the chart types Excel smooths by default (the <c>XYScatterSmoothLines*</c>
    /// types), which are written as smoothed unless the caller asked otherwise.
    /// </param>
    internal static C.Smooth? BuildSmooth(XLChartSeries series, bool smoothByChartType)
    {
        if ((series.AssignedFormat & XLChartSeriesFormat.Smooth) != 0)
            return new C.Smooth { Val = series.Smooth };

        return smoothByChartType ? new C.Smooth { Val = true } : null;
    }

    private static A.SolidFill? BuildSolidFill(XLColor? color) =>
        color == null || !color.HasValue ? null : new A.SolidFill(BuildColor(color));

    private static A.Outline? BuildOutline(XLColor? color, double? widthPt)
    {
        var fill = BuildSolidFill(color);
        if (fill == null && widthPt == null)
            return null;

        var outline = new A.Outline();
        if (widthPt != null)
            outline.Width = (int)Math.Round(widthPt.Value * EmuPerPoint);
        if (fill != null)
            outline.Append(fill);
        return outline;
    }

    private static OpenXmlElement BuildColor(XLColor color)
    {
        if (color.ColorType == XLColorType.Theme)
            return new A.SchemeColor { Val = MapSchemeColor(color.ThemeColor) };

        return new A.RgbColorModelHex { Val = color.Color.ToHex().Substring(2) };
    }

    // ── Reading ─────────────────────────────────────────────────────────

    /// <summary>
    /// Reads the formatting of one <c>c:ser</c> element into the model.
    /// </summary>
    internal static void ReadSeriesFormat(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, bool useSecondaryAxis)
    {
        var shapeProperties = seriesElement.Elements<C.ChartShapeProperties>().FirstOrDefault();
        var outline = shapeProperties?.Elements<A.Outline>().FirstOrDefault();
        var marker = seriesElement.Elements<C.Marker>().FirstOrDefault();
        var markerFill = marker?.Elements<C.ChartShapeProperties>().FirstOrDefault();
        var smooth = seriesElement.Elements<C.Smooth>().FirstOrDefault();

        series.SeedLoadedFormat(
            fillColor: ReadSolidFillColor(shapeProperties),
            lineColor: ReadSolidFillColor(outline),
            lineWidthPt: outline?.Width?.Value is { } width ? width / EmuPerPoint : null,
            markerStyle: ReadMarkerStyle(marker),
            markerSize: marker?.Elements<C.Size>().FirstOrDefault()?.Val?.Value,
            markerFillColor: ReadSolidFillColor(markerFill),
            // c:smooth defaults to true when the element is present without a val attribute.
            smooth: smooth != null && (smooth.Val?.Value ?? true),
            useSecondaryAxis: useSecondaryAxis);
    }

    private static XLColor? ReadSolidFillColor(OpenXmlCompositeElement? parent)
    {
        var solidFill = parent?.Elements<A.SolidFill>().FirstOrDefault();
        if (solidFill == null)
            return null;

        var rgb = solidFill.Elements<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
        if (IsHexRgb(rgb))
            return XLColor.FromHexRgb(rgb!);

        var scheme = solidFill.Elements<A.SchemeColor>().FirstOrDefault()?.Val;
        if (scheme != null && TryMapThemeColor(scheme.Value, out var themeColor))
            return XLColor.FromTheme(themeColor);

        return null;
    }

    /// <summary>
    /// Whether the value is a six digit hexadecimal RGB colour. Malformed chart XML must be read
    /// as "no explicit colour" rather than throwing.
    /// </summary>
    private static bool IsHexRgb(string? value)
    {
        if (value is not { Length: 6 })
            return false;

        foreach (var c in value)
        {
            var isHexDigit = c is >= '0' and <= '9' or >= 'a' and <= 'f' or >= 'A' and <= 'F';
            if (!isHexDigit)
                return false;
        }

        return true;
    }

    private static XLMarkerStyle ReadMarkerStyle(C.Marker? marker)
    {
        var symbol = marker?.Elements<C.Symbol>().FirstOrDefault()?.Val;
        if (symbol == null)
            return XLMarkerStyle.Auto;

        var value = symbol.Value;
        if (value == C.MarkerStyleValues.None) return XLMarkerStyle.None;
        if (value == C.MarkerStyleValues.Circle) return XLMarkerStyle.Circle;
        if (value == C.MarkerStyleValues.Dash) return XLMarkerStyle.Dash;
        if (value == C.MarkerStyleValues.Diamond) return XLMarkerStyle.Diamond;
        if (value == C.MarkerStyleValues.Dot) return XLMarkerStyle.Dot;
        if (value == C.MarkerStyleValues.Plus) return XLMarkerStyle.Plus;
        if (value == C.MarkerStyleValues.Square) return XLMarkerStyle.Square;
        if (value == C.MarkerStyleValues.Star) return XLMarkerStyle.Star;
        if (value == C.MarkerStyleValues.Triangle) return XLMarkerStyle.Triangle;
        if (value == C.MarkerStyleValues.X) return XLMarkerStyle.X;
        return XLMarkerStyle.Auto;
    }

    // ── Patching an existing series element ─────────────────────────────

    /// <summary>
    /// Applies the formatting properties the caller assigned onto an existing <c>c:ser</c> element,
    /// leaving every other child — trendlines, error bars, data point overrides, gradients on
    /// properties XLibur does not model — exactly as it was.
    /// </summary>
    /// <param name="seriesElement">The <c>c:ser</c> element to update.</param>
    /// <param name="series">The model series holding the values and what was assigned.</param>
    /// <param name="kind">
    /// The kind of chart group the series belongs to. Only some series types accept a marker or a
    /// smoothing flag, so the kind decides which properties can be written at all.
    /// </param>
    /// <param name="chartType">The chart type of the group, which constrains the data label position.</param>
    internal static void PatchSeriesFormat(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartGroupKind kind,
        XLChartType chartType)
    {
        var assigned = series.AssignedFormat;

        if ((assigned & (XLChartSeriesFormat.Fill | XLChartSeriesFormat.Line | XLChartSeriesFormat.LineWidth)) != 0)
            PatchShapeProperties(seriesElement, series, assigned);

        var markerAssigned = assigned &
            (XLChartSeriesFormat.Marker | XLChartSeriesFormat.MarkerSize | XLChartSeriesFormat.MarkerFill);
        if (markerAssigned != 0 && SupportsMarker(kind))
            PatchMarker(seriesElement, series, assigned);

        if ((assigned & XLChartSeriesFormat.Smooth) != 0 && SupportsSmooth(kind))
            PatchSmooth(seriesElement, series);

        PatchSeriesReferences(seriesElement, series, kind);

        // Not gated on `assigned`: the labels track their own assignments.
        if (ChartDataLabelsXml.Supports(kind))
            ChartDataLabelsXml.Apply(seriesElement, series.DataLabelsInternal, chartType);
    }

    /// <summary>
    /// Points a loaded series at a different range, when the caller re-pointed it.
    /// </summary>
    /// <remarks>
    /// A scatter or bubble series holds its references in <c>c:xVal</c>/<c>c:yVal</c> rather than
    /// <c>c:cat</c>/<c>c:val</c>, so the elements to patch depend on the group kind — the same
    /// distinction the reader makes when it loads them.
    /// </remarks>
    private static void PatchSeriesReferences(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartGroupKind kind)
    {
        var assigned = series.AssignedFormat;
        var isXyBased = kind is XLChartGroupKind.Scatter or XLChartGroupKind.Bubble;

        if ((assigned & XLChartSeriesFormat.ValueReferences) != 0)
        {
            var values = isXyBased
                ? (OpenXmlCompositeElement?)seriesElement.Elements<C.YValues>().FirstOrDefault()
                : seriesElement.Elements<C.Values>().FirstOrDefault();

            PatchReferenceFormula(values, series.ValueReferences);
        }

        if ((assigned & XLChartSeriesFormat.CategoryReferences) != 0)
        {
            var categories = isXyBased
                ? (OpenXmlCompositeElement?)seriesElement.Elements<C.XValues>().FirstOrDefault()
                : seriesElement.Elements<C.CategoryAxisData>().FirstOrDefault();

            PatchReferenceFormula(categories, series.CategoryReferences);
        }
    }

    /// <summary>
    /// Rewrites the <c>c:f</c> of whichever reference element a series data holder contains, and
    /// drops the cached values that went with the old range.
    /// </summary>
    /// <remarks>
    /// The cache is what Excel draws before it recalculates. Left alone it would describe the range
    /// the series used to point at, so a chart re-pointed at more rows would open showing the old
    /// ones. Both caches are optional in the schema, so removing them is valid, and Excel rebuilds
    /// them from the formula on open.
    /// </remarks>
    private static void PatchReferenceFormula(OpenXmlCompositeElement? holder, string? reference)
    {
        if (holder == null || string.IsNullOrWhiteSpace(reference))
            return;

        var numberReference = holder.Elements<C.NumberReference>().FirstOrDefault();
        if (numberReference != null)
        {
            numberReference.Formula = new C.Formula(reference);
            numberReference.Elements<C.NumberingCache>().ToList().ForEach(cache => cache.Remove());
            return;
        }

        var stringReference = holder.Elements<C.StringReference>().FirstOrDefault();
        if (stringReference != null)
        {
            stringReference.Formula = new C.Formula(reference);
            stringReference.Elements<C.StringCache>().ToList().ForEach(cache => cache.Remove());
            return;
        }

        var multiLevelReference = holder.Elements<C.MultiLevelStringReference>().FirstOrDefault();
        if (multiLevelReference != null)
        {
            multiLevelReference.Formula = new C.Formula(reference);
            multiLevelReference.Elements<C.MultiLevelStringCache>().ToList().ForEach(cache => cache.Remove());
        }
    }

    /// <summary>Whether the series type of a chart group accepts a <c>c:marker</c> child.</summary>
    private static bool SupportsMarker(XLChartGroupKind kind) => kind is XLChartGroupKind.Line
        or XLChartGroupKind.Line3D or XLChartGroupKind.Scatter or XLChartGroupKind.Radar
        or XLChartGroupKind.Stock;

    /// <summary>Whether the series type of a chart group accepts a <c>c:smooth</c> child.</summary>
    private static bool SupportsSmooth(XLChartGroupKind kind) => kind is XLChartGroupKind.Line
        or XLChartGroupKind.Line3D or XLChartGroupKind.Scatter or XLChartGroupKind.Stock;

    /// <summary>
    /// The schema order of the children of <c>c:ser</c> that this class writes. Only the elements it
    /// touches need to be listed; anything else keeps its place.
    /// </summary>
    internal static readonly Type[] SeriesChildOrder =
    [
        typeof(C.Index), typeof(C.Order), typeof(C.SeriesText), typeof(C.ChartShapeProperties),
        typeof(C.Marker), typeof(C.DataLabels), typeof(C.CategoryAxisData), typeof(C.XValues),
        typeof(C.Values), typeof(C.YValues), typeof(C.Smooth), typeof(C.ExtensionList)
    ];

    /// <summary>
    /// The schema order of the children of <c>c:chart</c> that this class touches.
    /// </summary>
    /// <remarks>
    /// The 3D elements between <c>c:autoTitleDeleted</c> and <c>c:plotArea</c> are listed even though
    /// nothing here writes them: <see cref="InsertOrdered"/> steps over any child it cannot rank, so
    /// leaving them out put a newly inserted title after the <c>c:view3D</c> block of a 3D chart —
    /// which the schema rejects outright.
    /// </remarks>
    internal static readonly Type[] ChartChildOrder =
    [
        typeof(C.Title), typeof(C.AutoTitleDeleted), typeof(C.PivotFormats), typeof(C.View3D),
        typeof(C.Floor), typeof(C.SideWall), typeof(C.BackWall), typeof(C.PlotArea),
        typeof(C.Legend), typeof(C.PlotVisibleOnly), typeof(C.DisplayBlanksAs),
        typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:title</c>.</summary>
    internal static readonly Type[] TitleChildOrder =
    [
        typeof(C.ChartText), typeof(C.Layout), typeof(C.Overlay),
        typeof(C.ChartShapeProperties), typeof(C.TextProperties), typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:legend</c>.</summary>
    internal static readonly Type[] LegendChildOrder =
    [
        typeof(C.LegendPosition), typeof(C.LegendEntry), typeof(C.Layout), typeof(C.Overlay),
        typeof(C.ChartShapeProperties), typeof(C.TextProperties), typeof(C.ExtensionList)
    ];

    /// <summary>
    /// The schema order of the children of <c>c:catAx</c> and <c>c:valAx</c> that this class touches.
    /// The two types agree on the elements they share; the unit elements exist on <c>c:valAx</c> only.
    /// </summary>
    internal static readonly Type[] AxisChildOrder =
    [
        typeof(C.AxisId), typeof(C.Scaling), typeof(C.Delete), typeof(C.AxisPosition),
        typeof(C.MajorGridlines), typeof(C.MinorGridlines), typeof(C.Title),
        typeof(C.NumberingFormat), typeof(C.MajorTickMark), typeof(C.MinorTickMark),
        typeof(C.TickLabelPosition), typeof(C.ChartShapeProperties), typeof(C.TextProperties),
        typeof(C.CrossingAxis), typeof(C.Crosses), typeof(C.CrossesAt), typeof(C.CrossBetween),
        typeof(C.MajorUnit), typeof(C.MinorUnit), typeof(C.DisplayUnits), typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:scaling</c>.</summary>
    internal static readonly Type[] ScalingChildOrder =
    [
        typeof(C.LogBase), typeof(C.Orientation), typeof(C.MaxAxisValue), typeof(C.MinAxisValue),
        typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:dLbls</c>.</summary>
    internal static readonly Type[] DataLabelsChildOrder =
    [
        typeof(C.DataLabel), typeof(C.Delete), typeof(C.NumberingFormat),
        typeof(C.ChartShapeProperties), typeof(C.TextProperties), typeof(C.DataLabelPosition),
        typeof(C.ShowLegendKey), typeof(C.ShowValue), typeof(C.ShowCategoryName),
        typeof(C.ShowSeriesName), typeof(C.ShowPercent), typeof(C.ShowBubbleSize),
        typeof(C.Separator), typeof(C.ExtensionList)
    ];

    /// <summary>
    /// Inserts <paramref name="element"/> at the position the schema order gives it: before the first
    /// child that must come after it, or at the end when there is none.
    /// </summary>
    internal static void InsertOrdered(OpenXmlCompositeElement parent, OpenXmlElement element, Type[] order)
    {
        var rank = Array.IndexOf(order, element.GetType());
        if (rank < 0)
        {
            parent.Append(element);
            return;
        }

        foreach (var child in parent.ChildElements)
        {
            var childRank = Array.IndexOf(order, child.GetType());
            if (childRank > rank)
            {
                parent.InsertBefore(element, child);
                return;
            }
        }

        parent.Append(element);
    }

    private static void PatchShapeProperties(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartSeriesFormat assigned)
    {
        var shapeProperties = EnsureShapeProperties(seriesElement);

        if ((assigned & XLChartSeriesFormat.Fill) != 0)
            SetFill(shapeProperties, series.FillColor);

        if ((assigned & (XLChartSeriesFormat.Line | XLChartSeriesFormat.LineWidth)) != 0)
            SetOutline(shapeProperties, series, assigned);
    }

    private static void PatchMarker(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartSeriesFormat assigned)
    {
        var marker = seriesElement.Elements<C.Marker>().FirstOrDefault();
        var markerIsNew = marker == null;
        if (marker == null)
        {
            marker = new C.Marker();
            InsertAfterLastOf(seriesElement, marker,
                typeof(C.ChartShapeProperties), typeof(C.SeriesText), typeof(C.Order), typeof(C.Index));
        }

        if ((assigned & XLChartSeriesFormat.Marker) != 0)
        {
            var symbol = MapMarkerSymbol(series.MarkerStyle);
            marker.Elements<C.Symbol>().ToList().ForEach(e => e.Remove());
            if (symbol != null)
                marker.InsertAt(new C.Symbol { Val = symbol }, 0);
        }

        if ((assigned & XLChartSeriesFormat.MarkerSize) != 0)
        {
            marker.Elements<C.Size>().ToList().ForEach(e => e.Remove());
            if (series.MarkerSize != null)
            {
                var size = new C.Size { Val = (byte)Math.Round(series.MarkerSize.Value) };
                InsertAfterLastOf(marker, size, typeof(C.Symbol));
            }
        }

        if ((assigned & XLChartSeriesFormat.MarkerFill) != 0)
            PatchMarkerFill(marker, series);

        // A marker element that ended up empty would read back as "this series has markers", so an
        // element created here for nothing is taken away again.
        if (markerIsNew && !marker.HasChildren)
            marker.Remove();
    }

    private static void PatchMarkerFill(C.Marker marker, XLChartSeries series)
    {
        var markerShapeProperties = marker.Elements<C.ChartShapeProperties>().FirstOrDefault();
        if (markerShapeProperties == null)
        {
            if (series.MarkerFillColor == null)
                return;

            markerShapeProperties = new C.ChartShapeProperties();
            InsertAfterLastOf(marker, markerShapeProperties, typeof(C.Size), typeof(C.Symbol));
        }

        SetFill(markerShapeProperties, series.MarkerFillColor);
    }

    private static void PatchSmooth(OpenXmlCompositeElement seriesElement, XLChartSeries series)
    {
        var existing = seriesElement.Elements<C.Smooth>().ToList();

        if (!series.Smooth)
        {
            // An explicit false is written, rather than dropped, so that a chart whose type smooths
            // by default (the XYScatterSmoothLines* types) can be turned back into straight lines.
            if (existing.Count == 0)
                AppendBeforeExtensionList(seriesElement, new C.Smooth { Val = false });
            else
                existing[0].Val = false;
        }
        else if (existing.Count == 0)
        {
            AppendBeforeExtensionList(seriesElement, new C.Smooth { Val = true });
        }
        else
        {
            existing[0].Val = true;
        }

        for (var i = 1; i < existing.Count; i++)
            existing[i].Remove();
    }

    private static C.ChartShapeProperties EnsureShapeProperties(OpenXmlCompositeElement seriesElement)
    {
        var shapeProperties = seriesElement.Elements<C.ChartShapeProperties>().FirstOrDefault();
        if (shapeProperties != null)
            return shapeProperties;

        shapeProperties = new C.ChartShapeProperties();
        InsertAfterLastOf(seriesElement, shapeProperties,
            typeof(C.SeriesText), typeof(C.Order), typeof(C.Index));
        return shapeProperties;
    }

    private static void SetFill(C.ChartShapeProperties shapeProperties, XLColor? color)
    {
        foreach (var existing in shapeProperties.ChildElements
                     .Where(IsFillElement).ToList())
            existing.Remove();

        if (color == null || !color.HasValue)
            return;

        // The fill precedes the outline and every effect in CT_ShapeProperties.
        var anchor = shapeProperties.ChildElements.FirstOrDefault(IsAfterFill);
        var fill = new A.SolidFill(BuildColor(color));
        if (anchor != null)
            shapeProperties.InsertBefore(fill, anchor);
        else
            shapeProperties.Append(fill);
    }

#pragma warning disable S3776 // Create-or-update outline, then one flat block per assigned property
    private static void SetOutline(
        C.ChartShapeProperties shapeProperties, XLChartSeries series, XLChartSeriesFormat assigned)
    {
        var outline = shapeProperties.Elements<A.Outline>().FirstOrDefault();
        if (outline == null)
        {
            // Nothing to clear and nothing to set: leave the element out entirely.
            if (series.LineColor == null && series.LineWidthPt == null)
                return;

            outline = new A.Outline();
            var anchor = shapeProperties.ChildElements.FirstOrDefault(IsAfterOutline);
            if (anchor != null)
                shapeProperties.InsertBefore(outline, anchor);
            else
                shapeProperties.Append(outline);
        }

        if ((assigned & XLChartSeriesFormat.LineWidth) != 0)
        {
            if (series.LineWidthPt == null)
                outline.Width = null;
            else
                outline.Width = (int)Math.Round(series.LineWidthPt.Value * EmuPerPoint);
        }

        if ((assigned & XLChartSeriesFormat.Line) != 0)
        {
            foreach (var existing in outline.ChildElements.Where(IsFillElement).ToList())
                existing.Remove();

            if (series.LineColor is { HasValue: true })
                outline.InsertAt(new A.SolidFill(BuildColor(series.LineColor)), 0);
        }
    }
#pragma warning restore S3776

    private static bool IsFillElement(OpenXmlElement element) =>
        element is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    private static bool IsAfterFill(OpenXmlElement element) =>
        element is A.Outline || IsEffectOrLater(element);

    private static bool IsAfterOutline(OpenXmlElement element) => IsEffectOrLater(element);

    private static bool IsEffectOrLater(OpenXmlElement element) =>
        element is A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList;

    /// <summary>
    /// Inserts <paramref name="element"/> directly after the last present element of the given types,
    /// which must be listed most- to least-preferred. Falls back to inserting first.
    /// </summary>
    private static void InsertAfterLastOf(
        OpenXmlCompositeElement parent, OpenXmlElement element, params Type[] precedingTypes)
    {
        foreach (var type in precedingTypes)
        {
            var anchor = parent.ChildElements.LastOrDefault(e => e.GetType() == type);
            if (anchor != null)
            {
                parent.InsertAfter(element, anchor);
                return;
            }
        }

        parent.InsertAt(element, 0);
    }

    private static void AppendBeforeExtensionList(OpenXmlCompositeElement parent, OpenXmlElement element)
    {
        var extensionList = parent.Elements<C.ExtensionList>().FirstOrDefault();
        if (extensionList != null)
            parent.InsertBefore(element, extensionList);
        else
            parent.Append(element);
    }

    // ── Enum mapping ────────────────────────────────────────────────────

    private static C.MarkerStyleValues? MapMarkerSymbol(XLMarkerStyle style) => style switch
    {
        XLMarkerStyle.Auto => null,
        XLMarkerStyle.None => C.MarkerStyleValues.None,
        XLMarkerStyle.Circle => C.MarkerStyleValues.Circle,
        XLMarkerStyle.Dash => C.MarkerStyleValues.Dash,
        XLMarkerStyle.Diamond => C.MarkerStyleValues.Diamond,
        XLMarkerStyle.Dot => C.MarkerStyleValues.Dot,
        XLMarkerStyle.Plus => C.MarkerStyleValues.Plus,
        XLMarkerStyle.Square => C.MarkerStyleValues.Square,
        XLMarkerStyle.Star => C.MarkerStyleValues.Star,
        XLMarkerStyle.Triangle => C.MarkerStyleValues.Triangle,
        XLMarkerStyle.X => C.MarkerStyleValues.X,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "Unknown marker style.")
    };

    private static A.SchemeColorValues MapSchemeColor(XLThemeColor themeColor) => themeColor switch
    {
        XLThemeColor.Background1 => A.SchemeColorValues.Background1,
        XLThemeColor.Text1 => A.SchemeColorValues.Text1,
        XLThemeColor.Background2 => A.SchemeColorValues.Background2,
        XLThemeColor.Text2 => A.SchemeColorValues.Text2,
        XLThemeColor.Accent1 => A.SchemeColorValues.Accent1,
        XLThemeColor.Accent2 => A.SchemeColorValues.Accent2,
        XLThemeColor.Accent3 => A.SchemeColorValues.Accent3,
        XLThemeColor.Accent4 => A.SchemeColorValues.Accent4,
        XLThemeColor.Accent5 => A.SchemeColorValues.Accent5,
        XLThemeColor.Accent6 => A.SchemeColorValues.Accent6,
        XLThemeColor.Hyperlink => A.SchemeColorValues.Hyperlink,
        XLThemeColor.FollowedHyperlink => A.SchemeColorValues.FollowedHyperlink,
        _ => throw new ArgumentOutOfRangeException(nameof(themeColor), themeColor, "Unknown theme colour.")
    };

    /// <summary>
    /// The scheme colours XLibur models, paired with the theme slot each maps onto. The first four
    /// slots answer to two names apiece: <c>bg1</c>/<c>lt1</c> and <c>tx1</c>/<c>dk1</c> are the same
    /// slot addressed from a colour map or from the theme itself, and a producer may write either.
    /// </summary>
    /// <remarks>
    /// A plain array rather than a dictionary: <c>SchemeColorValues</c> is an Open XML SDK 3 enum
    /// value type whose equality is defined by its <c>==</c> operator, not by GetHashCode, so a scan
    /// over sixteen entries is both correct and cheaper than hashing.
    /// </remarks>
    private static readonly (A.SchemeColorValues Scheme, XLThemeColor Theme)[] ThemeColorMap =
    [
        (A.SchemeColorValues.Background1, XLThemeColor.Background1),
        (A.SchemeColorValues.Light1, XLThemeColor.Background1),
        (A.SchemeColorValues.Text1, XLThemeColor.Text1),
        (A.SchemeColorValues.Dark1, XLThemeColor.Text1),
        (A.SchemeColorValues.Background2, XLThemeColor.Background2),
        (A.SchemeColorValues.Light2, XLThemeColor.Background2),
        (A.SchemeColorValues.Text2, XLThemeColor.Text2),
        (A.SchemeColorValues.Dark2, XLThemeColor.Text2),
        (A.SchemeColorValues.Accent1, XLThemeColor.Accent1),
        (A.SchemeColorValues.Accent2, XLThemeColor.Accent2),
        (A.SchemeColorValues.Accent3, XLThemeColor.Accent3),
        (A.SchemeColorValues.Accent4, XLThemeColor.Accent4),
        (A.SchemeColorValues.Accent5, XLThemeColor.Accent5),
        (A.SchemeColorValues.Accent6, XLThemeColor.Accent6),
        (A.SchemeColorValues.Hyperlink, XLThemeColor.Hyperlink),
        (A.SchemeColorValues.FollowedHyperlink, XLThemeColor.FollowedHyperlink),
    ];

    private static bool TryMapThemeColor(A.SchemeColorValues value, out XLThemeColor themeColor)
    {
        foreach (var (scheme, theme) in ThemeColorMap)
        {
            if (value == scheme)
            {
                themeColor = theme;
                return true;
            }
        }

        themeColor = default;
        return false;
    }
}
