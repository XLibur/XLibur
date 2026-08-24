using System;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.IO.DrawingML;
using XLibur.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// The formatting of one <c>c:ser</c>: its shape properties, its marker, its smoothing and the
/// references its data comes from.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Apply"/> covers both a chart being created and a chart loaded from a file. Creating a
/// child is the branch it takes when there is none, so the writer and the patcher stop being two
/// implementations of the same mapping.
/// </para>
/// <para>
/// A <c>null</c> property always means "omit the element" so that Excel falls back to its own
/// automatic formatting. Nothing here ever writes an explicit black or white. Everything else a
/// series carries — trendlines, error bars, data point overrides, gradients on properties XLibur
/// does not model — is left exactly as it was.
/// </para>
/// </remarks>
internal static class ChartSeriesFormatXml
{
    /// <summary>
    /// Seeds the model from one <c>c:ser</c> element.
    /// </summary>
    internal static void Read(
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
            lineWidthPt: outline?.Width?.Value is { } width ? width / DrawingUnits.EmuPerPoint : null,
            markerStyle: ReadMarkerStyle(marker),
            markerSize: marker?.Elements<C.Size>().FirstOrDefault()?.Val?.Value,
            markerFillColor: ReadSolidFillColor(markerFill),
            // c:smooth defaults to true when the element is present without a val attribute.
            smooth: smooth != null && (smooth.Val?.Value ?? true),
            useSecondaryAxis: useSecondaryAxis);
    }

    /// <summary>
    /// Writes the children a series of <paramref name="chartType"/> carries whether or not anybody
    /// asked for them, so that <see cref="Apply"/> is left with only what the caller assigned.
    /// </summary>
    /// <remarks>
    /// These are properties of the chart type rather than of the model, which is why they are not a
    /// branch of <see cref="Apply"/>: a series element that already exists in a file has already had
    /// its type's defaults applied by whoever wrote it, and conjuring them again would edit a chart
    /// nobody asked to change.
    /// </remarks>
    internal static void ApplyChartTypeDefaults(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartType chartType)
    {
        // A chart type that draws markers needs an explicit <c:symbol val="auto"/>; without one the
        // shape comes from the chart style.
        if (DrawsMarkers(chartType) && !seriesElement.Elements<C.Marker>().Any())
        {
            InsertAfterLastOf(seriesElement, new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Auto }),
                typeof(C.ChartShapeProperties), typeof(C.SeriesText), typeof(C.Order), typeof(C.Index));
        }

        if (SmoothsLines(chartType)
            && (series.AssignedFormat & XLChartSeriesFormat.Smooth) == 0
            && !seriesElement.Elements<C.Smooth>().Any())
        {
            AppendBeforeExtensionList(seriesElement, new C.Smooth { Val = true });
        }
    }

    /// <summary>
    /// Writes the formatting properties the caller assigned into <paramref name="seriesElement"/>.
    /// </summary>
    /// <param name="seriesElement">The <c>c:ser</c> element to write into.</param>
    /// <param name="series">The model series holding the values and what was assigned.</param>
    /// <param name="kind">
    /// The kind of chart group the series belongs to. Only some series types accept a marker or a
    /// smoothing flag, so the kind decides which properties can be written at all.
    /// </param>
    /// <param name="chartType">The chart type of the group, which constrains the data label position.</param>
    internal static void Apply(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartGroupKind kind,
        XLChartType chartType)
    {
        var assigned = series.AssignedFormat;

        if ((assigned & (XLChartSeriesFormat.Fill | XLChartSeriesFormat.Line | XLChartSeriesFormat.LineWidth)) != 0)
            ApplyShapeProperties(seriesElement, series, assigned);

        var markerAssigned = assigned &
            (XLChartSeriesFormat.Marker | XLChartSeriesFormat.MarkerSize | XLChartSeriesFormat.MarkerFill);
        if (markerAssigned != 0 && SupportsMarker(kind))
            ApplyMarker(seriesElement, series, assigned);

        if ((assigned & XLChartSeriesFormat.Smooth) != 0 && SupportsSmooth(kind))
            ApplySmooth(seriesElement, series);

        ApplyReferences(seriesElement, series, kind);

        // Not gated on `assigned`: the labels track their own assignments.
        if (ChartDataLabelsXml.Supports(kind))
            ChartDataLabelsXml.Apply(seriesElement, series.DataLabelsInternal, chartType);
    }

    /// <summary>
    /// Points a series at a different range, when the caller re-pointed it.
    /// </summary>
    /// <remarks>
    /// A scatter or bubble series holds its references in <c>c:xVal</c>/<c>c:yVal</c> rather than
    /// <c>c:cat</c>/<c>c:val</c>, so the elements to write depend on the group kind — the same
    /// distinction the reader makes when it loads them.
    /// </remarks>
    private static void ApplyReferences(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartGroupKind kind)
    {
        var assigned = series.AssignedFormat;
        var isXyBased = kind is XLChartGroupKind.Scatter or XLChartGroupKind.Bubble;

        if ((assigned & XLChartSeriesFormat.ValueReferences) != 0)
        {
            var values = isXyBased
                ? (OpenXmlCompositeElement?)seriesElement.Elements<C.YValues>().FirstOrDefault()
                : seriesElement.Elements<C.Values>().FirstOrDefault();

            SetReferenceFormula(values, series.ValueReferences);
        }

        if ((assigned & XLChartSeriesFormat.CategoryReferences) != 0)
        {
            var categories = isXyBased
                ? (OpenXmlCompositeElement?)seriesElement.Elements<C.XValues>().FirstOrDefault()
                : seriesElement.Elements<C.CategoryAxisData>().FirstOrDefault();

            SetReferenceFormula(categories, series.CategoryReferences);
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
    private static void SetReferenceFormula(OpenXmlCompositeElement? holder, string? reference)
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

    /// <summary>The chart types Excel draws markers for whether or not one was asked for.</summary>
    private static bool DrawsMarkers(XLChartType chartType) => chartType
        is XLChartType.LineWithMarkers or XLChartType.LineWithMarkersStacked
        or XLChartType.LineWithMarkersStacked100Percent;

    /// <summary>The chart types Excel smooths whether or not smoothing was asked for.</summary>
    private static bool SmoothsLines(XLChartType chartType) => chartType
        is XLChartType.XYScatterSmoothLinesNoMarkers or XLChartType.XYScatterSmoothLinesWithMarkers;

    private static void ApplyShapeProperties(
        OpenXmlCompositeElement seriesElement, XLChartSeries series, XLChartSeriesFormat assigned)
    {
        var existing = seriesElement.Elements<C.ChartShapeProperties>().FirstOrDefault();
        var shapeProperties = existing ?? NewShapeProperties(seriesElement);

        if ((assigned & XLChartSeriesFormat.Fill) != 0)
            ShapePropertiesWriter.SetFill(shapeProperties, series.FillColor);

        if ((assigned & (XLChartSeriesFormat.Line | XLChartSeriesFormat.LineWidth)) != 0)
            SetOutline(shapeProperties, series, assigned);

        // Shape properties created here for a fill and an outline that both turned out to be absent
        // would read back as "this series is explicitly formatted", so they are taken away again.
        if (existing == null && !shapeProperties.HasChildren)
            shapeProperties.Remove();
    }

    private static C.ChartShapeProperties NewShapeProperties(OpenXmlCompositeElement seriesElement)
    {
        var shapeProperties = new C.ChartShapeProperties();
        InsertAfterLastOf(seriesElement, shapeProperties,
            typeof(C.SeriesText), typeof(C.Order), typeof(C.Index));
        return shapeProperties;
    }

    private static void ApplyMarker(
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
            SetMarkerFill(marker, series);

        // A marker element built here says the series has markers, so it says which: an absent
        // c:symbol would leave the shape to the chart style rather than to the caller.
        if (markerIsNew && marker.HasChildren && !marker.Elements<C.Symbol>().Any())
            marker.InsertAt(new C.Symbol { Val = C.MarkerStyleValues.Auto }, 0);

        // A marker element that ended up empty would read back as "this series has markers", so an
        // element created here for nothing is taken away again.
        if (markerIsNew && !marker.HasChildren)
            marker.Remove();
    }

    private static void SetMarkerFill(C.Marker marker, XLChartSeries series)
    {
        var markerShapeProperties = marker.Elements<C.ChartShapeProperties>().FirstOrDefault();
        if (markerShapeProperties == null)
        {
            if (series.MarkerFillColor == null)
                return;

            markerShapeProperties = new C.ChartShapeProperties();
            InsertAfterLastOf(marker, markerShapeProperties, typeof(C.Size), typeof(C.Symbol));
        }

        ShapePropertiesWriter.SetFill(markerShapeProperties, series.MarkerFillColor);
    }

    private static void ApplySmooth(OpenXmlCompositeElement seriesElement, XLChartSeries series)
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

    /// <summary>
    /// Translates what the caller assigned into the outline operations that says.
    /// </summary>
    /// <remarks>
    /// The mask stops here. <see cref="ShapePropertiesWriter"/> is told to set a width or set a
    /// colour and works out what the schema requires; which of those to ask for is a question about
    /// <see cref="XLChartSeries.AssignedFormat"/>, which is chart knowledge and belongs on this side.
    /// </remarks>
    private static void SetOutline(
        C.ChartShapeProperties shapeProperties, XLChartSeries series, XLChartSeriesFormat assigned)
    {
        // Nothing to clear and nothing to set: leave the element out entirely.
        if (!shapeProperties.Elements<A.Outline>().Any()
            && series.LineColor == null && series.LineWidthPt == null)
            return;

        var outline = ShapePropertiesWriter.EnsureOutline(shapeProperties);

        if ((assigned & XLChartSeriesFormat.LineWidth) != 0)
            ShapePropertiesWriter.SetOutlineWidth(outline, series.LineWidthPt);

        if ((assigned & XLChartSeriesFormat.Line) != 0)
            ShapePropertiesWriter.SetOutlineColor(outline, series.LineColor);
    }

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
