using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.IO.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads chart definitions from an OpenXML worksheet part and populates the worksheet's chart collection.
/// Supports standard ChartPart charts and extended ExtendedChartPart charts (Office 2016+).
/// </summary>
internal static class ChartReader
{
    internal static void LoadCharts(WorksheetPart worksheetPart, XLWorksheet ws)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null)
            return;

        // Excel writes a chart under any of the three anchor kinds. A one-cell or absolute anchored
        // chart used to be skipped here, which left it out of ws.Charts entirely.
        foreach (var anchor in drawingsPart.WorksheetDrawing.ChildElements)
        {
            if (anchor is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
                continue;

            var composite = (OpenXmlCompositeElement)anchor;
            var xlChart = TryLoadChartFromAnchor(drawingsPart, composite, ws);
            if (xlChart == null)
                continue;

            ReadAnchor(composite, xlChart);
            ws.Charts.Add(xlChart);
        }
    }

    private static XLChart? TryLoadChartFromAnchor(
        DrawingsPart drawingsPart, OpenXmlCompositeElement anchor, XLWorksheet ws)
    {
        // GraphicFrame may be direct child or inside mc:AlternateContent > mc:Choice
        var graphicFrame = anchor.Elements<Xdr.GraphicFrame>().FirstOrDefault()
            ?? anchor.Descendants<Xdr.GraphicFrame>().FirstOrDefault();

        var graphicData = graphicFrame?.Graphic?.GraphicData;
        if (graphicData == null)
            return null;

        // Try standard chart reference
        var chartRef = graphicData.Elements<C.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value != null)
            return LoadStandardChart(drawingsPart, chartRef.Id.Value, ws);

        // Try extended chart reference (cx namespace)
        var cxRefId = ResolveExtendedChartRelId(graphicData);
        if (cxRefId != null)
            return LoadExtendedChart(drawingsPart, cxRefId, ws);

        return null;
    }

    private static string? ResolveExtendedChartRelId(A.GraphicData graphicData)
    {
        // GraphicData may deserialize cx:chart as OpenXmlUnknownElement, so also check by URI + r:id
        var cxRef = graphicData.Elements<Cx.RelId>().FirstOrDefault();
        var cxRefId = cxRef?.Id?.Value;

        if (cxRefId == null && graphicData.Uri == "http://schemas.microsoft.com/office/drawing/2014/chartex")
        {
            // Fallback: find the cx:chart element as an unknown element and extract r:id
            var unknownEl = graphicData.ChildElements.Count > 0 ? graphicData.ChildElements[0] : null;
            if (unknownEl != null)
            {
                var attr = unknownEl.GetAttribute("id", OpenXmlConst.RelationshipsNs);
                if (!string.IsNullOrWhiteSpace(attr.Value))
                    cxRefId = attr.Value;
            }
        }

        return cxRefId;
    }

    // ── Standard chart loading ──────────────────────────────────────────

    private static XLChart? LoadStandardChart(DrawingsPart drawingsPart, string relId, XLWorksheet ws)
    {
        var chartPart = (ChartPart)drawingsPart.GetPartById(relId);
        var chartSpace = chartPart.ChartSpace;
        if (chartSpace == null) return null;

        var chart = chartSpace.Elements<C.Chart>().FirstOrDefault();
        if (chart == null) return null;

        var xlChart = new XLChart(ws) { IsNew = false, RelId = relId };
        ReadTitle(chart, xlChart);
        ChartLegendXml.Read(chart, xlChart.LegendInternal);

        var plotArea = chart.PlotArea;
        if (plotArea != null)
            ReadPlotArea(plotArea, xlChart);

        // Set last: the series readers above assign properties that are refused on a loaded chart.
        xlChart.LoadedFromFile = true;
        return xlChart;
    }

    private static void ReadTitle(C.Chart chart, XLChart xlChart)
    {
        var title = chart.Title;
        if (title == null) return;

        var chartText = title.Elements<C.ChartText>().FirstOrDefault();
        var richText = chartText?.Elements<C.RichText>().FirstOrDefault();
        if (richText != null)
        {
            var text = string.Join("", richText.Descendants<A.Text>().Select(t => t.Text));
            if (!string.IsNullOrEmpty(text))
                xlChart.SeedLoadedTitle(text);
        }
    }

    private static void ReadPlotArea(C.PlotArea plotArea, XLChart xlChart)
    {
        var groups = ChartPlotAreaScanner.Scan(plotArea);
        var primaryKind = ChartPlotAreaScanner.ChoosePrimaryKind(groups);
        if (primaryKind == null)
            return;

        var primaryGroup = ChartPlotAreaScanner.PrimaryGroup(plotArea, groups, primaryKind.Value);
        var primaryValueAxisId = primaryGroup.ValueAxisId;
        ReadAxes(plotArea, groups, primaryGroup, primaryValueAxisId, xlChart);
        var primarySet = false;

        foreach (var group in groups)
        {
            var chartType = DetermineChartType(group);
            var isPrimary = group.Kind == primaryKind.Value;
            var useSecondaryAxis = group.ValueAxisId != null
                                   && primaryValueAxisId != null
                                   && group.ValueAxisId != primaryValueAxisId;

            if (isPrimary)
            {
                if (!primarySet)
                {
                    xlChart.ChartType = chartType;
                    primarySet = true;

                    // The chart-wide labels live on the primary chart group.
                    ChartDataLabelsXml.Read(
                        group.Element.Elements<C.DataLabels>().FirstOrDefault(), xlChart.DataLabelsInternal);
                }

                ReadGroupSeries(group, xlChart.SeriesInternal, useSecondaryAxis);
            }
            else
            {
                // The model carries two chart types. A plot area may hold more, in which case the
                // series of the third and later groups still land in the secondary collection but are
                // reported under the second group's type — see IXLChart.SecondaryChartType.
                xlChart.SecondaryChartType ??= chartType;
                ReadGroupSeries(group, xlChart.SecondarySeriesInternal, useSecondaryAxis);
            }
        }
    }

    /// <summary>
    /// Reads the horizontal axis, the value axis and — when a group hangs off a different value axis —
    /// the secondary value axis into the model.
    /// </summary>
    private static void ReadAxes(
        C.PlotArea plotArea, List<XLChartGroup> groups, XLChartGroup primaryGroup,
        uint? primaryValueAxisId, XLChart xlChart)
    {
        ChartAxisXml.Read(
            ChartPlotAreaScanner.FindAxis(plotArea, primaryGroup.CategoryAxisId),
            xlChart.CategoryAxisInternal);
        ChartAxisXml.Read(
            ChartPlotAreaScanner.FindAxis(plotArea, primaryValueAxisId),
            xlChart.ValueAxisInternal);

        var secondaryGroup = groups.Find(g =>
            g.ValueAxisId != null && primaryValueAxisId != null && g.ValueAxisId != primaryValueAxisId);
        if (secondaryGroup != null)
        {
            ChartAxisXml.Read(
                ChartPlotAreaScanner.FindAxis(plotArea, secondaryGroup.ValueAxisId),
                xlChart.SecondaryValueAxisInternal);
        }
    }

    private static XLChartType DetermineChartType(XLChartGroup group) => group.Kind switch
    {
        XLChartGroupKind.Bar => DetermineBarChartType((C.BarChart)group.Element),
        XLChartGroupKind.Bar3D => DetermineBar3DChartType((C.Bar3DChart)group.Element),
        XLChartGroupKind.Pie => XLChartType.Pie,
        XLChartGroupKind.Pie3D => XLChartType.Pie3D,
        XLChartGroupKind.OfPie => DetermineOfPieChartType((C.OfPieChart)group.Element),
        XLChartGroupKind.Doughnut => XLChartType.Doughnut,
        XLChartGroupKind.Area => DetermineAreaChartType((C.AreaChart)group.Element),
        XLChartGroupKind.Area3D => DetermineArea3DChartType((C.Area3DChart)group.Element),
        XLChartGroupKind.Line => DetermineLineChartType((C.LineChart)group.Element),
        XLChartGroupKind.Line3D => XLChartType.Line3D,
        XLChartGroupKind.Radar => DetermineRadarChartType((C.RadarChart)group.Element),
        XLChartGroupKind.Bubble => XLChartType.Bubble,
        XLChartGroupKind.Scatter => DetermineScatterChartType((C.ScatterChart)group.Element),
        XLChartGroupKind.Stock => XLChartType.StockHighLowClose,
        XLChartGroupKind.Surface => DetermineSurfaceChartType((C.SurfaceChart)group.Element),
        XLChartGroupKind.Surface3D => DetermineSurface3DChartType((C.Surface3DChart)group.Element),
        _ => XLChartType.ColumnClustered
    };

    /// <summary>
    /// Reads every series of one chart group — references, name and formatting — into the model.
    /// </summary>
    private static void ReadGroupSeries(
        XLChartGroup group, XLChartSeriesCollection target, bool useSecondaryAxis)
    {
        foreach (var seriesElement in group.SeriesElements)
        {
            var name = ExtractSeriesName(seriesElement.Elements<C.SeriesText>().FirstOrDefault());
            var (catRef, valRef) = group.IsXyBased
                ? ExtractXyReferences(seriesElement)
                : ExtractCategoryAndValueReferences(seriesElement);

            var series = (XLChartSeries)target.Add(name, valRef, catRef);
            ChartFormatting.ReadSeriesFormat(seriesElement, series, useSecondaryAxis);
            ChartDataLabelsXml.Read(
                seriesElement.Elements<C.DataLabels>().FirstOrDefault(), series.DataLabelsInternal);
        }
    }

    // ── Extended chart loading ──────────────────────────────────────────

    private static XLChart? LoadExtendedChart(DrawingsPart drawingsPart, string relId, XLWorksheet ws)
    {
        var extPart = (ExtendedChartPart)drawingsPart.GetPartById(relId);
        var chartSpace = extPart.ChartSpace;
        if (chartSpace == null) return null;

        var chartType = ReadExtendedChartType(chartSpace);
        if (chartType == null) return null;

        var xlChart = new XLChart(ws) { IsNew = false, RelId = relId, ChartType = chartType.Value };

        ReadExtendedTitle(chartSpace, xlChart);
        ReadExtendedSeries(chartSpace, xlChart);

        xlChart.LoadedFromFile = true;
        return xlChart;
    }

    private static void ReadExtendedTitle(Cx.ChartSpace chartSpace, XLChart xlChart)
    {
        var cxTitle = chartSpace.Descendants<Cx.ChartTitle>().FirstOrDefault();
        if (cxTitle == null) return;

        var titleText = string.Join("", cxTitle.Descendants<A.Text>().Select(t => t.Text));
        if (!string.IsNullOrEmpty(titleText))
            xlChart.SeedLoadedTitle(titleText);
    }

    private static XLChartType? ReadExtendedChartType(Cx.ChartSpace chartSpace)
    {
        var firstSeries = chartSpace.Descendants<Cx.Series>().FirstOrDefault();
        if (firstSeries == null) return null;

        var layoutId = firstSeries.GetAttribute("layoutId", string.Empty).Value ?? string.Empty;
        return layoutId switch
        {
            "sunburst" => XLChartType.Sunburst,
            "treemap" => XLChartType.Treemap,
            "waterfall" => XLChartType.Waterfall,
            "funnel" => XLChartType.Funnel,
            "boxWhisker" => XLChartType.BoxWhisker,
            _ => null
        };
    }

    private static void ReadExtendedSeries(Cx.ChartSpace chartSpace, XLChart xlChart)
    {
        var chartData = chartSpace.Descendants<Cx.ChartData>().FirstOrDefault();

        foreach (var cxSeries in chartSpace.Descendants<Cx.Series>())
        {
            var name = ReadExtendedSeriesName(cxSeries);
            var (catRef, valRef) = ReadExtendedSeriesRefs(cxSeries, chartData);
            xlChart.Series.Add(name, valRef, catRef);
        }
    }

    private static string ReadExtendedSeriesName(Cx.Series cxSeries)
    {
        var txData = cxSeries.Descendants<Cx.TextData>().FirstOrDefault();
        if (txData == null) return string.Empty;

        return txData.Descendants<Cx.VXsdstring>().FirstOrDefault()?.Text ?? string.Empty;
    }

    private static (string? catRef, string valRef) ReadExtendedSeriesRefs(
        Cx.Series cxSeries, Cx.ChartData? chartData)
    {
        var dataId = cxSeries.Descendants<Cx.DataId>().FirstOrDefault();
        if (dataId == null || chartData == null)
            return (null, string.Empty);

        var data = chartData.Elements<Cx.Data>()
            .FirstOrDefault(d => d.Id?.Value == dataId.Val?.Value);
        if (data == null)
            return (null, string.Empty);

        var catRef = data.Elements<Cx.StringDimension>().FirstOrDefault()
            ?.Elements<Cx.Formula>().FirstOrDefault()?.Text;

        var valRef = data.Elements<Cx.NumericDimension>().FirstOrDefault()
            ?.Elements<Cx.Formula>().FirstOrDefault()?.Text ?? string.Empty;

        return (catRef, valRef);
    }

    // ── Type determination helpers ──────────────────────────────────────

    private static XLChartType DetermineBarChartType(C.BarChart barChart)
    {
        var direction = barChart.BarDirection?.Val?.Value ?? C.BarDirectionValues.Column;
        var grouping = barChart.BarGrouping?.Val?.Value ?? C.BarGroupingValues.Clustered;

        if (direction == C.BarDirectionValues.Bar)
        {
            if (grouping == C.BarGroupingValues.Stacked) return XLChartType.BarStacked;
            if (grouping == C.BarGroupingValues.PercentStacked) return XLChartType.BarStacked100Percent;
            return XLChartType.BarClustered;
        }

        if (grouping == C.BarGroupingValues.Stacked) return XLChartType.ColumnStacked;
        if (grouping == C.BarGroupingValues.PercentStacked) return XLChartType.ColumnStacked100Percent;
        return XLChartType.ColumnClustered;
    }

    private static XLChartType DetermineLineChartType(C.LineChart lineChart)
    {
        var grouping = lineChart.Grouping?.Val?.Value;
        var hasMarkers = lineChart.Elements<C.LineChartSeries>().Any(HasVisibleMarker);

        if (grouping == C.GroupingValues.Stacked)
            return hasMarkers ? XLChartType.LineWithMarkersStacked : XLChartType.LineStacked;
        if (grouping == C.GroupingValues.PercentStacked)
            return hasMarkers ? XLChartType.LineWithMarkersStacked100Percent : XLChartType.LineStacked100Percent;

        return hasMarkers ? XLChartType.LineWithMarkers : XLChartType.Line;
    }

    /// <summary>
    /// Whether a series carries a marker that actually draws something. A <c>c:marker</c> holding
    /// <c>&lt;c:symbol val="none"/&gt;</c> switches markers off, so it must not turn a Line chart
    /// into a LineWithMarkers chart.
    /// </summary>
    private static bool HasVisibleMarker(C.LineChartSeries series)
    {
        var marker = series.Elements<C.Marker>().FirstOrDefault();
        if (marker == null)
            return false;

        var symbol = marker.Elements<C.Symbol>().FirstOrDefault()?.Val;
        return symbol == null || symbol.Value != C.MarkerStyleValues.None;
    }

    private static XLChartType DetermineRadarChartType(C.RadarChart radarChart) =>
        radarChart.RadarStyle?.Val?.Value == C.RadarStyleValues.Filled
            ? XLChartType.RadarFilled : XLChartType.Radar;

    private static XLChartType DetermineBar3DChartType(C.Bar3DChart bar3DChart)
    {
        var direction = bar3DChart.BarDirection?.Val?.Value ?? C.BarDirectionValues.Column;
        var grouping = bar3DChart.BarGrouping?.Val?.Value ?? C.BarGroupingValues.Clustered;
        var shape = bar3DChart.Elements<C.Shape>().FirstOrDefault()?.Val?.Value;
        var isHorizontal = direction == C.BarDirectionValues.Bar;

        if (shape == C.ShapeValues.Cone || shape == C.ShapeValues.ConeToMax)
            return ResolveBar3DGrouping(isHorizontal, grouping,
                horizontal: (XLChartType.ConeHorizontalClustered, XLChartType.ConeHorizontalStacked, XLChartType.ConeHorizontalStacked100Percent),
                vertical: (XLChartType.ConeClustered, XLChartType.ConeStacked, XLChartType.ConeStacked100Percent),
                verticalStandard: XLChartType.Cone);

        if (shape == C.ShapeValues.Cylinder)
            return ResolveBar3DGrouping(isHorizontal, grouping,
                horizontal: (XLChartType.CylinderHorizontalClustered, XLChartType.CylinderHorizontalStacked, XLChartType.CylinderHorizontalStacked100Percent),
                vertical: (XLChartType.CylinderClustered, XLChartType.CylinderStacked, XLChartType.CylinderStacked100Percent),
                verticalStandard: XLChartType.Cylinder);

        if (shape == C.ShapeValues.Pyramid || shape == C.ShapeValues.PyramidToMaximum)
            return ResolveBar3DGrouping(isHorizontal, grouping,
                horizontal: (XLChartType.PyramidHorizontalClustered, XLChartType.PyramidHorizontalStacked, XLChartType.PyramidHorizontalStacked100Percent),
                vertical: (XLChartType.PyramidClustered, XLChartType.PyramidStacked, XLChartType.PyramidStacked100Percent),
                verticalStandard: XLChartType.Pyramid);

        // Default: Box shape = standard 3D bar/column
        return ResolveBar3DBoxGrouping(isHorizontal, grouping);
    }

    private static XLChartType ResolveBar3DGrouping(
        bool isHorizontal, C.BarGroupingValues grouping,
        (XLChartType Clustered, XLChartType Stacked, XLChartType Stacked100) horizontal,
        (XLChartType Clustered, XLChartType Stacked, XLChartType Stacked100) vertical,
        XLChartType verticalStandard)
    {
        if (isHorizontal)
        {
            if (grouping == C.BarGroupingValues.Stacked) return horizontal.Stacked;
            if (grouping == C.BarGroupingValues.PercentStacked) return horizontal.Stacked100;
            return horizontal.Clustered;
        }

        if (grouping == C.BarGroupingValues.Stacked) return vertical.Stacked;
        if (grouping == C.BarGroupingValues.PercentStacked) return vertical.Stacked100;
        if (grouping == C.BarGroupingValues.Standard) return verticalStandard;
        return vertical.Clustered;
    }

    private static XLChartType ResolveBar3DBoxGrouping(bool isHorizontal, C.BarGroupingValues grouping)
    {
        if (isHorizontal)
        {
            if (grouping == C.BarGroupingValues.Stacked) return XLChartType.BarStacked3D;
            if (grouping == C.BarGroupingValues.PercentStacked) return XLChartType.BarStacked100Percent3D;
            return XLChartType.BarClustered3D;
        }

        if (grouping == C.BarGroupingValues.Stacked) return XLChartType.ColumnStacked3D;
        if (grouping == C.BarGroupingValues.PercentStacked) return XLChartType.ColumnStacked100Percent3D;
        if (grouping == C.BarGroupingValues.Standard) return XLChartType.Column3D;
        return XLChartType.ColumnClustered3D;
    }

    private static XLChartType DetermineAreaChartType(C.AreaChart areaChart)
    {
        var grouping = areaChart.Grouping?.Val?.Value;
        if (grouping == C.GroupingValues.Stacked) return XLChartType.AreaStacked;
        if (grouping == C.GroupingValues.PercentStacked) return XLChartType.AreaStacked100Percent;
        return XLChartType.Area;
    }

    /// <summary>
    /// A <c>c:surfaceChart</c> is strictly the flat contour variant, but XLibur's writer emits every
    /// surface type as one, so it is read back as the plain surface type its own writer meant.
    /// </summary>
    private static XLChartType DetermineSurfaceChartType(C.SurfaceChart surfaceChart)
    {
        var wireframe = surfaceChart.Elements<C.Wireframe>().FirstOrDefault()?.Val?.Value ?? false;
        return wireframe ? XLChartType.SurfaceWireframe : XLChartType.Surface;
    }

    private static XLChartType DetermineSurface3DChartType(C.Surface3DChart surfaceChart)
    {
        var wireframe = surfaceChart.Elements<C.Wireframe>().FirstOrDefault()?.Val?.Value ?? false;
        return wireframe ? XLChartType.SurfaceWireframe : XLChartType.Surface;
    }

    /// <summary>
    /// A <c>c:ofPieChart</c> is either pie-of-pie or bar-of-pie, told apart by <c>c:ofPieType</c>.
    /// </summary>
    private static XLChartType DetermineOfPieChartType(C.OfPieChart ofPieChart)
    {
        var type = ofPieChart.Elements<C.OfPieType>().FirstOrDefault()?.Val;
        return type != null && type.Value == C.OfPieValues.Bar
            ? XLChartType.PieToBar
            : XLChartType.PieToPie;
    }

    private static XLChartType DetermineArea3DChartType(C.Area3DChart areaChart)
    {
        var grouping = areaChart.Grouping?.Val?.Value;
        if (grouping == C.GroupingValues.Stacked) return XLChartType.AreaStacked3D;
        if (grouping == C.GroupingValues.PercentStacked) return XLChartType.AreaStacked100Percent3D;
        return XLChartType.Area3D;
    }

    private static XLChartType DetermineScatterChartType(C.ScatterChart scatterChart)
    {
        var style = scatterChart.ScatterStyle?.Val?.Value;
        if (style == C.ScatterStyleValues.SmoothMarker)
            return XLChartType.XYScatterSmoothLinesWithMarkers;
        return XLChartType.XYScatterMarkers;
    }

    // ── Shared extraction helpers ───────────────────────────────────────

    /// <summary>
    /// Extracts the category and value references of a series that uses <c>c:cat</c>/<c>c:val</c>.
    /// </summary>
    private static (string? catRef, string valRef) ExtractCategoryAndValueReferences(
        OpenXmlCompositeElement seriesElement)
    {
        var catData = seriesElement.Elements<C.CategoryAxisData>().FirstOrDefault();
        var valData = seriesElement.Elements<C.Values>().FirstOrDefault();

        string? catRef = null;
        if (catData != null)
        {
            catRef = catData.Elements<C.StringReference>().FirstOrDefault()?.Formula?.Text;
            catRef ??= catData.Elements<C.NumberReference>().FirstOrDefault()?.Formula?.Text;
        }

        var valRef = valData?.Elements<C.NumberReference>().FirstOrDefault()?.Formula?.Text ?? string.Empty;
        return (catRef, valRef);
    }

    /// <summary>
    /// Extracts the X and Y references of a scatter or bubble series, which uses
    /// <c>c:xVal</c>/<c>c:yVal</c> instead of categories and values.
    /// </summary>
    private static (string? catRef, string valRef) ExtractXyReferences(OpenXmlCompositeElement seriesElement)
    {
        string? xRef = null;
        var xValues = seriesElement.Elements<C.XValues>().FirstOrDefault();
        if (xValues != null)
        {
            xRef = xValues.Elements<C.NumberReference>().FirstOrDefault()?.Formula?.Text;
            xRef ??= xValues.Elements<C.StringReference>().FirstOrDefault()?.Formula?.Text;
        }

        var yValues = seriesElement.Elements<C.YValues>().FirstOrDefault();
        var yRef = yValues?.Elements<C.NumberReference>().FirstOrDefault()?.Formula?.Text ?? string.Empty;
        return (xRef, yRef);
    }

    /// <summary>
    /// Reads the series name from either form of <c>c:tx</c>: the literal <c>c:v</c> XLibur writes,
    /// or the <c>c:strRef</c> Excel writes when the name comes from a cell — in which case the
    /// cached value is the name the user sees.
    /// </summary>
    private static string ExtractSeriesName(C.SeriesText? seriesText)
    {
        if (seriesText == null) return string.Empty;

        var literal = seriesText.Elements<C.NumericValue>().FirstOrDefault()?.Text;
        if (literal != null) return literal;

        var strRef = seriesText.Elements<C.StringReference>().FirstOrDefault();
        var strCache = strRef?.Elements<C.StringCache>().FirstOrDefault();
        var pt = strCache?.Elements<C.StringPoint>().FirstOrDefault();
        return pt?.Elements<C.NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
    }

    // ── Position reading ────────────────────────────────────────────────

    /// <summary>EMU per pixel, the unit the drawing markers and extents are stored in.</summary>
    private const double EmuPerPixel = 9525;

    private static void ReadAnchor(OpenXmlCompositeElement anchor, XLChart xlChart)
    {
        switch (anchor)
        {
            case Xdr.TwoCellAnchor twoCell:
                xlChart.Anchor = XLDrawingAnchor.MoveAndSizeWithCells;
                ReadMarker(twoCell.FromMarker, xlChart.Position);
                ReadMarker(twoCell.ToMarker, xlChart.SecondPosition);
                break;

            case Xdr.OneCellAnchor oneCell:
                xlChart.Anchor = XLDrawingAnchor.MoveWithCells;
                ReadMarker(oneCell.FromMarker, xlChart.Position);
                ReadExtent(oneCell.Extent, xlChart);
                break;

            case Xdr.AbsoluteAnchor absolute:
                xlChart.Anchor = XLDrawingAnchor.Absolute;
                if (absolute.Position != null)
                {
                    xlChart.Left = ToPixels(absolute.Position.X?.Value);
                    xlChart.Top = ToPixels(absolute.Position.Y?.Value);
                }

                ReadExtent(absolute.Extent, xlChart);
                break;
        }
    }

    private static void ReadExtent(Xdr.Extent? extent, XLChart xlChart)
    {
        if (extent == null)
            return;

        xlChart.Width = ToPixels(extent.Cx?.Value);
        xlChart.Height = ToPixels(extent.Cy?.Value);
    }

    private static int ToPixels(long? emu) =>
        emu == null ? 0 : (int)System.Math.Round(emu.Value / EmuPerPixel);

    private static void ReadMarker(Xdr.MarkerType? marker, IXLDrawingPosition position)
    {
        if (marker == null)
            return;

        if (int.TryParse(marker.ColumnId?.Text, out var col)) position.Column = col;
        if (int.TryParse(marker.RowId?.Text, out var row)) position.Row = row;
        if (long.TryParse(marker.ColumnOffset?.Text, out var colOff)) position.ColumnOffset = colOff / 9525.0;
        if (long.TryParse(marker.RowOffset?.Text, out var rowOff)) position.RowOffset = rowOff / 9525.0;
    }
}
