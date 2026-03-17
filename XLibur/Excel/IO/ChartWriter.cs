using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.XLWorkbook;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes newly created charts to OpenXML. For each chart with <see cref="XLChart.IsNew"/> == <c>true</c>,
/// creates a ChartPart containing the ChartSpace DOM and a TwoCellAnchor with a GraphicFrame
/// referencing the chart. Ensures the worksheet's <c>&lt;drawing&gt;</c> element exists.
/// Charts loaded from an existing file (<see cref="XLChart.IsNew"/> == <c>false</c>) are skipped
/// and preserved untouched by OpenXML.
/// </summary>
internal static class ChartWriter
{
    /// <summary>
    /// Writes all new charts from the worksheet to the OpenXML worksheet part.
    /// Skips charts that were loaded from an existing file.
    /// </summary>
    internal static void WriteCharts(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        foreach (var chart in xlWorksheet.Charts)
        {
            var xlChart = (XLChart)chart;
            if (!xlChart.IsNew)
                continue;

            WriteChart(worksheet, cm, worksheetPart, xlChart, context);
        }
    }

    /// <summary>
    /// Writes a single chart: creates the ChartPart, builds the ChartSpace DOM,
    /// appends a TwoCellAnchor/GraphicFrame to the DrawingsPart, and ensures
    /// the worksheet XML contains a <c>&lt;drawing&gt;</c> reference.
    /// </summary>
    private static void WriteChart(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLChart xlChart,
        SaveContext context)
    {
        var drawingsPart = worksheetPart.DrawingsPart ??
                           worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        EnsureNamespaces(worksheetDrawing);

        // Create chart part
        var chartRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var chartPart = drawingsPart.AddNewPart<ChartPart>(chartRelId);

        // Build chart DOM
        chartPart.ChartSpace = BuildChartSpace(xlChart);

        // Create TwoCellAnchor with GraphicFrame
        var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
        var nvpId = nvps.Any()
            ? (UInt32Value)nvps.Max(p => p.Id!.Value) + 1
            : 1U;

        var fromPos = xlChart.Position;
        var toPos = xlChart.SecondPosition;

        var anchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker
            {
                ColumnId = new Xdr.ColumnId(fromPos.Column.ToString()),
                RowId = new Xdr.RowId(fromPos.Row.ToString()),
                ColumnOffset = new Xdr.ColumnOffset(((long)(fromPos.ColumnOffset * 9525)).ToString()),
                RowOffset = new Xdr.RowOffset(((long)(fromPos.RowOffset * 9525)).ToString())
            },
            new Xdr.ToMarker
            {
                ColumnId = new Xdr.ColumnId(toPos.Column.ToString()),
                RowId = new Xdr.RowId(toPos.Row.ToString()),
                ColumnOffset = new Xdr.ColumnOffset(((long)(toPos.ColumnOffset * 9525)).ToString()),
                RowOffset = new Xdr.RowOffset(((long)(toPos.RowOffset * 9525)).ToString())
            },
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = xlChart.Name },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()
                ),
                new Xdr.Transform(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new C.ChartReference { Id = chartRelId }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                )
            ),
            new Xdr.ClientData()
        );

        worksheetDrawing.Append(anchor);

        // Ensure <drawing> element in worksheet XML
        if (!worksheet.OfType<Drawing>().Any())
        {
            var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            var worksheetDrawingRef = new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) };
            worksheetDrawingRef.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            if (tableParts != null)
                worksheet.InsertBefore(worksheetDrawingRef, tableParts);
            else
                worksheet.AppendChild(worksheetDrawingRef);

            cm.SetElement(XLWorksheetContents.Drawing, worksheet.Elements<Drawing>().First());
        }
    }

    /// <summary>
    /// Builds the complete OpenXML ChartSpace DOM for a chart, including the BarChart element
    /// with series, category/value axes, and an optional title.
    /// </summary>
    private static ChartSpace BuildChartSpace(XLChart xlChart)
    {
        var barChart = new BarChart
        {
            BarDirection = new BarDirection { Val = GetBarDirection(xlChart) },
            BarGrouping = new BarGrouping { Val = GetGrouping(xlChart) }
        };

        foreach (var s in xlChart.Series)
        {
            var series = new BarChartSeries
            {
                Index = new Index { Val = s.Index },
                Order = new Order { Val = s.Order },
                SeriesText = new SeriesText(
                    new StringReference(
                        new StringCache(
                            new PointCount { Val = 1 },
                            new StringPoint(new NumericValue(s.Name)) { Index = 0 }
                        )
                    )
                )
            };

            if (s.CategoryReferences != null)
            {
                series.Append(new CategoryAxisData(
                    new StringReference
                    {
                        Formula = new C.Formula(s.CategoryReferences)
                    }
                ));
            }

            series.Append(new C.Values(
                new NumberReference
                {
                    Formula = new C.Formula(s.ValueReferences)
                }
            ));

            barChart.Append(series);
        }

        const uint catAxisId = 1u;
        const uint valAxisId = 2u;

        barChart.Append(new AxisId { Val = catAxisId });
        barChart.Append(new AxisId { Val = valAxisId });

        var plotArea = new PlotArea(
            new Layout(),
            barChart,
            new CategoryAxis(
                new AxisId { Val = catAxisId },
                new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Bottom },
                new CrossingAxis { Val = valAxisId }
            ),
            new ValueAxis(
                new AxisId { Val = valAxisId },
                new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Left },
                new CrossingAxis { Val = catAxisId }
            )
        );

        var chart = new C.Chart();

        if (xlChart.Title != null)
        {
            chart.Title = new Title(
                new ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text(xlChart.Title)
                            )
                        )
                    )
                ),
                new Overlay { Val = false }
            );
        }

        chart.Append(plotArea);
        chart.Append(new PlotVisibleOnly { Val = true });

        return new ChartSpace(chart);
    }

    /// <summary>
    /// Maps the chart's <see cref="XLBarOrientation"/> to the OpenXML <see cref="BarDirectionValues"/>.
    /// </summary>
    private static BarDirectionValues GetBarDirection(XLChart xlChart)
    {
        return xlChart.BarOrientation == XLBarOrientation.Horizontal
            ? BarDirectionValues.Bar
            : BarDirectionValues.Column;
    }

    /// <summary>
    /// Maps the chart's <see cref="XLBarGrouping"/> to the OpenXML <see cref="BarGroupingValues"/>.
    /// </summary>
    private static BarGroupingValues GetGrouping(XLChart xlChart)
    {
        return xlChart.BarGrouping switch
        {
            XLBarGrouping.Clustered => BarGroupingValues.Clustered,
            XLBarGrouping.Stacked => BarGroupingValues.Stacked,
            XLBarGrouping.Percent => BarGroupingValues.PercentStacked,
            _ => BarGroupingValues.Standard
        };
    }

    /// <summary>
    /// Ensures the WorksheetDrawing element has the required DrawingML and relationships namespace declarations.
    /// </summary>
    private static void EnsureNamespaces(Xdr.WorksheetDrawing worksheetDrawing)
    {
        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/main")))
            worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
            worksheetDrawing.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    }
}
