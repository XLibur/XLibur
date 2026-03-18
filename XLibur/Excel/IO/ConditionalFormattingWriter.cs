using XLibur.Excel.ContentManagers;
using XLibur.Extensions;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.ConditionalFormats;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class ConditionalFormattingWriter
{
    internal static void WriteConditionalFormatting(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        SaveContext context)
    {
        var xlSheetPivotCfs = xlWorksheet.PivotTables
            .SelectMany<XLPivotTable, XLConditionalFormat>(pt => pt.ConditionalFormats.Select(cf => cf.Format))
            .ToHashSet();

        // Elements in sheet.ConditionalFormats were sorted according to priority during load,
        // but new ones have priority 0. CFs are also interleaved with sheet CF. To deal with
        // these situations, set correct unique priority (also required for pivot CF).
        var xlConditionalFormats = xlWorksheet.ConditionalFormats.Cast<XLConditionalFormat>()
            .Concat(xlSheetPivotCfs)
            .OrderBy(x => x.Priority)
            .ToList();
        for (var i = 0; i < xlConditionalFormats.Count; ++i)
            xlConditionalFormats[i].Priority = i + 1;

        if (xlConditionalFormats.Count == 0)
        {
            worksheet.RemoveAllChildren<ConditionalFormatting>();
            cm.SetElement(XLWorksheetContents.ConditionalFormatting, null);
        }
        else
        {
            worksheet.RemoveAllChildren<ConditionalFormatting>();
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.ConditionalFormatting);

            foreach (var cfGroup in xlConditionalFormats
                         .GroupBy(
                             c => new
                             {
                                 SeqRefs = string.Join(" ",
                                     c.Ranges.Select(r => r.RangeAddress.ToStringRelative(false))),
                                 IsPivot = xlSheetPivotCfs.Contains(c),
                             },
                             c => c,
                             (key, g) => new { key.SeqRefs, key.IsPivot, CfList = g.ToList() }
                         )
                    )
            {
                var conditionalFormatting = new ConditionalFormatting
                {
                    SequenceOfReferences =
                        new ListValue<StringValue> { InnerText = cfGroup.SeqRefs },
                    Pivot = cfGroup.IsPivot ? true : null,
                };
                foreach (var cf in cfGroup.CfList)
                {
                    var xlCf = XLCFConverters.Convert(cf, cf.Priority, context);
                    conditionalFormatting.Append(xlCf);
                }

                worksheet.InsertAfter(conditionalFormatting, previousElement);
                previousElement = conditionalFormatting;
                cm.SetElement(XLWorksheetContents.ConditionalFormatting, conditionalFormatting);
            }
        }

        WriteExtensionDataBars(worksheet, cm, xlWorksheet, context);
    }

    private static void WriteExtensionDataBars(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        SaveContext context)
    {
        var exlst = xlWorksheet.ConditionalFormats
            .Where(c => c.ConditionalFormatType == XLConditionalFormatType.DataBar).ToArray();
        if (exlst.Length > 0)
        {
            if (!worksheet.Elements<WorksheetExtensionList>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.WorksheetExtensionList);
                worksheet.InsertAfter(new WorksheetExtensionList(), previousElement);
            }

            var worksheetExtensionList = worksheet.Elements<WorksheetExtensionList>().First();
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, worksheetExtensionList);

            var conditionalFormattings = worksheetExtensionList
                .Descendants<DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattings>().SingleOrDefault();
            if (conditionalFormattings == null || !conditionalFormattings.Any())
            {
                var worksheetExtension1 = new WorksheetExtension { Uri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}" };
                worksheetExtension1.AddNamespaceDeclaration("x14", X14Main2009SsNs);
                worksheetExtensionList.Append(worksheetExtension1);

                conditionalFormattings = new DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattings();
                worksheetExtension1.Append(conditionalFormattings);
            }

            foreach (var cfGroup in exlst
                         .GroupBy(
                             c => string.Join(" ", c.Ranges.Select(r => r.RangeAddress.ToStringRelative(false))),
                             c => c,
                             (key, g) => new { RangeId = key, CfList = g.ToList() }
                         )
                    )
            {
                foreach (var xlConditionalFormat in cfGroup.CfList.Cast<XLConditionalFormat>())
                {
                    var conditionalFormattingRule = conditionalFormattings
                        .Descendants<DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattingRule>()
                        .SingleOrDefault(r => r.Id == xlConditionalFormat.Id.WrapInBraces());
                    if (conditionalFormattingRule != null)
                    {
                        var conditionalFormat = conditionalFormattingRule
                            .Ancestors<DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormatting>()
                            .SingleOrDefault();
                        conditionalFormattings.RemoveChild(conditionalFormat);
                    }

                    var conditionalFormatting = new DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormatting();
                    conditionalFormatting.AddNamespaceDeclaration("xm", XmMain2006);
                    conditionalFormatting.Append(XLCFConvertersExtension.Convert(xlConditionalFormat, context));
                    var referenceSequence = new OfficeExcel.ReferenceSequence
                    { Text = cfGroup.RangeId };
                    conditionalFormatting.Append(referenceSequence);

                    conditionalFormattings.Append(conditionalFormatting);
                }
            }
        }
    }

    internal static void WriteSparklines(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        const string sparklineGroupsExtensionUri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

        if (!xlWorksheet.SparklineGroups.Any())
        {
            RemoveSparklineExtension(worksheet, cm, sparklineGroupsExtensionUri);
        }
        else
        {
            WriteSparklineGroups(worksheet, cm, xlWorksheet, sparklineGroupsExtensionUri);
        }
    }

    private static void RemoveSparklineExtension(Worksheet worksheet, XLWorksheetContentManager cm,
        string sparklineGroupsExtensionUri)
    {
        var worksheetExtensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
        var worksheetExtension = worksheetExtensionList?.Elements<WorksheetExtension>()
            .FirstOrDefault(ext =>
                string.Equals(ext.Uri, sparklineGroupsExtensionUri, StringComparison.InvariantCultureIgnoreCase));

        worksheetExtension?.RemoveAllChildren<X14.SparklineGroups>();

        if (worksheetExtensionList == null)
            return;

        if (worksheetExtension is { HasChildren: false })
            worksheetExtensionList.RemoveChild(worksheetExtension);

        if (!worksheetExtensionList.HasChildren)
        {
            worksheet.RemoveChild(worksheetExtensionList);
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, null);
        }
    }

    private static void WriteSparklineGroups(Worksheet worksheet, XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet, string sparklineGroupsExtensionUri)
    {
        if (!worksheet.Elements<WorksheetExtensionList>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.WorksheetExtensionList);
            worksheet.InsertAfter(new WorksheetExtensionList(), previousElement);
        }

        var worksheetExtensionList = worksheet.Elements<WorksheetExtensionList>().First();
        cm.SetElement(XLWorksheetContents.WorksheetExtensionList, worksheetExtensionList);

        var sparklineGroups = worksheetExtensionList.Descendants<X14.SparklineGroups>().SingleOrDefault();

        if (sparklineGroups == null || !sparklineGroups.Any())
        {
            var worksheetExtension1 = new WorksheetExtension() { Uri = sparklineGroupsExtensionUri };
            worksheetExtension1.AddNamespaceDeclaration("x14", X14Main2009SsNs);
            worksheetExtensionList.Append(worksheetExtension1);

            sparklineGroups = new X14.SparklineGroups();
            sparklineGroups.AddNamespaceDeclaration("xm", XmMain2006);
            worksheetExtension1.Append(sparklineGroups);
        }
        else
        {
            sparklineGroups.RemoveAllChildren();
        }

        foreach (var xlSparklineGroup in xlWorksheet.SparklineGroups)
        {
            if (!xlSparklineGroup.Any())
                continue;

            var sparklineGroup = CreateSparklineGroup(xlSparklineGroup);
            sparklineGroups.Append(sparklineGroup);
        }

        if (sparklineGroups.ChildElements.Count == 0)
            sparklineGroups.Remove();
    }

    private static X14.SparklineGroup CreateSparklineGroup(IXLSparklineGroup xlSparklineGroup)
    {
        var sparklineGroup = new X14.SparklineGroup();
        sparklineGroup.SetAttribute(new OpenXmlAttribute("xr2", "uid",
            "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
            "{A98FF5F8-AE60-43B5-8001-AD89004F45D3}"));

        SetSparklineColors(sparklineGroup, xlSparklineGroup);
        SetSparklineMarkers(sparklineGroup, xlSparklineGroup);
        SetSparklineDisplayOptions(sparklineGroup, xlSparklineGroup);
        SetSparklineAxes(sparklineGroup, xlSparklineGroup);

        var sparklines = new X14.Sparklines(xlSparklineGroup
            .Select(xlSparkline => new X14.Sparkline
            {
                Formula = new OfficeExcel.Formula(
                    xlSparkline.SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true)),
                ReferenceSequence =
                    new OfficeExcel.ReferenceSequence(xlSparkline.Location.Address.ToString()!)
            })
        );

        sparklineGroup.Append(sparklines);
        return sparklineGroup;
    }

    private static void SetSparklineColors(X14.SparklineGroup sparklineGroup, IXLSparklineGroup xlSparklineGroup)
    {
        sparklineGroup.FirstMarkerColor =
            new X14.FirstMarkerColor().FromXLiburColor<X14.FirstMarkerColor>(xlSparklineGroup.Style.FirstMarkerColor);
        sparklineGroup.LastMarkerColor =
            new X14.LastMarkerColor().FromXLiburColor<X14.LastMarkerColor>(xlSparklineGroup.Style.LastMarkerColor);
        sparklineGroup.HighMarkerColor =
            new X14.HighMarkerColor().FromXLiburColor<X14.HighMarkerColor>(xlSparklineGroup.Style.HighMarkerColor);
        sparklineGroup.LowMarkerColor =
            new X14.LowMarkerColor().FromXLiburColor<X14.LowMarkerColor>(xlSparklineGroup.Style.LowMarkerColor);
        sparklineGroup.SeriesColor =
            new X14.SeriesColor().FromXLiburColor<X14.SeriesColor>(xlSparklineGroup.Style.SeriesColor);
        sparklineGroup.NegativeColor =
            new X14.NegativeColor().FromXLiburColor<X14.NegativeColor>(xlSparklineGroup.Style.NegativeColor);
        sparklineGroup.MarkersColor =
            new X14.MarkersColor().FromXLiburColor<X14.MarkersColor>(xlSparklineGroup.Style.MarkersColor);
    }

    private static void SetSparklineMarkers(X14.SparklineGroup sparklineGroup, IXLSparklineGroup xlSparklineGroup)
    {
        sparklineGroup.High = xlSparklineGroup.ShowMarkers.HasFlag(XLSparklineMarkers.HighPoint);
        sparklineGroup.Low = xlSparklineGroup.ShowMarkers.HasFlag(XLSparklineMarkers.LowPoint);
        sparklineGroup.First = xlSparklineGroup.ShowMarkers.HasFlag(XLSparklineMarkers.FirstPoint);
        sparklineGroup.Last = xlSparklineGroup.ShowMarkers.HasFlag(XLSparklineMarkers.LastPoint);
        sparklineGroup.Negative = xlSparklineGroup.ShowMarkers.HasFlag(XLSparklineMarkers.NegativePoints);
        sparklineGroup.Markers = xlSparklineGroup.ShowMarkers.HasFlag(XLSparklineMarkers.Markers);
    }

    private static void SetSparklineDisplayOptions(X14.SparklineGroup sparklineGroup, IXLSparklineGroup xlSparklineGroup)
    {
        sparklineGroup.DisplayHidden = xlSparklineGroup.DisplayHidden;
        sparklineGroup.LineWeight = xlSparklineGroup.LineWeight;
        sparklineGroup.Type = xlSparklineGroup.Type.ToOpenXml();
        sparklineGroup.DisplayEmptyCellsAs = xlSparklineGroup.DisplayEmptyCellsAs.ToOpenXml();
    }

    private static void SetSparklineAxes(X14.SparklineGroup sparklineGroup, IXLSparklineGroup xlSparklineGroup)
    {
        sparklineGroup.AxisColor = new X14.AxisColor()
        { Rgb = xlSparklineGroup.HorizontalAxis.Color.Color.ToHex() };
        sparklineGroup.DisplayXAxis = xlSparklineGroup.HorizontalAxis.IsVisible;
        sparklineGroup.RightToLeft = xlSparklineGroup.HorizontalAxis.RightToLeft;
        sparklineGroup.DateAxis = xlSparklineGroup.HorizontalAxis.DateAxis;
        if (xlSparklineGroup.HorizontalAxis.DateAxis)
            sparklineGroup.Formula = new OfficeExcel.Formula(
                xlSparklineGroup.DateRange!.RangeAddress.ToString(XLReferenceStyle.A1, true));

        sparklineGroup.MinAxisType = xlSparklineGroup.VerticalAxis.MinAxisType.ToOpenXml();
        if (xlSparklineGroup.VerticalAxis.MinAxisType == XLSparklineAxisMinMax.Custom)
            sparklineGroup.ManualMin = xlSparklineGroup.VerticalAxis.ManualMin;

        sparklineGroup.MaxAxisType = xlSparklineGroup.VerticalAxis.MaxAxisType.ToOpenXml();
        if (xlSparklineGroup.VerticalAxis.MaxAxisType == XLSparklineAxisMinMax.Custom)
            sparklineGroup.ManualMax = xlSparklineGroup.VerticalAxis.ManualMax;
    }
}
