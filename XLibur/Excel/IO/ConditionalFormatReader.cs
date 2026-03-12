using XLibur.Extensions;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads conditional formatting rules and worksheet extensions (sparklines, X14 data validations, data bars).
/// </summary>
internal static class ConditionalFormatReader
{
    /// <summary>
    /// Loads the conditional formatting.
    /// </summary>
    // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
    internal static void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XLWorksheet ws,
        Dictionary<int, DifferentialFormat> differentialFormats, LoadContext context)
    {
        if (conditionalFormatting == null) return;

        foreach (var fr in conditionalFormatting.Elements<ConditionalFormattingRule>())
        {
            var ranges = conditionalFormatting.SequenceOfReferences!.Items
                .Select(sor => ws.Range(sor.Value!)!);
            var conditionalFormat = new XLConditionalFormat(ranges);

            conditionalFormat.StopIfTrue = OpenXmlHelper.GetBooleanValueAsBool(fr.StopIfTrue, false);

            if (fr.FormatId != null)
            {
                OpenXmlHelper.LoadFont(differentialFormats[(int)fr.FormatId.Value].Font, conditionalFormat.Style.Font);
                OpenXmlHelper.LoadFill(differentialFormats[(int)fr.FormatId.Value].Fill, conditionalFormat.Style.Fill,
                    differentialFillFormat: true);
                OpenXmlHelper.LoadBorder(differentialFormats[(int)fr.FormatId.Value].Border,
                    conditionalFormat.Style.Border);
                OpenXmlHelper.LoadNumberFormat(differentialFormats[(int)fr.FormatId.Value].NumberingFormat,
                    conditionalFormat.Style.NumberFormat);
            }

            // The conditional formatting type is compulsory. If it doesn't exist, skip the entire rule.
            if (fr.Type == null) continue;
            conditionalFormat.ConditionalFormatType = fr.Type.Value.ToXLibur();
            conditionalFormat.Priority = fr.Priority?.Value ?? int.MaxValue;

            // Although formulas are directly used only by CellIs and Expression type, other
            // format types also write them for evaluation to the workbook, e.g. rule to
            // IsBlank writes `LEN(TRIM(A2))=0` or ContainsText writes `NOT(ISERROR(SEARCH("hello",A2)))`.
            if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.CellIs)
            {
                conditionalFormat.Operator = fr.Operator!.Value.ToXLibur();

                // The XML schema allows up to three <formula> tags, but at most two are used.
                // Some producers emit empty <formula> tags that should be ignored and extra
                // non-empty formulas should also be ignored (Excel behavior).
                var nonEmptyFormulas = fr.Elements<Formula>()
                    .Where(static f => !string.IsNullOrEmpty(f.Text))
                    .Select(f => GetFormula(f.Text!))
                    .ToList();
                if (conditionalFormat.Operator is XLCFOperator.Between or XLCFOperator.NotBetween)
                {
                    var formulas = nonEmptyFormulas.Take(2).ToList();
                    if (formulas.Count != 2)
                        throw PartStructureException.IncorrectElementsCount();

                    conditionalFormat.Values.Add(formulas[0]);
                    conditionalFormat.Values.Add(formulas[1]);
                }
                else
                {
                    // Other XLCFOperators expect one argument.
                    var operatorArg = nonEmptyFormulas.FirstOrDefault();
                    if (operatorArg is null)
                        throw PartStructureException.IncorrectElementsCount();

                    conditionalFormat.Values.Add(operatorArg);
                }
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Expression)
            {
                var formula = fr.Elements<Formula>()
                    .Where(static f => !string.IsNullOrEmpty(f.Text))
                    .FirstOrDefault();

                if (formula is null)
                    throw PartStructureException.IncorrectElementsCount();

                conditionalFormat.Values.Add(GetFormula(formula.Text!));
            }

            if (!string.IsNullOrWhiteSpace(fr.Text))
                conditionalFormat.Values.Add(GetFormula(fr.Text!.Value!));

            if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Top10)
            {
                if (fr.Percent != null)
                    conditionalFormat.Percent = fr.Percent.Value;
                if (fr.Bottom != null)
                    conditionalFormat.Bottom = fr.Bottom.Value;
                if (fr.Rank != null)
                    conditionalFormat.Values.Add(GetFormula(fr.Rank.Value.ToString()));
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.TimePeriod)
            {
                if (fr.TimePeriod != null)
                    conditionalFormat.TimePeriod = fr.TimePeriod.Value.ToXLibur();
                else
                    conditionalFormat.TimePeriod = XLTimePeriod.Yesterday;
            }

            if (fr.Elements<ColorScale>().Any())
            {
                var colorScale = fr.Elements<ColorScale>().First();
                ExtractConditionalFormatValueObjects(conditionalFormat, colorScale);
            }
            else if (fr.Elements<DataBar>().Any())
            {
                var dataBar = fr.Elements<DataBar>().First();
                if (dataBar.ShowValue != null)
                    conditionalFormat.ShowBarOnly = !dataBar.ShowValue.Value;

                var id = fr.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Id>().FirstOrDefault();
                if (id is { Text: not null } && !string.IsNullOrWhiteSpace(id.Text))
                    conditionalFormat.Id = new Guid(id.Text.Substring(1, id.Text.Length - 2));

                ExtractConditionalFormatValueObjects(conditionalFormat, dataBar);
            }
            else if (fr.Elements<IconSet>().Any())
            {
                var iconSet = fr.Elements<IconSet>().First();
                if (iconSet.ShowValue != null)
                    conditionalFormat.ShowIconOnly = !iconSet.ShowValue.Value;
                if (iconSet.Reverse != null)
                    conditionalFormat.ReverseIconOrder = iconSet.Reverse.Value;

                if (iconSet.IconSetValue != null)
                    conditionalFormat.IconSetStyle = iconSet.IconSetValue.Value.ToXLibur();
                else
                    conditionalFormat.IconSetStyle = XLIconSetStyle.ThreeTrafficLights1;

                ExtractConditionalFormatValueObjects(conditionalFormat, iconSet);
            }

            var isPivotTableFormatting = conditionalFormatting.Pivot?.Value ?? false;
            if (isPivotTableFormatting)
                context.AddPivotTableCf(ws.Name, conditionalFormat);
            else
                ws.ConditionalFormats.Add(conditionalFormat);
        }
    }

    internal static void LoadExtensions(WorksheetExtensionList extensions, XLWorksheet ws, XLWorkbook workbook)
    {
        if (extensions == null)
        {
            return;
        }

        foreach (var dvs in extensions
                     .Descendants<X14.DataValidations>()
                     .SelectMany(dataValidations => dataValidations.Descendants<X14.DataValidation>()))
        {
            var txt = dvs.ReferenceSequence!.InnerText;
            if (string.IsNullOrWhiteSpace(txt)) continue;
            foreach (var rangeAddress in txt.Split(' '))
            {
                var dvt = new XLDataValidation(ws.Range(rangeAddress)!);
                ws.DataValidations.Add(dvt, skipIntersectionsCheck: true);
                if (dvs.AllowBlank != null) dvt.IgnoreBlanks = dvs.AllowBlank;
                if (dvs.ShowDropDown != null) dvt.InCellDropdown = !dvs.ShowDropDown.Value;
                if (dvs.ShowErrorMessage != null) dvt.ShowErrorMessage = dvs.ShowErrorMessage;
                if (dvs.ShowInputMessage != null) dvt.ShowInputMessage = dvs.ShowInputMessage;
                if (dvs.PromptTitle != null) dvt.InputTitle = dvs.PromptTitle.Value!;
                if (dvs.Prompt != null) dvt.InputMessage = dvs.Prompt.Value!;
                if (dvs.ErrorTitle != null) dvt.ErrorTitle = dvs.ErrorTitle.Value!;
                if (dvs.Error != null) dvt.ErrorMessage = dvs.Error.Value!;
                if (dvs.ErrorStyle != null) dvt.ErrorStyle = dvs.ErrorStyle.Value.ToXLibur();
                if (dvs.Type != null) dvt.AllowedValues = dvs.Type.Value.ToXLibur();
                if (dvs.Operator != null) dvt.Operator = dvs.Operator.Value.ToXLibur();
                if (dvs.DataValidationForumla1 != null) dvt.MinValue = dvs.DataValidationForumla1.InnerText;
                if (dvs.DataValidationForumla2 != null) dvt.MaxValue = dvs.DataValidationForumla2.InnerText;
            }
        }

        foreach (var conditionalFormattingRule in extensions
                     .Descendants<DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattingRule>()
                     .Where(cf =>
                         cf.Type is { HasValue: true }
                         && cf.Type.Value == ConditionalFormatValues.DataBar))
        {
            var xlConditionalFormat = ws.ConditionalFormats
                .Cast<XLConditionalFormat>()
                .SingleOrDefault(cf => cf.Id.WrapInBraces() == conditionalFormattingRule.Id);
            if (xlConditionalFormat != null)
            {
                var negativeFillColor = conditionalFormattingRule
                    .Descendants<DocumentFormat.OpenXml.Office2010.Excel.NegativeFillColor>().SingleOrDefault();
                xlConditionalFormat.Colors.Add(negativeFillColor!.ToXLiburColor());

                var x14DataBar = conditionalFormattingRule
                    .Descendants<DocumentFormat.OpenXml.Office2010.Excel.DataBar>().SingleOrDefault();
                if (x14DataBar?.Gradient != null)
                    xlConditionalFormat.Gradient = x14DataBar.Gradient.Value;
            }
        }

        foreach (var slg in extensions
                     .Descendants<X14.SparklineGroups>()
                     .SelectMany(sparklineGroups => sparklineGroups.Descendants<X14.SparklineGroup>()))
        {
            var xlSparklineGroup = ((XLSparklineGroups)ws.SparklineGroups).Add();

            if (slg.Formula != null)
                xlSparklineGroup.DateRange = workbook.Range(slg.Formula.Text);

            var xlSparklineStyle = xlSparklineGroup.Style;
            if (slg.FirstMarkerColor != null)
                xlSparklineStyle.FirstMarkerColor = slg.FirstMarkerColor.ToXLiburColor();
            if (slg.LastMarkerColor != null) xlSparklineStyle.LastMarkerColor = slg.LastMarkerColor.ToXLiburColor();
            if (slg.HighMarkerColor != null) xlSparklineStyle.HighMarkerColor = slg.HighMarkerColor.ToXLiburColor();
            if (slg.LowMarkerColor != null) xlSparklineStyle.LowMarkerColor = slg.LowMarkerColor.ToXLiburColor();
            if (slg.SeriesColor != null) xlSparklineStyle.SeriesColor = slg.SeriesColor.ToXLiburColor();
            if (slg.NegativeColor != null) xlSparklineStyle.NegativeColor = slg.NegativeColor.ToXLiburColor();
            if (slg.MarkersColor != null) xlSparklineStyle.MarkersColor = slg.MarkersColor.ToXLiburColor();
            xlSparklineGroup.Style = xlSparklineStyle;

            if (slg.DisplayHidden != null) xlSparklineGroup.DisplayHidden = slg.DisplayHidden;
            if (slg.LineWeight != null) xlSparklineGroup.LineWeight = slg.LineWeight;
            if (slg.Type != null) xlSparklineGroup.Type = slg.Type.Value.ToXLibur();
            if (slg.DisplayEmptyCellsAs != null)
                xlSparklineGroup.DisplayEmptyCellsAs = slg.DisplayEmptyCellsAs.Value.ToXLibur();

            xlSparklineGroup.ShowMarkers = XLSparklineMarkers.None;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Markers, false))
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.Markers;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.High, false))
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.HighPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Low, false))
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.LowPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.First, false))
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.FirstPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Last, false))
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.LastPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Negative, false))
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.NegativePoints;

            if (slg.AxisColor != null)
                xlSparklineGroup.HorizontalAxis.Color = XLColor.FromHtml(slg.AxisColor.Rgb!.Value!);
            if (slg.DisplayXAxis != null) xlSparklineGroup.HorizontalAxis.IsVisible = slg.DisplayXAxis;
            if (slg.RightToLeft != null) xlSparklineGroup.HorizontalAxis.RightToLeft = slg.RightToLeft;

            if (slg.ManualMax != null) xlSparklineGroup.VerticalAxis.ManualMax = slg.ManualMax;
            if (slg.ManualMin != null) xlSparklineGroup.VerticalAxis.ManualMin = slg.ManualMin;
            if (slg.MinAxisType != null)
                xlSparklineGroup.VerticalAxis.MinAxisType = slg.MinAxisType.Value.ToXLibur();
            if (slg.MaxAxisType != null)
                xlSparklineGroup.VerticalAxis.MaxAxisType = slg.MaxAxisType.Value.ToXLibur();

            slg.Descendants<X14.Sparklines>().SelectMany(sls => sls.Descendants<X14.Sparkline>())
                .ForEach(sl => xlSparklineGroup.Add(sl.ReferenceSequence!.Text!, sl.Formula!.Text!));
        }
    }

    internal static XLFormula GetFormula(string value)
    {
        var formula = new XLFormula();
        formula._value = value;
        formula.IsFormula = !(value[0] == '"' && value.EndsWith("\""));
        return formula;
    }

    internal static void ExtractConditionalFormatValueObjects(XLConditionalFormat conditionalFormat,
        OpenXmlElement element)
    {
        foreach (var c in element.Elements<ConditionalFormatValueObject>())
        {
            if (c.Type != null)
                conditionalFormat.ContentTypes.Add(c.Type.Value.ToXLibur());
            conditionalFormat.Values.Add(c.Val != null ? new XLFormula { Value = c.Val!.Value! } : null!);

            if (c.GreaterThanOrEqual != null)
                conditionalFormat.IconSetOperators.Add(c.GreaterThanOrEqual.Value
                    ? XLCFIconSetOperator.EqualOrGreaterThan
                    : XLCFIconSetOperator.GreaterThan);
            else
                conditionalFormat.IconSetOperators.Add(XLCFIconSetOperator.EqualOrGreaterThan);
        }

        foreach (var c in element.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>())
        {
            conditionalFormat.Colors.Add(c.ToXLiburColor());
        }
    }
}
