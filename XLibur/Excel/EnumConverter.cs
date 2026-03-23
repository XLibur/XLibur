using XLibur.Excel.Drawings;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using Vml = DocumentFormat.OpenXml.Vml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel;

internal static class EnumConverter
{
    private static readonly Dictionary<XLPictureFormat, PartTypeInfo> PictureFormatMap =
        new()
        {
            { XLPictureFormat.Unknown, new PartTypeInfo("image/unknown", ".bin") },
            { XLPictureFormat.Bmp, ImagePartType.Bmp },
            { XLPictureFormat.Gif, ImagePartType.Gif },
            { XLPictureFormat.Png, ImagePartType.Png },
            { XLPictureFormat.Tiff, ImagePartType.Tiff },
            { XLPictureFormat.Icon, ImagePartType.Icon },
            { XLPictureFormat.Pcx, ImagePartType.Pcx },
            { XLPictureFormat.Jpeg, ImagePartType.Jpeg },
            { XLPictureFormat.Emf, ImagePartType.Emf },
            { XLPictureFormat.Wmf, ImagePartType.Wmf },
            { XLPictureFormat.Webp, new PartTypeInfo("image/webp", ".webp") },
            { XLPictureFormat.Svg, new PartTypeInfo("image/svg+xml", ".svg") }
        };

    private static readonly string[] XLFontUnderlineValuesStrings =
    [
        "double",
        "doubleAccounting",
        "none",
        "single",
        "singleAccounting"
    ];

    private static readonly string[] XLFontVerticalTextAlignmentValuesStrings =
    [
        "baseline",
        "subscript",
        "superscript"
    ];

    private static readonly string[] XLPhoneticAlignmentStrings =
    [
        "center",
        "distributed",
        "left",
        "noControl"
    ];

    private static readonly string[] XLPhoneticTypeStrings =
    [
        "fullwidthKatakana",
        "halfwidthKatakana",
        "Hiragana",
        "noConversion"
    ];

    private static readonly string[] XLFontSchemeStrings =
    [
        "none",
        "major",
        "minor"
    ];

    private static readonly Dictionary<UnderlineValues, XLFontUnderlineValues> UnderlineValuesMap =
        new()
        {
            { UnderlineValues.Double, XLFontUnderlineValues.Double },
            { UnderlineValues.DoubleAccounting, XLFontUnderlineValues.DoubleAccounting },
            { UnderlineValues.None, XLFontUnderlineValues.None },
            { UnderlineValues.Single, XLFontUnderlineValues.Single },
            { UnderlineValues.SingleAccounting, XLFontUnderlineValues.SingleAccounting },
        };

    private static readonly Dictionary<FontSchemeValues, XLFontScheme> FontSchemeMap =
        new()
        {
            { FontSchemeValues.None, XLFontScheme.None },
            { FontSchemeValues.Major, XLFontScheme.Major },
            { FontSchemeValues.Minor, XLFontScheme.Minor },
        };

    private static readonly Dictionary<OrientationValues, XLPageOrientation> OrientationMap =
        new()
        {
            { OrientationValues.Default, XLPageOrientation.Default },
            { OrientationValues.Landscape, XLPageOrientation.Landscape },
            { OrientationValues.Portrait, XLPageOrientation.Portrait },
        };

    private static readonly Dictionary<VerticalAlignmentRunValues, XLFontVerticalTextAlignmentValues>
        VerticalAlignmentRunMap =
            new()
            {
                { VerticalAlignmentRunValues.Baseline, XLFontVerticalTextAlignmentValues.Baseline },
                { VerticalAlignmentRunValues.Subscript, XLFontVerticalTextAlignmentValues.Subscript },
                { VerticalAlignmentRunValues.Superscript, XLFontVerticalTextAlignmentValues.Superscript },
            };

    private static readonly Dictionary<PatternValues, XLFillPatternValues> PatternMap =
        new()
        {
            { PatternValues.DarkDown, XLFillPatternValues.DarkDown },
            { PatternValues.DarkGray, XLFillPatternValues.DarkGray },
            { PatternValues.DarkGrid, XLFillPatternValues.DarkGrid },
            { PatternValues.DarkHorizontal, XLFillPatternValues.DarkHorizontal },
            { PatternValues.DarkTrellis, XLFillPatternValues.DarkTrellis },
            { PatternValues.DarkUp, XLFillPatternValues.DarkUp },
            { PatternValues.DarkVertical, XLFillPatternValues.DarkVertical },
            { PatternValues.Gray0625, XLFillPatternValues.Gray0625 },
            { PatternValues.Gray125, XLFillPatternValues.Gray125 },
            { PatternValues.LightDown, XLFillPatternValues.LightDown },
            { PatternValues.LightGray, XLFillPatternValues.LightGray },
            { PatternValues.LightGrid, XLFillPatternValues.LightGrid },
            { PatternValues.LightHorizontal, XLFillPatternValues.LightHorizontal },
            { PatternValues.LightTrellis, XLFillPatternValues.LightTrellis },
            { PatternValues.LightUp, XLFillPatternValues.LightUp },
            { PatternValues.LightVertical, XLFillPatternValues.LightVertical },
            { PatternValues.MediumGray, XLFillPatternValues.MediumGray },
            { PatternValues.None, XLFillPatternValues.None },
            { PatternValues.Solid, XLFillPatternValues.Solid },
        };

    private static readonly Dictionary<BorderStyleValues, XLBorderStyleValues> BorderStyleMap =
        new()
        {
            { BorderStyleValues.DashDot, XLBorderStyleValues.DashDot },
            { BorderStyleValues.DashDotDot, XLBorderStyleValues.DashDotDot },
            { BorderStyleValues.Dashed, XLBorderStyleValues.Dashed },
            { BorderStyleValues.Dotted, XLBorderStyleValues.Dotted },
            { BorderStyleValues.Double, XLBorderStyleValues.Double },
            { BorderStyleValues.Hair, XLBorderStyleValues.Hair },
            { BorderStyleValues.Medium, XLBorderStyleValues.Medium },
            { BorderStyleValues.MediumDashDot, XLBorderStyleValues.MediumDashDot },
            { BorderStyleValues.MediumDashDotDot, XLBorderStyleValues.MediumDashDotDot },
            { BorderStyleValues.MediumDashed, XLBorderStyleValues.MediumDashed },
            { BorderStyleValues.None, XLBorderStyleValues.None },
            { BorderStyleValues.SlantDashDot, XLBorderStyleValues.SlantDashDot },
            { BorderStyleValues.Thick, XLBorderStyleValues.Thick },
            { BorderStyleValues.Thin, XLBorderStyleValues.Thin },
        };

    private static readonly Dictionary<HorizontalAlignmentValues, XLAlignmentHorizontalValues>
        HorizontalAlignmentMap =
            new()
            {
                { HorizontalAlignmentValues.Center, XLAlignmentHorizontalValues.Center },
                { HorizontalAlignmentValues.CenterContinuous, XLAlignmentHorizontalValues.CenterContinuous },
                { HorizontalAlignmentValues.Distributed, XLAlignmentHorizontalValues.Distributed },
                { HorizontalAlignmentValues.Fill, XLAlignmentHorizontalValues.Fill },
                { HorizontalAlignmentValues.General, XLAlignmentHorizontalValues.General },
                { HorizontalAlignmentValues.Justify, XLAlignmentHorizontalValues.Justify },
                { HorizontalAlignmentValues.Left, XLAlignmentHorizontalValues.Left },
                { HorizontalAlignmentValues.Right, XLAlignmentHorizontalValues.Right },
            };

    private static readonly Dictionary<VerticalAlignmentValues, XLAlignmentVerticalValues>
        VerticalAlignmentMap =
            new()
            {
                { VerticalAlignmentValues.Bottom, XLAlignmentVerticalValues.Bottom },
                { VerticalAlignmentValues.Center, XLAlignmentVerticalValues.Center },
                { VerticalAlignmentValues.Distributed, XLAlignmentVerticalValues.Distributed },
                { VerticalAlignmentValues.Justify, XLAlignmentVerticalValues.Justify },
                { VerticalAlignmentValues.Top, XLAlignmentVerticalValues.Top },
            };

    private static readonly Dictionary<PageOrderValues, XLPageOrderValues> PageOrdersMap =
        new()
        {
            { PageOrderValues.DownThenOver, XLPageOrderValues.DownThenOver },
            { PageOrderValues.OverThenDown, XLPageOrderValues.OverThenDown },
        };

    private static readonly Dictionary<CellCommentsValues, XLShowCommentsValues> CellCommentsMap =
        new()
        {
            { CellCommentsValues.AsDisplayed, XLShowCommentsValues.AsDisplayed },
            { CellCommentsValues.AtEnd, XLShowCommentsValues.AtEnd },
            { CellCommentsValues.None, XLShowCommentsValues.None },
        };

    private static readonly Dictionary<PrintErrorValues, XLPrintErrorValues> PrintErrorMap =
        new()
        {
            { PrintErrorValues.Blank, XLPrintErrorValues.Blank },
            { PrintErrorValues.Dash, XLPrintErrorValues.Dash },
            { PrintErrorValues.Displayed, XLPrintErrorValues.Displayed },
            { PrintErrorValues.NA, XLPrintErrorValues.NA },
        };

    private static readonly Dictionary<CalculateModeValues, XLCalculateMode> CalculateModeMap =
        new()
        {
            { CalculateModeValues.Auto, XLCalculateMode.Auto },
            { CalculateModeValues.AutoNoTable, XLCalculateMode.AutoNoTable },
            { CalculateModeValues.Manual, XLCalculateMode.Manual },
        };

    private static readonly Dictionary<ReferenceModeValues, XLReferenceStyle> ReferenceModeMap =
        new()
        {
            { ReferenceModeValues.R1C1, XLReferenceStyle.R1C1 },
            { ReferenceModeValues.A1, XLReferenceStyle.A1 },
        };

    private static readonly Dictionary<TotalsRowFunctionValues, XLTotalsRowFunction> TotalsRowFunctionMap =
        new()
        {
            { TotalsRowFunctionValues.None, XLTotalsRowFunction.None },
            { TotalsRowFunctionValues.Sum, XLTotalsRowFunction.Sum },
            { TotalsRowFunctionValues.Minimum, XLTotalsRowFunction.Minimum },
            { TotalsRowFunctionValues.Maximum, XLTotalsRowFunction.Maximum },
            { TotalsRowFunctionValues.Average, XLTotalsRowFunction.Average },
            { TotalsRowFunctionValues.Count, XLTotalsRowFunction.Count },
            { TotalsRowFunctionValues.CountNumbers, XLTotalsRowFunction.CountNumbers },
            { TotalsRowFunctionValues.StandardDeviation, XLTotalsRowFunction.StandardDeviation },
            { TotalsRowFunctionValues.Variance, XLTotalsRowFunction.Variance },
            { TotalsRowFunctionValues.Custom, XLTotalsRowFunction.Custom },
        };

    private static readonly Dictionary<DataValidationValues, XLAllowedValues> DataValidationMap =
        new()
        {
            { DataValidationValues.None, XLAllowedValues.AnyValue },
            { DataValidationValues.Custom, XLAllowedValues.Custom },
            { DataValidationValues.Date, XLAllowedValues.Date },
            { DataValidationValues.Decimal, XLAllowedValues.Decimal },
            { DataValidationValues.List, XLAllowedValues.List },
            { DataValidationValues.TextLength, XLAllowedValues.TextLength },
            { DataValidationValues.Time, XLAllowedValues.Time },
            { DataValidationValues.Whole, XLAllowedValues.WholeNumber },
        };

    private static readonly Dictionary<DataValidationErrorStyleValues, XLErrorStyle>
        DataValidationErrorStyleMap =
            new()
            {
                { DataValidationErrorStyleValues.Information, XLErrorStyle.Information },
                { DataValidationErrorStyleValues.Warning, XLErrorStyle.Warning },
                { DataValidationErrorStyleValues.Stop, XLErrorStyle.Stop },
            };

    private static readonly Dictionary<DataValidationOperatorValues, XLOperator> DataValidationOperatorMap =
        new()
        {
            { DataValidationOperatorValues.Between, XLOperator.Between },
            { DataValidationOperatorValues.GreaterThanOrEqual, XLOperator.EqualOrGreaterThan },
            { DataValidationOperatorValues.LessThanOrEqual, XLOperator.EqualOrLessThan },
            { DataValidationOperatorValues.Equal, XLOperator.EqualTo },
            { DataValidationOperatorValues.GreaterThan, XLOperator.GreaterThan },
            { DataValidationOperatorValues.LessThan, XLOperator.LessThan },
            { DataValidationOperatorValues.NotBetween, XLOperator.NotBetween },
            { DataValidationOperatorValues.NotEqual, XLOperator.NotEqualTo },
        };

    private static readonly Dictionary<SheetStateValues, XLWorksheetVisibility> SheetStateMap =
        new()
        {
            { SheetStateValues.Visible, XLWorksheetVisibility.Visible },
            { SheetStateValues.Hidden, XLWorksheetVisibility.Hidden },
            { SheetStateValues.VeryHidden, XLWorksheetVisibility.VeryHidden },
        };

    private static readonly Dictionary<PhoneticAlignmentValues, XLPhoneticAlignment> PhoneticAlignmentMap =
        new()
        {
            { PhoneticAlignmentValues.Center, XLPhoneticAlignment.Center },
            { PhoneticAlignmentValues.Distributed, XLPhoneticAlignment.Distributed },
            { PhoneticAlignmentValues.Left, XLPhoneticAlignment.Left },
            { PhoneticAlignmentValues.NoControl, XLPhoneticAlignment.NoControl },
        };

    private static readonly Dictionary<PhoneticValues, XLPhoneticType> PhoneticMap =
        new()
        {
            { PhoneticValues.FullWidthKatakana, XLPhoneticType.FullWidthKatakana },
            { PhoneticValues.HalfWidthKatakana, XLPhoneticType.HalfWidthKatakana },
            { PhoneticValues.Hiragana, XLPhoneticType.Hiragana },
            { PhoneticValues.NoConversion, XLPhoneticType.NoConversion },
        };

    private static readonly Dictionary<DataConsolidateFunctionValues, XLPivotSummary>
        DataConsolidateFunctionMap =
            new()
            {
                { DataConsolidateFunctionValues.Sum, XLPivotSummary.Sum },
                { DataConsolidateFunctionValues.Count, XLPivotSummary.Count },
                { DataConsolidateFunctionValues.Average, XLPivotSummary.Average },
                { DataConsolidateFunctionValues.Minimum, XLPivotSummary.Minimum },
                { DataConsolidateFunctionValues.Maximum, XLPivotSummary.Maximum },
                { DataConsolidateFunctionValues.Product, XLPivotSummary.Product },
                { DataConsolidateFunctionValues.CountNumbers, XLPivotSummary.CountNumbers },
                { DataConsolidateFunctionValues.StandardDeviation, XLPivotSummary.StandardDeviation },
                { DataConsolidateFunctionValues.StandardDeviationP, XLPivotSummary.PopulationStandardDeviation },
                { DataConsolidateFunctionValues.Variance, XLPivotSummary.Variance },
                { DataConsolidateFunctionValues.VarianceP, XLPivotSummary.PopulationVariance },
            };

    private static readonly Dictionary<ShowDataAsValues, XLPivotCalculation> ShowDataAsMap =
        new()
        {
            { ShowDataAsValues.Normal, XLPivotCalculation.Normal },
            { ShowDataAsValues.Difference, XLPivotCalculation.DifferenceFrom },
            { ShowDataAsValues.Percent, XLPivotCalculation.PercentageOf },
            { ShowDataAsValues.PercentageDifference, XLPivotCalculation.PercentageDifferenceFrom },
            { ShowDataAsValues.RunTotal, XLPivotCalculation.RunningTotal },
            {
                ShowDataAsValues.PercentOfRaw, XLPivotCalculation.PercentageOfRow
            }, // There's a typo in the OpenXML SDK =)
            { ShowDataAsValues.PercentOfColumn, XLPivotCalculation.PercentageOfColumn },
            { ShowDataAsValues.PercentOfTotal, XLPivotCalculation.PercentageOfTotal },
            { ShowDataAsValues.Index, XLPivotCalculation.Index },
        };

    private static readonly Dictionary<FilterOperatorValues, XLFilterOperator> FilterOperatorMap =
        new()
        {
            { FilterOperatorValues.Equal, XLFilterOperator.Equal },
            { FilterOperatorValues.NotEqual, XLFilterOperator.NotEqual },
            { FilterOperatorValues.GreaterThan, XLFilterOperator.GreaterThan },
            { FilterOperatorValues.LessThan, XLFilterOperator.LessThan },
            { FilterOperatorValues.GreaterThanOrEqual, XLFilterOperator.EqualOrGreaterThan },
            { FilterOperatorValues.LessThanOrEqual, XLFilterOperator.EqualOrLessThan },
        };

    private static readonly Dictionary<DynamicFilterValues, XLFilterDynamicType> DynamicFilterMap =
        new()
        {
            { DynamicFilterValues.AboveAverage, XLFilterDynamicType.AboveAverage },
            { DynamicFilterValues.BelowAverage, XLFilterDynamicType.BelowAverage },
        };

    private static readonly Dictionary<DateTimeGroupingValues, XLDateTimeGrouping> DateTimeGroupingMap =
        new()
        {
            { DateTimeGroupingValues.Year, XLDateTimeGrouping.Year },
            { DateTimeGroupingValues.Month, XLDateTimeGrouping.Month },
            { DateTimeGroupingValues.Day, XLDateTimeGrouping.Day },
            { DateTimeGroupingValues.Hour, XLDateTimeGrouping.Hour },
            { DateTimeGroupingValues.Minute, XLDateTimeGrouping.Minute },
            { DateTimeGroupingValues.Second, XLDateTimeGrouping.Second },
        };

    private static readonly Dictionary<SheetViewValues, XLSheetViewOptions> SheetViewMap =
        new()
        {
            { SheetViewValues.Normal, XLSheetViewOptions.Normal },
            { SheetViewValues.PageBreakPreview, XLSheetViewOptions.PageBreakPreview },
            { SheetViewValues.PageLayout, XLSheetViewOptions.PageLayout },
        };

    private static readonly Dictionary<Vml.StrokeLineStyleValues, XLLineStyle> StrokeLineStyleMap =
        new()
        {
            { Vml.StrokeLineStyleValues.Single, XLLineStyle.Single },
            { Vml.StrokeLineStyleValues.ThickBetweenThin, XLLineStyle.ThickBetweenThin },
            { Vml.StrokeLineStyleValues.ThickThin, XLLineStyle.ThickThin },
            { Vml.StrokeLineStyleValues.ThinThick, XLLineStyle.ThinThick },
            { Vml.StrokeLineStyleValues.ThinThin, XLLineStyle.ThinThin },
        };

    private static readonly Dictionary<ConditionalFormatValues, XLConditionalFormatType> ConditionalFormatMap =
        new()
        {
            { ConditionalFormatValues.Expression, XLConditionalFormatType.Expression },
            { ConditionalFormatValues.CellIs, XLConditionalFormatType.CellIs },
            { ConditionalFormatValues.ColorScale, XLConditionalFormatType.ColorScale },
            { ConditionalFormatValues.DataBar, XLConditionalFormatType.DataBar },
            { ConditionalFormatValues.IconSet, XLConditionalFormatType.IconSet },
            { ConditionalFormatValues.Top10, XLConditionalFormatType.Top10 },
            { ConditionalFormatValues.UniqueValues, XLConditionalFormatType.IsUnique },
            { ConditionalFormatValues.DuplicateValues, XLConditionalFormatType.IsDuplicate },
            { ConditionalFormatValues.ContainsText, XLConditionalFormatType.ContainsText },
            { ConditionalFormatValues.NotContainsText, XLConditionalFormatType.NotContainsText },
            { ConditionalFormatValues.BeginsWith, XLConditionalFormatType.StartsWith },
            { ConditionalFormatValues.EndsWith, XLConditionalFormatType.EndsWith },
            { ConditionalFormatValues.ContainsBlanks, XLConditionalFormatType.IsBlank },
            { ConditionalFormatValues.NotContainsBlanks, XLConditionalFormatType.NotBlank },
            { ConditionalFormatValues.ContainsErrors, XLConditionalFormatType.IsError },
            { ConditionalFormatValues.NotContainsErrors, XLConditionalFormatType.NotError },
            { ConditionalFormatValues.TimePeriod, XLConditionalFormatType.TimePeriod },
            { ConditionalFormatValues.AboveAverage, XLConditionalFormatType.AboveAverage },
        };

    private static readonly Dictionary<ConditionalFormatValueObjectValues, XLCFContentType>
        ConditionalFormatValueObjectMap =
            new()
            {
                { ConditionalFormatValueObjectValues.Number, XLCFContentType.Number },
                { ConditionalFormatValueObjectValues.Percent, XLCFContentType.Percent },
                { ConditionalFormatValueObjectValues.Max, XLCFContentType.Maximum },
                { ConditionalFormatValueObjectValues.Min, XLCFContentType.Minimum },
                { ConditionalFormatValueObjectValues.Formula, XLCFContentType.Formula },
                { ConditionalFormatValueObjectValues.Percentile, XLCFContentType.Percentile },
            };

    private static readonly Dictionary<ConditionalFormattingOperatorValues, XLCFOperator>
        ConditionalFormattingOperatorMap =
            new()
            {
                { ConditionalFormattingOperatorValues.LessThan, XLCFOperator.LessThan },
                { ConditionalFormattingOperatorValues.LessThanOrEqual, XLCFOperator.EqualOrLessThan },
                { ConditionalFormattingOperatorValues.Equal, XLCFOperator.Equal },
                { ConditionalFormattingOperatorValues.NotEqual, XLCFOperator.NotEqual },
                { ConditionalFormattingOperatorValues.GreaterThanOrEqual, XLCFOperator.EqualOrGreaterThan },
                { ConditionalFormattingOperatorValues.GreaterThan, XLCFOperator.GreaterThan },
                { ConditionalFormattingOperatorValues.Between, XLCFOperator.Between },
                { ConditionalFormattingOperatorValues.NotBetween, XLCFOperator.NotBetween },
                { ConditionalFormattingOperatorValues.ContainsText, XLCFOperator.Contains },
                { ConditionalFormattingOperatorValues.NotContains, XLCFOperator.NotContains },
                { ConditionalFormattingOperatorValues.BeginsWith, XLCFOperator.StartsWith },
                { ConditionalFormattingOperatorValues.EndsWith, XLCFOperator.EndsWith },
            };

    private static readonly Dictionary<IconSetValues, XLIconSetStyle> IconSetMap =
        new()
        {
            { IconSetValues.ThreeArrows, XLIconSetStyle.ThreeArrows },
            { IconSetValues.ThreeArrowsGray, XLIconSetStyle.ThreeArrowsGray },
            { IconSetValues.ThreeFlags, XLIconSetStyle.ThreeFlags },
            { IconSetValues.ThreeTrafficLights1, XLIconSetStyle.ThreeTrafficLights1 },
            { IconSetValues.ThreeTrafficLights2, XLIconSetStyle.ThreeTrafficLights2 },
            { IconSetValues.ThreeSigns, XLIconSetStyle.ThreeSigns },
            { IconSetValues.ThreeSymbols, XLIconSetStyle.ThreeSymbols },
            { IconSetValues.ThreeSymbols2, XLIconSetStyle.ThreeSymbols2 },
            { IconSetValues.FourArrows, XLIconSetStyle.FourArrows },
            { IconSetValues.FourArrowsGray, XLIconSetStyle.FourArrowsGray },
            { IconSetValues.FourRedToBlack, XLIconSetStyle.FourRedToBlack },
            { IconSetValues.FourRating, XLIconSetStyle.FourRating },
            { IconSetValues.FourTrafficLights, XLIconSetStyle.FourTrafficLights },
            { IconSetValues.FiveArrows, XLIconSetStyle.FiveArrows },
            { IconSetValues.FiveArrowsGray, XLIconSetStyle.FiveArrowsGray },
            { IconSetValues.FiveRating, XLIconSetStyle.FiveRating },
            { IconSetValues.FiveQuarters, XLIconSetStyle.FiveQuarters },
        };

    private static readonly Dictionary<TimePeriodValues, XLTimePeriod> TimePeriodMap =
        new()
        {
            { TimePeriodValues.Yesterday, XLTimePeriod.Yesterday },
            { TimePeriodValues.Today, XLTimePeriod.Today },
            { TimePeriodValues.Tomorrow, XLTimePeriod.Tomorrow },
            { TimePeriodValues.Last7Days, XLTimePeriod.InTheLast7Days },
            { TimePeriodValues.LastWeek, XLTimePeriod.LastWeek },
            { TimePeriodValues.ThisWeek, XLTimePeriod.ThisWeek },
            { TimePeriodValues.NextWeek, XLTimePeriod.NextWeek },
            { TimePeriodValues.LastMonth, XLTimePeriod.LastMonth },
            { TimePeriodValues.ThisMonth, XLTimePeriod.ThisMonth },
            { TimePeriodValues.NextMonth, XLTimePeriod.NextMonth },
        };

    private static readonly Dictionary<PivotAreaValues, XLPivotAreaType> PivotAreaMap =
        new()
        {
            { PivotAreaValues.None, XLPivotAreaType.None },
            { PivotAreaValues.Normal, XLPivotAreaType.Normal },
            { PivotAreaValues.Data, XLPivotAreaType.Data },
            { PivotAreaValues.All, XLPivotAreaType.All },
            { PivotAreaValues.Origin, XLPivotAreaType.Origin },
            { PivotAreaValues.Button, XLPivotAreaType.Button },
            { PivotAreaValues.TopRight, XLPivotAreaType.TopRight },
            { PivotAreaValues.TopEnd, XLPivotAreaType.TopEnd },
        };

    private static readonly Dictionary<X14.SparklineTypeValues, XLSparklineType> SparklineTypeMap =
        new()
        {
            { X14.SparklineTypeValues.Line, XLSparklineType.Line },
            { X14.SparklineTypeValues.Column, XLSparklineType.Column },
            { X14.SparklineTypeValues.Stacked, XLSparklineType.Stacked },
        };

    private static readonly Dictionary<X14.SparklineAxisMinMaxValues, XLSparklineAxisMinMax>
        SparklineAxisMinMaxMap =
            new()
            {
                { X14.SparklineAxisMinMaxValues.Individual, XLSparklineAxisMinMax.Automatic },
                { X14.SparklineAxisMinMaxValues.Group, XLSparklineAxisMinMax.SameForAll },
                { X14.SparklineAxisMinMaxValues.Custom, XLSparklineAxisMinMax.Custom },
            };

    private static readonly Dictionary<X14.DisplayBlanksAsValues, XLDisplayBlanksAsValues> DisplayBlanksAsMap =
        new()
        {
            { X14.DisplayBlanksAsValues.Span, XLDisplayBlanksAsValues.Interpolate },
            { X14.DisplayBlanksAsValues.Gap, XLDisplayBlanksAsValues.NotPlotted },
            { X14.DisplayBlanksAsValues.Zero, XLDisplayBlanksAsValues.Zero },
        };

    private static readonly Dictionary<FieldSortValues, XLPivotSortType> FieldSortMap =
        new()
        {
            { FieldSortValues.Manual, XLPivotSortType.Default },
            { FieldSortValues.Ascending, XLPivotSortType.Ascending },
            { FieldSortValues.Descending, XLPivotSortType.Descending },
        };

    private static readonly Dictionary<PivotTableAxisValues, XLPivotAxis> PivotTableAxisMap =
        new()
        {
            { PivotTableAxisValues.AxisRow, XLPivotAxis.AxisRow },
            { PivotTableAxisValues.AxisColumn, XLPivotAxis.AxisCol },
            { PivotTableAxisValues.AxisPage, XLPivotAxis.AxisPage },
            { PivotTableAxisValues.AxisValues, XLPivotAxis.AxisValues },
        };

    private static readonly Dictionary<ItemValues, XLPivotItemType> ItemMap =
        new()
        {
            { ItemValues.Data, XLPivotItemType.Data },
            { ItemValues.Default, XLPivotItemType.Default },
            { ItemValues.Sum, XLPivotItemType.Sum },
            { ItemValues.CountA, XLPivotItemType.CountA },
            { ItemValues.Average, XLPivotItemType.Avg },
            { ItemValues.Maximum, XLPivotItemType.Max },
            { ItemValues.Minimum, XLPivotItemType.Min },
            { ItemValues.Product, XLPivotItemType.Product },
            { ItemValues.Count, XLPivotItemType.Count },
            { ItemValues.StandardDeviation, XLPivotItemType.StdDev },
            { ItemValues.StandardDeviationP, XLPivotItemType.StdDevP },
            { ItemValues.Variance, XLPivotItemType.Var },
            { ItemValues.VarianceP, XLPivotItemType.VarP },
            { ItemValues.Grand, XLPivotItemType.Grand },
            { ItemValues.Blank, XLPivotItemType.Blank },
        };

    private static readonly Dictionary<FormatActionValues, XLPivotFormatAction> FormatActionMap =
        new()
        {
            { FormatActionValues.Blank, XLPivotFormatAction.Blank },
            { FormatActionValues.Formatting, XLPivotFormatAction.Formatting },
        };

    private static readonly Dictionary<ScopeValues, XLPivotCfScope> ScopeMap =
        new()
        {
            { ScopeValues.Selection, XLPivotCfScope.SelectedCells },
            { ScopeValues.Data, XLPivotCfScope.DataFields },
            { ScopeValues.Field, XLPivotCfScope.FieldIntersections },
        };

    private static readonly Dictionary<RuleValues, XLPivotCfRuleType> RuleMap =
        new()
        {
            { RuleValues.None, XLPivotCfRuleType.None },
            { RuleValues.All, XLPivotCfRuleType.All },
            { RuleValues.Row, XLPivotCfRuleType.Row },
            { RuleValues.Column, XLPivotCfRuleType.Column },
        };

    public static string ToOpenXmlString(this XLFontUnderlineValues value)
        => XLFontUnderlineValuesStrings[(int)value];

    public static string ToOpenXmlString(this XLFontVerticalTextAlignmentValues value)
        => XLFontVerticalTextAlignmentValuesStrings[(int)value];

    public static string ToOpenXmlString(this XLPhoneticAlignment value)
        => XLPhoneticAlignmentStrings[(int)value];

    public static string ToOpenXmlString(this XLPhoneticType value)
        => XLPhoneticTypeStrings[(int)value];

    extension(XLFontScheme value)
    {
        public string ToOpenXml()
            => XLFontSchemeStrings[(int)value];

        public FontSchemeValues ToOpenXmlEnum()
        {
            return value switch
            {
                XLFontScheme.None => FontSchemeValues.None,
                XLFontScheme.Major => FontSchemeValues.Major,
                XLFontScheme.Minor => FontSchemeValues.Minor,
                _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported font scheme value.")
            };
        }
    }

    public static UnderlineValues ToOpenXml(this XLFontUnderlineValues value) => value switch
    {
        XLFontUnderlineValues.Double => UnderlineValues.Double,
        XLFontUnderlineValues.DoubleAccounting => UnderlineValues.DoubleAccounting,
        XLFontUnderlineValues.None => UnderlineValues.None,
        XLFontUnderlineValues.Single => UnderlineValues.Single,
        XLFontUnderlineValues.SingleAccounting => UnderlineValues.SingleAccounting,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static OrientationValues ToOpenXml(this XLPageOrientation value) => value switch
    {
        XLPageOrientation.Default => OrientationValues.Default,
        XLPageOrientation.Landscape => OrientationValues.Landscape,
        XLPageOrientation.Portrait => OrientationValues.Portrait,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static VerticalAlignmentRunValues ToOpenXml(this XLFontVerticalTextAlignmentValues value) => value switch
    {
        XLFontVerticalTextAlignmentValues.Baseline => VerticalAlignmentRunValues.Baseline,
        XLFontVerticalTextAlignmentValues.Subscript => VerticalAlignmentRunValues.Subscript,
        XLFontVerticalTextAlignmentValues.Superscript => VerticalAlignmentRunValues.Superscript,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static PatternValues ToOpenXml(this XLFillPatternValues value) => value switch
    {
        XLFillPatternValues.DarkDown => PatternValues.DarkDown,
        XLFillPatternValues.DarkGray => PatternValues.DarkGray,
        XLFillPatternValues.DarkGrid => PatternValues.DarkGrid,
        XLFillPatternValues.DarkHorizontal => PatternValues.DarkHorizontal,
        XLFillPatternValues.DarkTrellis => PatternValues.DarkTrellis,
        XLFillPatternValues.DarkUp => PatternValues.DarkUp,
        XLFillPatternValues.DarkVertical => PatternValues.DarkVertical,
        XLFillPatternValues.Gray0625 => PatternValues.Gray0625,
        XLFillPatternValues.Gray125 => PatternValues.Gray125,
        XLFillPatternValues.LightDown => PatternValues.LightDown,
        XLFillPatternValues.LightGray => PatternValues.LightGray,
        XLFillPatternValues.LightGrid => PatternValues.LightGrid,
        XLFillPatternValues.LightHorizontal => PatternValues.LightHorizontal,
        XLFillPatternValues.LightTrellis => PatternValues.LightTrellis,
        XLFillPatternValues.LightUp => PatternValues.LightUp,
        XLFillPatternValues.LightVertical => PatternValues.LightVertical,
        XLFillPatternValues.MediumGray => PatternValues.MediumGray,
        XLFillPatternValues.None => PatternValues.None,
        XLFillPatternValues.Solid => PatternValues.Solid,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static BorderStyleValues ToOpenXml(this XLBorderStyleValues value) => value switch
    {
        XLBorderStyleValues.DashDot => BorderStyleValues.DashDot,
        XLBorderStyleValues.DashDotDot => BorderStyleValues.DashDotDot,
        XLBorderStyleValues.Dashed => BorderStyleValues.Dashed,
        XLBorderStyleValues.Dotted => BorderStyleValues.Dotted,
        XLBorderStyleValues.Double => BorderStyleValues.Double,
        XLBorderStyleValues.Hair => BorderStyleValues.Hair,
        XLBorderStyleValues.Medium => BorderStyleValues.Medium,
        XLBorderStyleValues.MediumDashDot => BorderStyleValues.MediumDashDot,
        XLBorderStyleValues.MediumDashDotDot => BorderStyleValues.MediumDashDotDot,
        XLBorderStyleValues.MediumDashed => BorderStyleValues.MediumDashed,
        XLBorderStyleValues.None => BorderStyleValues.None,
        XLBorderStyleValues.SlantDashDot => BorderStyleValues.SlantDashDot,
        XLBorderStyleValues.Thick => BorderStyleValues.Thick,
        XLBorderStyleValues.Thin => BorderStyleValues.Thin,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static HorizontalAlignmentValues ToOpenXml(this XLAlignmentHorizontalValues value) => value switch
    {
        XLAlignmentHorizontalValues.Center => HorizontalAlignmentValues.Center,
        XLAlignmentHorizontalValues.CenterContinuous => HorizontalAlignmentValues.CenterContinuous,
        XLAlignmentHorizontalValues.Distributed => HorizontalAlignmentValues.Distributed,
        XLAlignmentHorizontalValues.Fill => HorizontalAlignmentValues.Fill,
        XLAlignmentHorizontalValues.General => HorizontalAlignmentValues.General,
        XLAlignmentHorizontalValues.Justify => HorizontalAlignmentValues.Justify,
        XLAlignmentHorizontalValues.Left => HorizontalAlignmentValues.Left,
        XLAlignmentHorizontalValues.Right => HorizontalAlignmentValues.Right,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static VerticalAlignmentValues ToOpenXml(this XLAlignmentVerticalValues value) => value switch
    {
        XLAlignmentVerticalValues.Bottom => VerticalAlignmentValues.Bottom,
        XLAlignmentVerticalValues.Center => VerticalAlignmentValues.Center,
        XLAlignmentVerticalValues.Distributed => VerticalAlignmentValues.Distributed,
        XLAlignmentVerticalValues.Justify => VerticalAlignmentValues.Justify,
        XLAlignmentVerticalValues.Top => VerticalAlignmentValues.Top,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static PageOrderValues ToOpenXml(this XLPageOrderValues value) => value switch
    {
        XLPageOrderValues.DownThenOver => PageOrderValues.DownThenOver,
        XLPageOrderValues.OverThenDown => PageOrderValues.OverThenDown,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static CellCommentsValues ToOpenXml(this XLShowCommentsValues value) => value switch
    {
        XLShowCommentsValues.AsDisplayed => CellCommentsValues.AsDisplayed,
        XLShowCommentsValues.AtEnd => CellCommentsValues.AtEnd,
        XLShowCommentsValues.None => CellCommentsValues.None,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static PrintErrorValues ToOpenXml(this XLPrintErrorValues value) => value switch
    {
        XLPrintErrorValues.Blank => PrintErrorValues.Blank,
        XLPrintErrorValues.Dash => PrintErrorValues.Dash,
        XLPrintErrorValues.Displayed => PrintErrorValues.Displayed,
        XLPrintErrorValues.NA => PrintErrorValues.NA,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static CalculateModeValues ToOpenXml(this XLCalculateMode value) => value switch
    {
        XLCalculateMode.Auto => CalculateModeValues.Auto,
        XLCalculateMode.AutoNoTable => CalculateModeValues.AutoNoTable,
        XLCalculateMode.Manual => CalculateModeValues.Manual,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static ReferenceModeValues ToOpenXml(this XLReferenceStyle value) => value switch
    {
        XLReferenceStyle.R1C1 => ReferenceModeValues.R1C1,
        XLReferenceStyle.A1 => ReferenceModeValues.A1,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static uint ToOpenXml(this XLAlignmentReadingOrderValues value) => value switch
    {
        XLAlignmentReadingOrderValues.ContextDependent => 0,
        XLAlignmentReadingOrderValues.LeftToRight => 1,
        XLAlignmentReadingOrderValues.RightToLeft => 2,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static TotalsRowFunctionValues ToOpenXml(this XLTotalsRowFunction value) => value switch
    {
        XLTotalsRowFunction.None => TotalsRowFunctionValues.None,
        XLTotalsRowFunction.Sum => TotalsRowFunctionValues.Sum,
        XLTotalsRowFunction.Minimum => TotalsRowFunctionValues.Minimum,
        XLTotalsRowFunction.Maximum => TotalsRowFunctionValues.Maximum,
        XLTotalsRowFunction.Average => TotalsRowFunctionValues.Average,
        XLTotalsRowFunction.Count => TotalsRowFunctionValues.Count,
        XLTotalsRowFunction.CountNumbers => TotalsRowFunctionValues.CountNumbers,
        XLTotalsRowFunction.StandardDeviation => TotalsRowFunctionValues.StandardDeviation,
        XLTotalsRowFunction.Variance => TotalsRowFunctionValues.Variance,
        XLTotalsRowFunction.Custom => TotalsRowFunctionValues.Custom,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static DataValidationValues ToOpenXml(this XLAllowedValues value) => value switch
    {
        XLAllowedValues.AnyValue => DataValidationValues.None,
        XLAllowedValues.Custom => DataValidationValues.Custom,
        XLAllowedValues.Date => DataValidationValues.Date,
        XLAllowedValues.Decimal => DataValidationValues.Decimal,
        XLAllowedValues.List => DataValidationValues.List,
        XLAllowedValues.TextLength => DataValidationValues.TextLength,
        XLAllowedValues.Time => DataValidationValues.Time,
        XLAllowedValues.WholeNumber => DataValidationValues.Whole,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static DataValidationErrorStyleValues ToOpenXml(this XLErrorStyle value) => value switch
    {
        XLErrorStyle.Information => DataValidationErrorStyleValues.Information,
        XLErrorStyle.Warning => DataValidationErrorStyleValues.Warning,
        XLErrorStyle.Stop => DataValidationErrorStyleValues.Stop,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static DataValidationOperatorValues ToOpenXml(this XLOperator value) => value switch
    {
        XLOperator.Between => DataValidationOperatorValues.Between,
        XLOperator.EqualOrGreaterThan => DataValidationOperatorValues.GreaterThanOrEqual,
        XLOperator.EqualOrLessThan => DataValidationOperatorValues.LessThanOrEqual,
        XLOperator.EqualTo => DataValidationOperatorValues.Equal,
        XLOperator.GreaterThan => DataValidationOperatorValues.GreaterThan,
        XLOperator.LessThan => DataValidationOperatorValues.LessThan,
        XLOperator.NotBetween => DataValidationOperatorValues.NotBetween,
        XLOperator.NotEqualTo => DataValidationOperatorValues.NotEqual,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static SheetStateValues ToOpenXml(this XLWorksheetVisibility value) => value switch
    {
        XLWorksheetVisibility.Visible => SheetStateValues.Visible,
        XLWorksheetVisibility.Hidden => SheetStateValues.Hidden,
        XLWorksheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static PhoneticValues ToOpenXml(this XLPhoneticType value) => value switch
    {
        XLPhoneticType.FullWidthKatakana => PhoneticValues.FullWidthKatakana,
        XLPhoneticType.HalfWidthKatakana => PhoneticValues.HalfWidthKatakana,
        XLPhoneticType.Hiragana => PhoneticValues.Hiragana,
        XLPhoneticType.NoConversion => PhoneticValues.NoConversion,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static DataConsolidateFunctionValues ToOpenXml(this XLPivotSummary value) => value switch
    {
        XLPivotSummary.Sum => DataConsolidateFunctionValues.Sum,
        XLPivotSummary.Count => DataConsolidateFunctionValues.Count,
        XLPivotSummary.Average => DataConsolidateFunctionValues.Average,
        XLPivotSummary.Minimum => DataConsolidateFunctionValues.Minimum,
        XLPivotSummary.Maximum => DataConsolidateFunctionValues.Maximum,
        XLPivotSummary.Product => DataConsolidateFunctionValues.Product,
        XLPivotSummary.CountNumbers => DataConsolidateFunctionValues.CountNumbers,
        XLPivotSummary.StandardDeviation => DataConsolidateFunctionValues.StandardDeviation,
        XLPivotSummary.PopulationStandardDeviation => DataConsolidateFunctionValues.StandardDeviationP,
        XLPivotSummary.Variance => DataConsolidateFunctionValues.Variance,
        XLPivotSummary.PopulationVariance => DataConsolidateFunctionValues.VarianceP,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static ShowDataAsValues ToOpenXml(this XLPivotCalculation value) => value switch
    {
        XLPivotCalculation.Normal => ShowDataAsValues.Normal,
        XLPivotCalculation.DifferenceFrom => ShowDataAsValues.Difference,
        XLPivotCalculation.PercentageOf => ShowDataAsValues.Percent,
        XLPivotCalculation.PercentageDifferenceFrom => ShowDataAsValues.PercentageDifference,
        XLPivotCalculation.RunningTotal => ShowDataAsValues.RunTotal,
        XLPivotCalculation.PercentageOfRow => ShowDataAsValues.PercentOfRaw, // There's a typo in the OpenXML SDK =)
        XLPivotCalculation.PercentageOfColumn => ShowDataAsValues.PercentOfColumn,
        XLPivotCalculation.PercentageOfTotal => ShowDataAsValues.PercentOfTotal,
        XLPivotCalculation.Index => ShowDataAsValues.Index,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static FilterOperatorValues ToOpenXml(this XLFilterOperator value) => value switch
    {
        XLFilterOperator.Equal => FilterOperatorValues.Equal,
        XLFilterOperator.NotEqual => FilterOperatorValues.NotEqual,
        XLFilterOperator.GreaterThan => FilterOperatorValues.GreaterThan,
        XLFilterOperator.EqualOrGreaterThan => FilterOperatorValues.GreaterThanOrEqual,
        XLFilterOperator.LessThan => FilterOperatorValues.LessThan,
        XLFilterOperator.EqualOrLessThan => FilterOperatorValues.LessThanOrEqual,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static DynamicFilterValues ToOpenXml(this XLFilterDynamicType value) => value switch
    {
        XLFilterDynamicType.AboveAverage => DynamicFilterValues.AboveAverage,
        XLFilterDynamicType.BelowAverage => DynamicFilterValues.BelowAverage,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static DateTimeGroupingValues ToOpenXml(this XLDateTimeGrouping value) => value switch
    {
        XLDateTimeGrouping.Year => DateTimeGroupingValues.Year,
        XLDateTimeGrouping.Month => DateTimeGroupingValues.Month,
        XLDateTimeGrouping.Day => DateTimeGroupingValues.Day,
        XLDateTimeGrouping.Hour => DateTimeGroupingValues.Hour,
        XLDateTimeGrouping.Minute => DateTimeGroupingValues.Minute,
        XLDateTimeGrouping.Second => DateTimeGroupingValues.Second,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static SheetViewValues ToOpenXml(this XLSheetViewOptions value) => value switch
    {
        XLSheetViewOptions.Normal => SheetViewValues.Normal,
        XLSheetViewOptions.PageBreakPreview => SheetViewValues.PageBreakPreview,
        XLSheetViewOptions.PageLayout => SheetViewValues.PageLayout,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static Vml.StrokeLineStyleValues ToOpenXml(this XLLineStyle value) => value switch
    {
        XLLineStyle.Single => Vml.StrokeLineStyleValues.Single,
        XLLineStyle.ThickBetweenThin => Vml.StrokeLineStyleValues.ThickBetweenThin,
        XLLineStyle.ThickThin => Vml.StrokeLineStyleValues.ThickThin,
        XLLineStyle.ThinThick => Vml.StrokeLineStyleValues.ThinThick,
        XLLineStyle.ThinThin => Vml.StrokeLineStyleValues.ThinThin,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static ConditionalFormatValues ToOpenXml(this XLConditionalFormatType value) => value switch
    {
        XLConditionalFormatType.Expression => ConditionalFormatValues.Expression,
        XLConditionalFormatType.CellIs => ConditionalFormatValues.CellIs,
        XLConditionalFormatType.ColorScale => ConditionalFormatValues.ColorScale,
        XLConditionalFormatType.DataBar => ConditionalFormatValues.DataBar,
        XLConditionalFormatType.IconSet => ConditionalFormatValues.IconSet,
        XLConditionalFormatType.Top10 => ConditionalFormatValues.Top10,
        XLConditionalFormatType.IsUnique => ConditionalFormatValues.UniqueValues,
        XLConditionalFormatType.IsDuplicate => ConditionalFormatValues.DuplicateValues,
        XLConditionalFormatType.ContainsText => ConditionalFormatValues.ContainsText,
        XLConditionalFormatType.NotContainsText => ConditionalFormatValues.NotContainsText,
        XLConditionalFormatType.StartsWith => ConditionalFormatValues.BeginsWith,
        XLConditionalFormatType.EndsWith => ConditionalFormatValues.EndsWith,
        XLConditionalFormatType.IsBlank => ConditionalFormatValues.ContainsBlanks,
        XLConditionalFormatType.NotBlank => ConditionalFormatValues.NotContainsBlanks,
        XLConditionalFormatType.IsError => ConditionalFormatValues.ContainsErrors,
        XLConditionalFormatType.NotError => ConditionalFormatValues.NotContainsErrors,
        XLConditionalFormatType.TimePeriod => ConditionalFormatValues.TimePeriod,
        XLConditionalFormatType.AboveAverage => ConditionalFormatValues.AboveAverage,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static ConditionalFormatValueObjectValues ToOpenXml(this XLCFContentType value) => value switch
    {
        XLCFContentType.Number => ConditionalFormatValueObjectValues.Number,
        XLCFContentType.Percent => ConditionalFormatValueObjectValues.Percent,
        XLCFContentType.Maximum => ConditionalFormatValueObjectValues.Max,
        XLCFContentType.Minimum => ConditionalFormatValueObjectValues.Min,
        XLCFContentType.Formula => ConditionalFormatValueObjectValues.Formula,
        XLCFContentType.Percentile => ConditionalFormatValueObjectValues.Percentile,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static ConditionalFormattingOperatorValues ToOpenXml(this XLCFOperator value) => value switch
    {
        XLCFOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
        XLCFOperator.EqualOrLessThan => ConditionalFormattingOperatorValues.LessThanOrEqual,
        XLCFOperator.Equal => ConditionalFormattingOperatorValues.Equal,
        XLCFOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
        XLCFOperator.EqualOrGreaterThan => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
        XLCFOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
        XLCFOperator.Between => ConditionalFormattingOperatorValues.Between,
        XLCFOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
        XLCFOperator.Contains => ConditionalFormattingOperatorValues.ContainsText,
        XLCFOperator.NotContains => ConditionalFormattingOperatorValues.NotContains,
        XLCFOperator.StartsWith => ConditionalFormattingOperatorValues.BeginsWith,
        XLCFOperator.EndsWith => ConditionalFormattingOperatorValues.EndsWith,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static IconSetValues ToOpenXml(this XLIconSetStyle value) => value switch
    {
        XLIconSetStyle.ThreeArrows => IconSetValues.ThreeArrows,
        XLIconSetStyle.ThreeArrowsGray => IconSetValues.ThreeArrowsGray,
        XLIconSetStyle.ThreeFlags => IconSetValues.ThreeFlags,
        XLIconSetStyle.ThreeTrafficLights1 => IconSetValues.ThreeTrafficLights1,
        XLIconSetStyle.ThreeTrafficLights2 => IconSetValues.ThreeTrafficLights2,
        XLIconSetStyle.ThreeSigns => IconSetValues.ThreeSigns,
        XLIconSetStyle.ThreeSymbols => IconSetValues.ThreeSymbols,
        XLIconSetStyle.ThreeSymbols2 => IconSetValues.ThreeSymbols2,
        XLIconSetStyle.FourArrows => IconSetValues.FourArrows,
        XLIconSetStyle.FourArrowsGray => IconSetValues.FourArrowsGray,
        XLIconSetStyle.FourRedToBlack => IconSetValues.FourRedToBlack,
        XLIconSetStyle.FourRating => IconSetValues.FourRating,
        XLIconSetStyle.FourTrafficLights => IconSetValues.FourTrafficLights,
        XLIconSetStyle.FiveArrows => IconSetValues.FiveArrows,
        XLIconSetStyle.FiveArrowsGray => IconSetValues.FiveArrowsGray,
        XLIconSetStyle.FiveRating => IconSetValues.FiveRating,
        XLIconSetStyle.FiveQuarters => IconSetValues.FiveQuarters,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static TimePeriodValues ToOpenXml(this XLTimePeriod value) => value switch
    {
        XLTimePeriod.Yesterday => TimePeriodValues.Yesterday,
        XLTimePeriod.Today => TimePeriodValues.Today,
        XLTimePeriod.Tomorrow => TimePeriodValues.Tomorrow,
        XLTimePeriod.InTheLast7Days => TimePeriodValues.Last7Days,
        XLTimePeriod.LastWeek => TimePeriodValues.LastWeek,
        XLTimePeriod.ThisWeek => TimePeriodValues.ThisWeek,
        XLTimePeriod.NextWeek => TimePeriodValues.NextWeek,
        XLTimePeriod.LastMonth => TimePeriodValues.LastMonth,
        XLTimePeriod.ThisMonth => TimePeriodValues.ThisMonth,
        XLTimePeriod.NextMonth => TimePeriodValues.NextMonth,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static PartTypeInfo ToOpenXml(this XLPictureFormat value)
    {
        return PictureFormatMap[value];
    }

    public static Xdr.EditAsValues ToOpenXml(this XLPicturePlacement value) => value switch
    {
        XLPicturePlacement.FreeFloating => Xdr.EditAsValues.Absolute,
        XLPicturePlacement.Move => Xdr.EditAsValues.OneCell,
        XLPicturePlacement.MoveAndSize => Xdr.EditAsValues.TwoCell,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static PivotAreaValues ToOpenXml(this XLPivotAreaType value) => value switch
    {
        XLPivotAreaType.None => PivotAreaValues.None,
        XLPivotAreaType.Normal => PivotAreaValues.Normal,
        XLPivotAreaType.Data => PivotAreaValues.Data,
        XLPivotAreaType.All => PivotAreaValues.All,
        XLPivotAreaType.Origin => PivotAreaValues.Origin,
        XLPivotAreaType.Button => PivotAreaValues.Button,
        XLPivotAreaType.TopRight => PivotAreaValues.TopRight,
        XLPivotAreaType.TopEnd => PivotAreaValues.TopEnd,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "XLPivotAreaValues value not implemented"),
    };

    public static X14.SparklineTypeValues ToOpenXml(this XLSparklineType value) => value switch
    {
        XLSparklineType.Line => X14.SparklineTypeValues.Line,
        XLSparklineType.Column => X14.SparklineTypeValues.Column,
        XLSparklineType.Stacked => X14.SparklineTypeValues.Stacked,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static X14.SparklineAxisMinMaxValues ToOpenXml(this XLSparklineAxisMinMax value) => value switch
    {
        XLSparklineAxisMinMax.Automatic => X14.SparklineAxisMinMaxValues.Individual,
        XLSparklineAxisMinMax.SameForAll => X14.SparklineAxisMinMaxValues.Group,
        XLSparklineAxisMinMax.Custom => X14.SparklineAxisMinMaxValues.Custom,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static X14.DisplayBlanksAsValues ToOpenXml(this XLDisplayBlanksAsValues value) => value switch
    {
        XLDisplayBlanksAsValues.Interpolate => X14.DisplayBlanksAsValues.Span,
        XLDisplayBlanksAsValues.NotPlotted => X14.DisplayBlanksAsValues.Gap,
        XLDisplayBlanksAsValues.Zero => X14.DisplayBlanksAsValues.Zero,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static FieldSortValues ToOpenXml(this XLPivotSortType value) => value switch
    {
        XLPivotSortType.Default => FieldSortValues.Manual,
        XLPivotSortType.Ascending => FieldSortValues.Ascending,
        XLPivotSortType.Descending => FieldSortValues.Descending,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static X14.DataBarAxisPositionValues ToOpenXml(this XLDataBarAxisPosition value)
    {
        if (value == XLDataBarAxisPosition.Automatic) return X14.DataBarAxisPositionValues.Automatic;
        if (value == XLDataBarAxisPosition.Middle) return X14.DataBarAxisPositionValues.Middle;
        if (value == XLDataBarAxisPosition.None) return X14.DataBarAxisPositionValues.None;
        throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
    }

    public static XLFontUnderlineValues ToXLibur(this UnderlineValues value)
    {
        return UnderlineValuesMap[value];
    }

    public static XLFontScheme ToXLibur(this FontSchemeValues value)
    {
        return FontSchemeMap[value];
    }

    public static XLPageOrientation ToXLibur(this OrientationValues value)
    {
        return OrientationMap[value];
    }

    public static XLFontVerticalTextAlignmentValues ToXLibur(this VerticalAlignmentRunValues value)
    {
        return VerticalAlignmentRunMap[value];
    }

    public static XLFillPatternValues ToXLibur(this PatternValues value)
    {
        return PatternMap[value];
    }

    public static XLBorderStyleValues ToXLibur(this BorderStyleValues value)
    {
        return BorderStyleMap[value];
    }

    public static XLAlignmentHorizontalValues ToXLibur(this HorizontalAlignmentValues value)
    {
        return HorizontalAlignmentMap[value];
    }

    public static XLAlignmentVerticalValues ToXLibur(this VerticalAlignmentValues value)
    {
        return VerticalAlignmentMap[value];
    }

    public static XLPageOrderValues ToXLibur(this PageOrderValues value)
    {
        return PageOrdersMap[value];
    }

    public static XLShowCommentsValues ToXLibur(this CellCommentsValues value)
    {
        return CellCommentsMap[value];
    }

    public static XLPrintErrorValues ToXLibur(this PrintErrorValues value)
    {
        return PrintErrorMap[value];
    }

    public static XLCalculateMode ToXLibur(this CalculateModeValues value)
    {
        return CalculateModeMap[value];
    }

    public static XLReferenceStyle ToXLibur(this ReferenceModeValues value)
    {
        return ReferenceModeMap[value];
    }

    public static XLAlignmentReadingOrderValues ToXLibur(this uint value) => value switch
    {
        0 => XLAlignmentReadingOrderValues.ContextDependent,
        1 => XLAlignmentReadingOrderValues.LeftToRight,
        2 => XLAlignmentReadingOrderValues.RightToLeft,
        _ => throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!"),
    };

    public static XLTotalsRowFunction ToXLibur(this TotalsRowFunctionValues value)
    {
        return TotalsRowFunctionMap[value];
    }

    public static XLAllowedValues ToXLibur(this DataValidationValues value)
    {
        return DataValidationMap[value];
    }

    public static XLErrorStyle ToXLibur(this DataValidationErrorStyleValues value)
    {
        return DataValidationErrorStyleMap[value];
    }

    public static XLOperator ToXLibur(this DataValidationOperatorValues value)
    {
        return DataValidationOperatorMap[value];
    }

    public static XLWorksheetVisibility ToXLibur(this SheetStateValues value)
    {
        return SheetStateMap[value];
    }

    public static XLPhoneticAlignment ToXLibur(this PhoneticAlignmentValues value)
    {
        return PhoneticAlignmentMap[value];
    }

    public static XLPhoneticType ToXLibur(this PhoneticValues value)
    {
        return PhoneticMap[value];
    }

    public static XLPivotSummary ToXLibur(this DataConsolidateFunctionValues value)
    {
        return DataConsolidateFunctionMap[value];
    }

    public static XLPivotCalculation ToXLibur(this ShowDataAsValues value)
    {
        return ShowDataAsMap[value];
    }

    public static XLFilterOperator ToXLibur(this FilterOperatorValues value)
    {
        return FilterOperatorMap[value];
    }

    public static XLFilterDynamicType ToXLibur(this DynamicFilterValues value)
    {
        return DynamicFilterMap[value];
    }

    public static XLDateTimeGrouping ToXLibur(this DateTimeGroupingValues value)
    {
        return DateTimeGroupingMap[value];
    }

    public static XLSheetViewOptions ToXLibur(this SheetViewValues value)
    {
        return SheetViewMap[value];
    }

    public static XLLineStyle ToXLibur(this Vml.StrokeLineStyleValues value)
    {
        return StrokeLineStyleMap[value];
    }

    public static XLConditionalFormatType ToXLibur(this ConditionalFormatValues value)
    {
        return ConditionalFormatMap[value];
    }

    public static XLCFContentType ToXLibur(this ConditionalFormatValueObjectValues value)
    {
        return ConditionalFormatValueObjectMap[value];
    }

    public static XLCFOperator ToXLibur(this ConditionalFormattingOperatorValues value)
    {
        return ConditionalFormattingOperatorMap[value];
    }

    public static XLIconSetStyle ToXLibur(this IconSetValues value)
    {
        return IconSetMap[value];
    }

    public static XLTimePeriod ToXLibur(this TimePeriodValues value)
    {
        return TimePeriodMap[value];
    }

    public static XLPivotAreaType ToXLibur(this PivotAreaValues value)
    {
        return PivotAreaMap[value];
    }

    public static XLSparklineType ToXLibur(this X14.SparklineTypeValues value)
    {
        return SparklineTypeMap[value];
    }

    public static XLSparklineAxisMinMax ToXLibur(this X14.SparklineAxisMinMaxValues value)
    {
        return SparklineAxisMinMaxMap[value];
    }

    public static XLDisplayBlanksAsValues ToXLibur(this X14.DisplayBlanksAsValues value)
    {
        return DisplayBlanksAsMap[value];
    }

    public static XLPivotSortType ToXLibur(this FieldSortValues value)
    {
        return FieldSortMap[value];
    }

    internal static XLPivotAxis ToXLibur(this PivotTableAxisValues value)
    {
        return PivotTableAxisMap[value];
    }

    internal static XLPivotItemType ToXLibur(this ItemValues value)
    {
        return ItemMap[value];
    }

    internal static XLPivotFormatAction ToXLibur(this FormatActionValues value)
    {
        return FormatActionMap[value];
    }

    internal static XLPivotCfScope ToXLibur(this ScopeValues value)
    {
        return ScopeMap[value];
    }

    internal static XLPivotCfRuleType ToXLibur(this RuleValues value)
    {
        return RuleMap[value];
    }

    public static XLDataBarAxisPosition ToXLibur(this X14.DataBarAxisPositionValues value)
    {
        if (value == X14.DataBarAxisPositionValues.Automatic) return XLDataBarAxisPosition.Automatic;
        if (value == X14.DataBarAxisPositionValues.Middle) return XLDataBarAxisPosition.Middle;
        if (value == X14.DataBarAxisPositionValues.None) return XLDataBarAxisPosition.None;
        throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
    }

}