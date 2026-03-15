using XLibur.Excel.AutoFilters;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads worksheet-level elements: sheet views, page setup, protection, data validation, autofilter, hyperlinks, and breaks.
/// </summary>
internal static class WorksheetElementReader
{
    internal static void LoadSheetViews(SheetViews sheetViews, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(ws);

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();

        if (sheetView == null) return;

        LoadSheetViewProperties(sheetView, ws);
        LoadSheetViewSelection(sheetView, ws);
        LoadSheetViewZoom(sheetView, ws);
        LoadSheetViewPane(sheetView, ws);

        if (sheetView.TopLeftCell?.Value is { } topLeftCell && XLHelper.IsValidA1Address(topLeftCell))
            ws.SheetView.TopLeftCellAddress = ws.Cell(topLeftCell)!.Address;
    }

    private static void LoadSheetViewProperties(SheetView sheetView, XLWorksheet ws)
    {
        if (sheetView.RightToLeft != null) ws.RightToLeft = sheetView.RightToLeft.Value;
        if (sheetView.ShowFormulas != null) ws.ShowFormulas = sheetView.ShowFormulas.Value;
        if (sheetView.ShowGridLines != null) ws.ShowGridLines = sheetView.ShowGridLines.Value;
        if (sheetView.ShowOutlineSymbols != null)
            ws.ShowOutlineSymbols = sheetView.ShowOutlineSymbols.Value;
        if (sheetView.ShowRowColHeaders != null) ws.ShowRowColHeaders = sheetView.ShowRowColHeaders.Value;
        if (sheetView.ShowRuler != null) ws.ShowRuler = sheetView.ShowRuler.Value;
        if (sheetView.ShowWhiteSpace != null) ws.ShowWhiteSpace = sheetView.ShowWhiteSpace.Value;
        if (sheetView.ShowZeros != null) ws.ShowZeros = sheetView.ShowZeros.Value;
        if (sheetView.TabSelected != null) ws.TabSelected = sheetView.TabSelected.Value;
    }

    private static void LoadSheetViewSelection(SheetView sheetView, XLWorksheet ws)
    {
        var selection = sheetView.Elements<Selection>().FirstOrDefault();
        if (selection == null) return;

        if (selection.SequenceOfReferences != null)
            ws.Ranges(selection.SequenceOfReferences!.InnerText!.Replace(" ", ",")).Select();

        if (selection.ActiveCell != null)
            ws.Cell(selection.ActiveCell!.Value!)!.SetActive();
    }

    private static void LoadSheetViewZoom(SheetView sheetView, XLWorksheet ws)
    {
        if (sheetView.ZoomScale != null)
            ws.SheetView.ZoomScale = (int)UInt32Value.ToUInt32(sheetView.ZoomScale);
        if (sheetView.ZoomScaleNormal != null)
            ws.SheetView.ZoomScaleNormal = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleNormal);
        if (sheetView.ZoomScalePageLayoutView != null)
            ws.SheetView.ZoomScalePageLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScalePageLayoutView);
        if (sheetView.ZoomScaleSheetLayoutView != null)
            ws.SheetView.ZoomScaleSheetLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleSheetLayoutView);
    }

    private static void LoadSheetViewPane(SheetView sheetView, XLWorksheet ws)
    {
        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (new[] { PaneStateValues.Frozen, PaneStateValues.FrozenSplit }.Contains(pane?.State?.Value ??
                PaneStateValues.Split))
        {
            if (pane!.HorizontalSplit != null)
                ws.SheetView.SplitColumn = (int)pane.HorizontalSplit.Value;
            if (pane!.VerticalSplit != null)
                ws.SheetView.SplitRow = (int)pane.VerticalSplit.Value;
        }
    }

    internal static void LoadPrintOptions(PrintOptions printOptions, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(printOptions);

        if (printOptions.GridLines != null)
            ws.PageSetup.ShowGridlines = printOptions.GridLines;
        if (printOptions.HorizontalCentered != null)
            ws.PageSetup.CenterHorizontally = printOptions.HorizontalCentered;
        if (printOptions.VerticalCentered != null)
            ws.PageSetup.CenterVertically = printOptions.VerticalCentered;
        if (printOptions.Headings != null)
            ws.PageSetup.ShowRowAndColumnHeadings = printOptions.Headings;
    }

    internal static void LoadPageMargins(PageMargins pageMargins, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(pageMargins);

        if (pageMargins.Bottom != null)
            ws.PageSetup.Margins.Bottom = pageMargins.Bottom;
        if (pageMargins.Footer != null)
            ws.PageSetup.Margins.Footer = pageMargins.Footer;
        if (pageMargins.Header != null)
            ws.PageSetup.Margins.Header = pageMargins.Header;
        if (pageMargins.Left != null)
            ws.PageSetup.Margins.Left = pageMargins.Left;
        if (pageMargins.Right != null)
            ws.PageSetup.Margins.Right = pageMargins.Right;
        if (pageMargins.Top != null)
            ws.PageSetup.Margins.Top = pageMargins.Top;
    }

    internal static void LoadPageSetup(PageSetup pageSetup, XLWorksheet ws, PageSetupProperties? pageSetupProperties)
    {
        ArgumentNullException.ThrowIfNull(pageSetup);

        if (pageSetup.PaperSize != null)
            ws.PageSetup.PaperSize = (XLPaperSize)int.Parse(pageSetup.PaperSize.InnerText!);
        if (pageSetup.Scale != null)
            ws.PageSetup.Scale = int.Parse(pageSetup.Scale.InnerText!);

        LoadFitToPage(pageSetup, ws, pageSetupProperties);
        LoadPageSetupOptions(pageSetup, ws);
    }

    private static void LoadFitToPage(PageSetup pageSetup, XLWorksheet ws, PageSetupProperties? pageSetupProperties)
    {
        if (pageSetupProperties?.FitToPage != null && pageSetupProperties.FitToPage.Value)
        {
            ws.PageSetup.PagesWide = pageSetup.FitToWidth == null ? 1 : int.Parse(pageSetup.FitToWidth.InnerText!);
            ws.PageSetup.PagesTall = pageSetup.FitToHeight == null ? 1 : int.Parse(pageSetup.FitToHeight.InnerText!);
        }
    }

    private static void LoadPageSetupOptions(PageSetup pageSetup, XLWorksheet ws)
    {
        if (pageSetup.PageOrder != null)
            ws.PageSetup.PageOrder = pageSetup.PageOrder.Value.ToXLibur();
        if (pageSetup.Orientation != null)
            ws.PageSetup.PageOrientation = pageSetup.Orientation.Value.ToXLibur();
        if (pageSetup.BlackAndWhite != null)
            ws.PageSetup.BlackAndWhite = pageSetup.BlackAndWhite;
        if (pageSetup.Draft != null)
            ws.PageSetup.DraftQuality = pageSetup.Draft;
        if (pageSetup.CellComments != null)
            ws.PageSetup.ShowComments = pageSetup.CellComments.Value.ToXLibur();
        if (pageSetup.Errors != null)
            ws.PageSetup.PrintErrorValue = pageSetup.Errors.Value.ToXLibur();
        if (pageSetup.HorizontalDpi != null) ws.PageSetup.HorizontalDpi = (int)pageSetup.HorizontalDpi.Value;
        if (pageSetup.VerticalDpi != null) ws.PageSetup.VerticalDpi = (int)pageSetup.VerticalDpi.Value;
        if (pageSetup.FirstPageNumber?.HasValue ?? false)
            ws.PageSetup.FirstPageNumber = (int)pageSetup.FirstPageNumber.Value;
    }

    internal static void LoadHeaderFooter(HeaderFooter headerFooter, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(headerFooter);

        if (headerFooter.AlignWithMargins != null)
            ws.PageSetup.AlignHFWithMargins = headerFooter.AlignWithMargins;
        if (headerFooter.ScaleWithDoc != null)
            ws.PageSetup.ScaleHFWithDocument = headerFooter.ScaleWithDoc;

        if (headerFooter.DifferentFirst != null)
            ws.PageSetup.DifferentFirstPageOnHF = headerFooter.DifferentFirst;
        if (headerFooter.DifferentOddEven != null)
            ws.PageSetup.DifferentOddEvenPagesOnHF = headerFooter.DifferentOddEven;

        // Footers
        var xlFooter = (XLHeaderFooter)ws.PageSetup.Footer;
        var evenFooter = headerFooter.EvenFooter;
        if (evenFooter != null)
            xlFooter.SetInnerText(XLHFOccurrence.EvenPages, evenFooter.Text);
        var oddFooter = headerFooter.OddFooter;
        if (oddFooter != null)
            xlFooter.SetInnerText(XLHFOccurrence.OddPages, oddFooter.Text);
        var firstFooter = headerFooter.FirstFooter;
        if (firstFooter != null)
            xlFooter.SetInnerText(XLHFOccurrence.FirstPage, firstFooter.Text);
        // Headers
        var xlHeader = (XLHeaderFooter)ws.PageSetup.Header;
        var evenHeader = headerFooter.EvenHeader;
        if (evenHeader != null)
            xlHeader.SetInnerText(XLHFOccurrence.EvenPages, evenHeader.Text);
        var oddHeader = headerFooter.OddHeader;
        if (oddHeader != null)
            xlHeader.SetInnerText(XLHFOccurrence.OddPages, oddHeader.Text);
        var firstHeader = headerFooter.FirstHeader;
        if (firstHeader != null)
            xlHeader.SetInnerText(XLHFOccurrence.FirstPage, firstHeader.Text);

        ((XLHeaderFooter)ws.PageSetup.Header).SetAsInitial();
        ((XLHeaderFooter)ws.PageSetup.Footer).SetAsInitial();
    }

    internal static void LoadSheetProperties(SheetProperties sheetProperty, XLWorksheet ws,
        out PageSetupProperties? pageSetupProperties)
    {
        ArgumentNullException.ThrowIfNull(ws);
        pageSetupProperties = null;

        if (sheetProperty.TabColor != null)
            ws.TabColor = sheetProperty.TabColor.ToXLiburColor();

        if (sheetProperty.OutlineProperties != null)
        {
            if (sheetProperty.OutlineProperties.SummaryBelow != null)
            {
                ws.Outline.SummaryVLocation = sheetProperty.OutlineProperties.SummaryBelow
                    ? XLOutlineSummaryVLocation.Bottom
                    : XLOutlineSummaryVLocation.Top;
            }

            if (sheetProperty.OutlineProperties.SummaryRight != null)
            {
                ws.Outline.SummaryHLocation = sheetProperty.OutlineProperties.SummaryRight
                    ? XLOutlineSummaryHLocation.Right
                    : XLOutlineSummaryHLocation.Left;
            }
        }

        if (sheetProperty.PageSetupProperties != null)
            pageSetupProperties = sheetProperty.PageSetupProperties;
    }

    internal static void LoadRowBreaks(RowBreaks rowBreaks, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(rowBreaks);

        foreach (var rowBreak in rowBreaks.Elements<Break>())
            ws.PageSetup.RowBreaks.Add(int.Parse(rowBreak.Id!.InnerText!));
    }

    internal static void LoadColumnBreaks(ColumnBreaks columnBreaks, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(columnBreaks);
        foreach (var columnBreak in columnBreaks.Elements<Break>().Where(columnBreak => columnBreak.Id != null))
        {
            ws.PageSetup.ColumnBreaks.Add(int.Parse(columnBreak.Id!.InnerText!));
        }
    }

    internal static void LoadSheetProtection(SheetProtection sp, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(ws);

        ws.Protection.IsProtected = OpenXmlHelper.GetBooleanValueAsBool(sp.Sheet, false);

        var algorithmName = sp.AlgorithmName?.Value ?? string.Empty;
        if (string.IsNullOrEmpty(algorithmName))
        {
            ws.Protection.PasswordHash = sp.Password?.Value ?? string.Empty;
            ws.Protection.Base64EncodedSalt = string.Empty;
        }
        else if (DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName))
        {
            ws.Protection.Algorithm =
                DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
            ws.Protection.PasswordHash = sp.HashValue?.Value ?? string.Empty;
            ws.Protection.SpinCount = sp.SpinCount?.Value ?? 0;
            ws.Protection.Base64EncodedSalt = sp.SaltValue?.Value ?? string.Empty;
        }

        ws.Protection.AllowElement(XLSheetProtectionElements.FormatCells,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.FormatCells, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.FormatColumns,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.FormatColumns, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.FormatRows,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.FormatRows, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.InsertColumns,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.InsertColumns, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.InsertHyperlinks,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.InsertHyperlinks, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.InsertRows,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.InsertRows, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.DeleteColumns,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.DeleteColumns, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.DeleteRows,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.DeleteRows, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.AutoFilter,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.AutoFilter, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.PivotTables,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.PivotTables, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.Sort, !OpenXmlHelper.GetBooleanValueAsBool(sp.Sort, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.EditScenarios,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.Scenarios, true));

        ws.Protection.AllowElement(XLSheetProtectionElements.EditObjects,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.Objects, false));
        ws.Protection.AllowElement(XLSheetProtectionElements.SelectLockedCells,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.SelectLockedCells, false));
        ws.Protection.AllowElement(XLSheetProtectionElements.SelectUnlockedCells,
            !OpenXmlHelper.GetBooleanValueAsBool(sp.SelectUnlockedCells, false));
    }

    internal static void LoadDataValidations(DataValidations dataValidations, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(ws);

        foreach (var dvs in dataValidations.Elements<DataValidation>())
        {
            var txt = dvs.SequenceOfReferences!.InnerText;
            if (string.IsNullOrWhiteSpace(txt)) continue;
            foreach (var rangeAddress in txt.Split(' '))
            {
                var dvt = new XLDataValidation(ws.Range(rangeAddress)!);
                ws.DataValidations.Add(dvt, skipIntersectionsCheck: true);
                ApplyDataValidationProperties(dvs, dvt);
            }
        }
    }

    private static void ApplyDataValidationProperties(DataValidation dvs, XLDataValidation dvt)
    {
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
        if (dvs.Formula1 != null) dvt.MinValue = dvs.Formula1.Text;
        if (dvs.Formula2 != null) dvt.MaxValue = dvs.Formula2.Text;
    }

    internal static void LoadHyperlinks(Hyperlinks hyperlinks, WorksheetPart worksheetPart, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(ws);
        ArgumentNullException.ThrowIfNull(hyperlinks);
        ArgumentNullException.ThrowIfNull(worksheetPart);
        var hyperlinkDictionary = worksheetPart.HyperlinkRelationships.ToDictionary(hr => hr.Id, hr => hr.Uri);

        foreach (var hl in hyperlinks.Elements<Hyperlink>())
        {
            if (hl.Reference!.Value!.Equals("#REF")) continue;
            var tooltip = hl.Tooltip != null ? hl.Tooltip.Value : string.Empty;
            var xlRange = ws.Range(hl.Reference!.Value!);
            foreach (var xlCell1 in xlRange!.Cells())
            {
                var xlCell = (XLCell)xlCell1;
                if (hl.Id != null)
                    xlCell.SetCellHyperlink(new XLHyperlink(hyperlinkDictionary[hl.Id.Value!], tooltip!));
                else if (hl.Location != null)
                    xlCell.SetCellHyperlink(new XLHyperlink(hl.Location.Value!, tooltip!));
                else
                    xlCell.SetCellHyperlink(new XLHyperlink(hl.Reference!.Value!, tooltip!));
            }
        }
    }

    internal static void LoadAutoFilter(AutoFilter af, XLWorksheet ws,
        Dictionary<int, DifferentialFormat> differentialFormats)
    {
        if (af != null)
        {
            ws.Range(af.Reference!.Value!)!.SetAutoFilter();
            var autoFilter = ws.AutoFilter;
            LoadAutoFilterSort(af, ws, autoFilter);
            LoadAutoFilterColumns(af, autoFilter, differentialFormats);
        }
    }

    internal static void LoadAutoFilterColumns(AutoFilter af, XLAutoFilter autoFilter,
        Dictionary<int, DifferentialFormat>? differentialFormats = null)
    {
        foreach (var filterColumn in af.Elements<FilterColumn>())
        {
            var column = (int)filterColumn.ColumnId!.Value + 1;
            var xlFilterColumn = autoFilter.Column(column);
            if (filterColumn.CustomFilters is { } customFilters)
                LoadCustomFilters(xlFilterColumn, customFilters);
            else if (filterColumn.Filters is { } filters)
                LoadRegularFilters(xlFilterColumn, filters);
            else if (filterColumn.Top10 is { } top10)
                LoadTop10Filter(xlFilterColumn, top10);
            else if (filterColumn.DynamicFilter is { } dynamicFilter)
                LoadDynamicFilter(xlFilterColumn, filterColumn, dynamicFilter);
            else if (filterColumn.Elements<ColorFilter>().FirstOrDefault() is { } colorFilter
                     && differentialFormats is not null)
                LoadColorFilter(xlFilterColumn, colorFilter, differentialFormats);
        }
    }

    private static void LoadCustomFilters(XLFilterColumn xlFilterColumn, CustomFilters customFilters)
    {
        xlFilterColumn.FilterType = XLFilterType.Custom;
        var connector = OpenXmlHelper.GetBooleanValueAsBool(customFilters.And, false)
            ? XLConnector.And
            : XLConnector.Or;

        foreach (var filter in customFilters.OfType<CustomFilter>())
        {
            var op = filter.Operator is not null ? filter.Operator.Value.ToXLibur() : XLFilterOperator.Equal;
            var filterValue = filter.Val!.Value!;
            var xlFilter = CreateCustomXLFilter(op, filterValue, connector);
            xlFilterColumn.AddFilter(xlFilter);
        }
    }

    private static XLFilter CreateCustomXLFilter(XLFilterOperator op, string filterValue, XLConnector connector)
    {
        switch (op)
        {
            case XLFilterOperator.Equal:
                return XLFilter.CreateCustomPatternFilter(filterValue, true, connector);
            case XLFilterOperator.NotEqual:
                return XLFilter.CreateCustomPatternFilter(filterValue, false, connector);
            default:
                // OOXML allows only string, so do your best to convert back to a properly typed
                // variable. It's not perfect, but let's mimic Excel.
                var customValue = XLCellValue.FromText(filterValue, CultureInfo.InvariantCulture);
                return XLFilter.CreateCustomFilter(customValue, op, connector);
        }
    }

    private static void LoadRegularFilters(XLFilterColumn xlFilterColumn, Filters filters)
    {
        xlFilterColumn.FilterType = XLFilterType.Regular;
        foreach (var filter in filters.OfType<Filter>())
        {
            xlFilterColumn.AddFilter(XLFilter.CreateRegularFilter(filter.Val!.Value!));
        }

        foreach (var dateGroupItem in filters.OfType<DateGroupItem>())
        {
            LoadDateGroupItem(xlFilterColumn, dateGroupItem);
        }
    }

    private static void LoadDateGroupItem(XLFilterColumn xlFilterColumn, DateGroupItem dateGroupItem)
    {
        if (dateGroupItem.DateTimeGrouping is null || !dateGroupItem.DateTimeGrouping.HasValue)
            return;

        var xlGrouping = dateGroupItem.DateTimeGrouping.Value.ToXLibur();
        var year = 1900;
        var month = 1;
        var day = 1;
        var hour = 0;
        var minute = 0;
        var second = 0;

        var valid = true;

        if (xlGrouping >= XLDateTimeGrouping.Year)
            valid = TryGetDatePart(dateGroupItem.Year, ref year) && valid;

        if (xlGrouping >= XLDateTimeGrouping.Month)
            valid = TryGetDatePart(dateGroupItem.Month, ref month) && valid;

        if (xlGrouping >= XLDateTimeGrouping.Day)
            valid = TryGetDatePart(dateGroupItem.Day, ref day) && valid;

        if (xlGrouping >= XLDateTimeGrouping.Hour)
            valid = TryGetDatePart(dateGroupItem.Hour, ref hour) && valid;

        if (xlGrouping >= XLDateTimeGrouping.Minute)
            valid = TryGetDatePart(dateGroupItem.Minute, ref minute) && valid;

        if (xlGrouping >= XLDateTimeGrouping.Second)
            valid = TryGetDatePart(dateGroupItem.Second, ref second) && valid;

        if (valid)
        {
            var date = new DateTime(year, month, day, hour, minute, second);
            var xlDateGroupFilter = XLFilter.CreateDateGroupFilter(date, xlGrouping);
            xlFilterColumn.AddFilter(xlDateGroupFilter);
        }
    }

    private static bool TryGetDatePart(UInt16Value? value, ref int result)
    {
        if (value?.HasValue ?? false)
        {
            result = value.Value;
            return true;
        }

        return false;
    }

    private static void LoadTop10Filter(XLFilterColumn xlFilterColumn, Top10 top10)
    {
        xlFilterColumn.FilterType = XLFilterType.TopBottom;
        xlFilterColumn.TopBottomType = OpenXmlHelper.GetBooleanValueAsBool(top10.Percent, false)
            ? XLTopBottomType.Percent
            : XLTopBottomType.Items;
        var takeTop = OpenXmlHelper.GetBooleanValueAsBool(top10.Top, true);
        xlFilterColumn.TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

        // Value contains how many percent or items, so it can only be int.
        // Filter value is optional, so we don't rely on it.
        var percentsOrItems = (int)top10.Val!.Value;
        xlFilterColumn.TopBottomValue = percentsOrItems;
        xlFilterColumn.AddFilter(XLFilter.CreateTopBottom(takeTop, percentsOrItems));
    }

    private static void LoadDynamicFilter(XLFilterColumn xlFilterColumn, FilterColumn filterColumn,
        DynamicFilter dynamicFilter)
    {
        xlFilterColumn.FilterType = XLFilterType.Dynamic;
        var dynamicType = dynamicFilter.Type is { } dynamicFilterType
            ? dynamicFilterType.Value.ToXLibur()
            : XLFilterDynamicType.AboveAverage;
        var dynamicValue = filterColumn.DynamicFilter!.Val!.Value;

        xlFilterColumn.DynamicType = dynamicType;
        xlFilterColumn.DynamicValue = dynamicValue;
        xlFilterColumn.AddFilter(XLFilter.CreateAverage(dynamicValue,
            dynamicType == XLFilterDynamicType.AboveAverage));
    }

    private static void LoadColorFilter(XLFilterColumn xlFilterColumn, ColorFilter colorFilter,
        Dictionary<int, DifferentialFormat> differentialFormats)
    {
        xlFilterColumn.FilterType = XLFilterType.Color;
        var byCellColor = OpenXmlHelper.GetBooleanValueAsBool(colorFilter.CellColor, true);
        xlFilterColumn.FilterByCellColor = byCellColor;

        var formatId = (int)colorFilter.FormatId!.Value;
        if (differentialFormats.TryGetValue(formatId, out var dxf))
        {
            var filterColorValue = ResolveFilterColor(dxf, byCellColor);
            xlFilterColumn.FilterColor = filterColorValue;
            xlFilterColumn.AddFilter(XLFilter.CreateColorFilter(filterColorValue, byCellColor));
        }
    }

    private static XLColor ResolveFilterColor(DifferentialFormat dxf, bool byCellColor)
    {
        if (byCellColor)
        {
            var patternFill = dxf.Fill?.PatternFill;
            if (patternFill?.BackgroundColor is { } bgColor)
                return bgColor.ToXLiburColor();
            if (patternFill?.ForegroundColor is { } fgColor)
                return fgColor.ToXLiburColor();
            return XLColor.NoColor;
        }

        var fontColor = dxf.Font?.Color;
        return fontColor is not null ? fontColor.ToXLiburColor() : XLColor.NoColor;
    }

    internal static void LoadAutoFilterSort(AutoFilter af, XLWorksheet ws, XLAutoFilter autoFilter)
    {
        var sort = af.Elements<SortState>().FirstOrDefault();
        if (sort != null)
        {
            var condition = sort.Elements<SortCondition>().FirstOrDefault();
            if (condition != null)
            {
                var column = ws.Range(condition.Reference!.Value!)!.FirstCell().Address.ColumnNumber -
                    autoFilter.Range.FirstCell().Address.ColumnNumber + 1;
                autoFilter.SortColumn = column;
                autoFilter.Sorted = true;
                autoFilter.SortOrder = condition.Descending != null && condition.Descending.Value
                    ? XLSortOrder.Descending
                    : XLSortOrder.Ascending;
            }
        }
    }
}
