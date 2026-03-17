using XLibur.Excel.ContentManagers;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel.IO;

internal static class SheetViewWriter
{
    internal static void WriteSheetProperties(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        worksheet.SheetProperties ??= new SheetProperties();

        worksheet.SheetProperties.TabColor = xlWorksheet.TabColor.HasValue
            ? new TabColor().FromXLiburColor<TabColor>(xlWorksheet.TabColor)
            : null;

        cm.SetElement(XLWorksheetContents.SheetProperties, worksheet.SheetProperties);

        worksheet.SheetProperties.OutlineProperties ??= new OutlineProperties();

        worksheet.SheetProperties.OutlineProperties.SummaryBelow =
            (xlWorksheet.Outline.SummaryVLocation ==
             XLOutlineSummaryVLocation.Bottom);
        worksheet.SheetProperties.OutlineProperties.SummaryRight =
            (xlWorksheet.Outline.SummaryHLocation ==
             XLOutlineSummaryHLocation.Right);

        if (worksheet.SheetProperties.PageSetupProperties == null
            && (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0))
            worksheet.SheetProperties.PageSetupProperties = new PageSetupProperties { FitToPage = true };
    }

    internal static void WriteSheetDimension(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        // Empty worksheets have dimension A1 (not A1:A1)
        var sheetDimensionReference = "A1";
        if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
        {
            var maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
            var maxRow = xlWorksheet.Internals.CellsCollection.MaxRowUsed;
            sheetDimensionReference = "A1:" + XLHelper.GetColumnLetterFromNumber(maxColumn) +
                                      maxRow.ToInvariantString();
        }

        worksheet.SheetDimension ??= new SheetDimension { Reference = sheetDimensionReference };

        cm.SetElement(XLWorksheetContents.SheetDimension, worksheet.SheetDimension);
    }

    internal static void WriteSheetViews(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        worksheet.SheetViews ??= new SheetViews();

        cm.SetElement(XLWorksheetContents.SheetViews, worksheet.SheetViews);

        var sheetView = (SheetView?)worksheet.SheetViews.FirstOrDefault();
        if (sheetView == null)
        {
            sheetView = new SheetView { WorkbookViewId = 0U };
            worksheet.SheetViews.AppendChild(sheetView);
        }

        var svcm = new XLSheetViewContentManager(sheetView);

        SetBooleanViewProperties(sheetView, xlWorksheet);

        if (xlWorksheet.SheetView.View == XLSheetViewOptions.Normal)
            sheetView.View = null;
        else
            sheetView.View = xlWorksheet.SheetView.View.ToOpenXml();

        var pane = SetupPane(sheetView, svcm, xlWorksheet);

        SetTopLeftCell(sheetView, xlWorksheet);

        sheetView.RemoveAllChildren<Selection>();
        svcm.SetElement(XLSheetViewContents.Selection, null);

        if (xlWorksheet.SelectedRanges.Count > 0 || xlWorksheet.ActiveCell is not null)
            SetupSelections(sheetView, svcm, xlWorksheet, pane);

        SetZoomScales(sheetView, xlWorksheet);
    }

    private static void SetBooleanViewProperties(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        sheetView.TabSelected = xlWorksheet.TabSelected ? true : null;
        sheetView.RightToLeft = xlWorksheet.RightToLeft ? true : null;
        sheetView.ShowFormulas = xlWorksheet.ShowFormulas ? true : null;
        sheetView.ShowGridLines = xlWorksheet.ShowGridLines ? null : false;
        sheetView.ShowOutlineSymbols = xlWorksheet.ShowOutlineSymbols ? null : false;
        sheetView.ShowRowColHeaders = xlWorksheet.ShowRowColHeaders ? null : false;
        sheetView.ShowRuler = xlWorksheet.ShowRuler ? null : false;
        sheetView.ShowWhiteSpace = xlWorksheet.ShowWhiteSpace ? null : false;
        sheetView.ShowZeros = xlWorksheet.ShowZeros ? null : false;
    }

    private static Pane? SetupPane(SheetView sheetView, XLSheetViewContentManager svcm, XLWorksheet xlWorksheet)
    {
        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (pane == null)
        {
            pane = new Pane();
            sheetView.InsertAt(pane, 0);
        }

        svcm.SetElement(XLSheetViewContents.Pane, pane);

        pane.State = PaneStateValues.FrozenSplit;
        var hSplit = xlWorksheet.SheetView.SplitColumn;
        var ySplit = xlWorksheet.SheetView.SplitRow;

        pane.HorizontalSplit = hSplit;
        pane.VerticalSplit = ySplit;

        pane.ActivePane = GetActivePaneValue(hSplit, ySplit);

        pane.TopLeftCell = XLHelper.GetColumnLetterFromNumber(xlWorksheet.SheetView.SplitColumn + 1)
                           + (xlWorksheet.SheetView.SplitRow + 1);

        if (hSplit == 0 && ySplit == 0)
        {
            sheetView.RemoveAllChildren<Pane>();
            svcm.SetElement(XLSheetViewContents.Pane, null);
            return null;
        }

        return pane;
    }

    private static PaneValues GetActivePaneValue(int hSplit, int ySplit)
    {
        if (ySplit == 0 && hSplit == 0)
            return PaneValues.TopLeft;
        if (ySplit == 0)
            return PaneValues.TopRight;
        if (hSplit == 0)
            return PaneValues.BottomLeft;
        return PaneValues.BottomRight;
    }

    private static void SetTopLeftCell(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        if (!xlWorksheet.SheetView.TopLeftCellAddress.IsValid
            || xlWorksheet.SheetView.TopLeftCellAddress == new XLAddress(1, 1, fixedRow: false, fixedColumn: false))
            sheetView.TopLeftCell = null;
        else
            sheetView.TopLeftCell = xlWorksheet.SheetView.TopLeftCellAddress.ToString();
    }

    private static void SetupSelections(SheetView sheetView, XLSheetViewContentManager svcm,
        XLWorksheet xlWorksheet, Pane? pane)
    {
        var firstSelection = xlWorksheet.SelectedRanges.FirstOrDefault();

        if (pane != null)
        {
            PopulateSelection(new Selection()
            {
                Pane = pane.ActivePane
            });
        }

        PopulateSelection(new Selection());
        return;

        void PopulateSelection(Selection selection)
        {
            if (xlWorksheet.ActiveCell is not null)
                selection.ActiveCell = xlWorksheet.ActiveCell.Value.ToString();
            else if (firstSelection != null)
                selection.ActiveCell = firstSelection.RangeAddress.FirstAddress.ToStringRelative(false);

            var seqRef = new List<string> { selection.ActiveCell!.Value! };
            seqRef.AddRange(xlWorksheet.SelectedRanges.Select(range =>
                range.RangeAddress.FirstAddress.Equals(range.RangeAddress.LastAddress)
                    ? range.RangeAddress.FirstAddress.ToStringRelative(false)
                    : range.RangeAddress.ToStringRelative(false)));

            selection.SequenceOfReferences = new ListValue<StringValue>
                { InnerText = string.Join(" ", seqRef.Distinct().ToArray()) };

            sheetView.InsertAfter(selection, svcm.GetPreviousElementFor(XLSheetViewContents.Selection));
            svcm.SetElement(XLSheetViewContents.Selection, selection);
        }
    }

    private static void SetZoomScales(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        sheetView.ZoomScale = xlWorksheet.SheetView.ZoomScale == 100
            ? null
            : (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScale));

        sheetView.ZoomScaleNormal = xlWorksheet.SheetView.ZoomScaleNormal == 100
            ? null
            : (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScaleNormal));

        sheetView.ZoomScalePageLayoutView = xlWorksheet.SheetView.ZoomScalePageLayoutView == 100
            ? null
            : (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScalePageLayoutView));

        sheetView.ZoomScaleSheetLayoutView = xlWorksheet.SheetView.ZoomScaleSheetLayoutView == 100
            ? null
            : (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScaleSheetLayoutView));
    }

    internal static void WriteSheetFormatProperties(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        int maxOutlineColumn,
        int maxOutlineRow,
        out double worksheetColumnWidth)
    {
        worksheet.SheetFormatProperties ??= new SheetFormatProperties();

        cm.SetElement(XLWorksheetContents.SheetFormatProperties,
            worksheet.SheetFormatProperties);

        worksheet.SheetFormatProperties.DefaultRowHeight = xlWorksheet.RowHeight.SaveRound();

        if (xlWorksheet.RowHeightChanged)
            worksheet.SheetFormatProperties.CustomHeight = true;
        else
            worksheet.SheetFormatProperties.CustomHeight = null;

        worksheetColumnWidth = ColumnWriter.GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound();
        if (xlWorksheet.ColumnWidthChanged)
            worksheet.SheetFormatProperties.DefaultColumnWidth = worksheetColumnWidth;
        else
            worksheet.SheetFormatProperties.DefaultColumnWidth = null;

        if (maxOutlineColumn > 0)
            worksheet.SheetFormatProperties.OutlineLevelColumn = (byte)maxOutlineColumn;
        else
            worksheet.SheetFormatProperties.OutlineLevelColumn = null;

        if (maxOutlineRow > 0)
            worksheet.SheetFormatProperties.OutlineLevelRow = (byte)maxOutlineRow;
        else
            worksheet.SheetFormatProperties.OutlineLevelRow = null;
    }
}