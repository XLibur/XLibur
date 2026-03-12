using XLibur.Excel.ContentManagers;
using XLibur.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal sealed class SheetViewWriter
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

        if (xlWorksheet.TabSelected)
            sheetView.TabSelected = true;
        else
            sheetView.TabSelected = null;

        if (xlWorksheet.RightToLeft)
            sheetView.RightToLeft = true;
        else
            sheetView.RightToLeft = null;

        if (xlWorksheet.ShowFormulas)
            sheetView.ShowFormulas = true;
        else
            sheetView.ShowFormulas = null;

        if (xlWorksheet.ShowGridLines)
            sheetView.ShowGridLines = null;
        else
            sheetView.ShowGridLines = false;

        if (xlWorksheet.ShowOutlineSymbols)
            sheetView.ShowOutlineSymbols = null;
        else
            sheetView.ShowOutlineSymbols = false;

        if (xlWorksheet.ShowRowColHeaders)
            sheetView.ShowRowColHeaders = null;
        else
            sheetView.ShowRowColHeaders = false;

        if (xlWorksheet.ShowRuler)
            sheetView.ShowRuler = null;
        else
            sheetView.ShowRuler = false;

        if (xlWorksheet.ShowWhiteSpace)
            sheetView.ShowWhiteSpace = null;
        else
            sheetView.ShowWhiteSpace = false;

        if (xlWorksheet.ShowZeros)
            sheetView.ShowZeros = null;
        else
            sheetView.ShowZeros = false;

        if (xlWorksheet.RightToLeft)
            sheetView.RightToLeft = true;
        else
            sheetView.RightToLeft = null;

        if (xlWorksheet.SheetView.View == XLSheetViewOptions.Normal)
            sheetView.View = null;
        else
            sheetView.View = xlWorksheet.SheetView.View.ToOpenXml();

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

        // When panes are frozen, which part should move.
        PaneValues split;
        if (ySplit == 0 && hSplit == 0)
            split = PaneValues.TopLeft;
        else if (ySplit == 0 && hSplit != 0)
            split = PaneValues.TopRight;
        else if (ySplit != 0 && hSplit == 0)
            split = PaneValues.BottomLeft;
        else
            split = PaneValues.BottomRight;

        pane.ActivePane = split;

        pane.TopLeftCell = XLHelper.GetColumnLetterFromNumber(xlWorksheet.SheetView.SplitColumn + 1)
                           + (xlWorksheet.SheetView.SplitRow + 1);

        if (hSplit == 0 && ySplit == 0)
        {
            // We don't have a pane. Just a regular sheet.
            pane = null;
            sheetView.RemoveAllChildren<Pane>();
            svcm.SetElement(XLSheetViewContents.Pane, null);
        }

        // Do sheet view. Whether it's for a regular sheet or for the bottom-right pane
        if (!xlWorksheet.SheetView.TopLeftCellAddress.IsValid
            || xlWorksheet.SheetView.TopLeftCellAddress == new XLAddress(1, 1, fixedRow: false, fixedColumn: false))
            sheetView.TopLeftCell = null;
        else
            sheetView.TopLeftCell = xlWorksheet.SheetView.TopLeftCellAddress.ToString();

        if (xlWorksheet.SelectedRanges.Any() || xlWorksheet.ActiveCell is not null)
        {
            sheetView.RemoveAllChildren<Selection>();
            svcm.SetElement(XLSheetViewContents.Selection, null);

            var firstSelection = xlWorksheet.SelectedRanges.FirstOrDefault();

            void PopulateSelection(Selection selection)
            {
                if (xlWorksheet.ActiveCell is not null)
                    selection.ActiveCell = xlWorksheet.ActiveCell.Value.ToString();
                else if (firstSelection != null)
                    selection.ActiveCell = firstSelection.RangeAddress.FirstAddress.ToStringRelative(false);

                var seqRef = new List<string> { selection.ActiveCell!.Value! };
                seqRef.AddRange(xlWorksheet.SelectedRanges.Select(range =>
                {
                    return range.RangeAddress.FirstAddress.Equals(range.RangeAddress.LastAddress)
                        ? range.RangeAddress.FirstAddress.ToStringRelative(false)
                        : range.RangeAddress.ToStringRelative(false);
                }));

                selection.SequenceOfReferences = new ListValue<StringValue>
                    { InnerText = string.Join(" ", seqRef.Distinct().ToArray()) };

                sheetView.InsertAfter(selection, svcm.GetPreviousElementFor(XLSheetViewContents.Selection));
                svcm.SetElement(XLSheetViewContents.Selection, selection);
            }

            // If a pane exists, we need to set the active pane too
            // Yes, this might lead to 2 Selection elements!
            if (pane != null)
            {
                PopulateSelection(new Selection()
                {
                    Pane = pane.ActivePane
                });
            }

            PopulateSelection(new Selection());
        }

        if (xlWorksheet.SheetView.ZoomScale == 100)
            sheetView.ZoomScale = null;
        else
            sheetView.ZoomScale = (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScale));

        if (xlWorksheet.SheetView.ZoomScaleNormal == 100)
            sheetView.ZoomScaleNormal = null;
        else
            sheetView.ZoomScaleNormal = (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScaleNormal));

        if (xlWorksheet.SheetView.ZoomScalePageLayoutView == 100)
            sheetView.ZoomScalePageLayoutView = null;
        else
            sheetView.ZoomScalePageLayoutView =
                (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScalePageLayoutView));

        if (xlWorksheet.SheetView.ZoomScaleSheetLayoutView == 100)
            sheetView.ZoomScaleSheetLayoutView = null;
        else
            sheetView.ZoomScaleSheetLayoutView =
                (uint)System.Math.Max(10, System.Math.Min(400, xlWorksheet.SheetView.ZoomScaleSheetLayoutView));
    }

    internal static void WriteSheetFormatProperties(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        int maxOutlineColumn,
        int maxOutlineRow,
        out double worksheetColumnWidth)
    {
        if (worksheet.SheetFormatProperties == null)
            worksheet.SheetFormatProperties = new SheetFormatProperties();

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
