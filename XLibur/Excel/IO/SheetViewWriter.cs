using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;
using XLibur.Utils;

namespace XLibur.Excel.IO;

internal static class SheetViewWriter
{
    internal static void WriteSheetProperties(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        var sheetProperties = worksheet.SheetProperties;
        var loadedEmpty = sheetProperties is not null && SchemaDefault.IsEmpty(sheetProperties);
        sheetProperties ??= new SheetProperties();

        sheetProperties.TabColor = xlWorksheet.SheetView.TabColor.HasValue
            ? new TabColor().FromXLiburColor<TabColor>(xlWorksheet.SheetView.TabColor)
            : null;

        WriteOutlineProperties(sheetProperties, xlWorksheet);

        if (sheetProperties.PageSetupProperties == null
            && (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0))
            sheetProperties.PageSetupProperties = new PageSetupProperties { FitToPage = true };

        SchemaDefault.Place(worksheet, cm, XLWorksheetContents.SheetProperties, sheetProperties, loadedEmpty);
    }

    /// <summary>
    /// Writes <c>&lt;outlinePr&gt;</c> only when summary rows are above or summary columns are on
    /// the left, as Excel does. Excel writes neither <c>summaryBelow="1"</c> nor
    /// <c>summaryRight="1"</c>, whether or not the sheet has an outline, and drops them from a file
    /// it saves again. A loaded file that has them keeps them (see <see cref="SchemaDefault"/>).
    /// </summary>
    private static void WriteOutlineProperties(SheetProperties sheetProperties, XLWorksheet xlWorksheet)
    {
        var outlineProperties = sheetProperties.OutlineProperties;
        var loadedEmpty = outlineProperties is not null && SchemaDefault.IsEmpty(outlineProperties);
        outlineProperties ??= new OutlineProperties();

        outlineProperties.SummaryBelow = SchemaDefault.Bool(outlineProperties.SummaryBelow,
            xlWorksheet.Outline.SummaryVLocation == XLOutlineSummaryVLocation.Bottom, schemaDefault: true);
        outlineProperties.SummaryRight = SchemaDefault.Bool(outlineProperties.SummaryRight,
            xlWorksheet.Outline.SummaryHLocation == XLOutlineSummaryHLocation.Right, schemaDefault: true);

        if (SchemaDefault.IsEmpty(outlineProperties) && !loadedEmpty)
            sheetProperties.OutlineProperties = null;
        else if (outlineProperties.Parent is null)
            sheetProperties.OutlineProperties = outlineProperties;
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

        // A selection the file wrote without an active cell is written without one again, while it
        // is the selection the file had. Excel reads the absence as a choice of its own, which an
        // active cell taken from the selection would replace.
        var loadedSelections = sheetView.Elements<Selection>().ToList();
        var deriveActiveCell = loadedSelections.Count == 0
                               || loadedSelections.Exists(s => s.ActiveCell is not null)
                               || loadedSelections.Exists(s =>
                                   s.SequenceOfReferences?.InnerText != SelectedRangesText(xlWorksheet));

        sheetView.RemoveAllChildren<Selection>();
        svcm.SetElement(XLSheetViewContents.Selection, null);

        if (xlWorksheet.SelectedRanges.Count > 0 || xlWorksheet.ActiveCell is not null)
            SetupSelections(sheetView, svcm, xlWorksheet, pane, deriveActiveCell);

        SetZoomScales(sheetView, xlWorksheet);
    }

    private static void SetBooleanViewProperties(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        var view = xlWorksheet.SheetView;
        sheetView.TabSelected = view.TabSelected ? true : null;
        sheetView.RightToLeft = view.RightToLeft ? true : null;
        sheetView.ShowFormulas = view.ShowFormulas ? true : null;
        sheetView.ShowGridLines = view.ShowGridLines ? null : false;
        sheetView.ShowOutlineSymbols = view.ShowOutlineSymbols ? null : false;
        sheetView.ShowRowColHeaders = view.ShowRowColHeaders ? null : false;
        sheetView.ShowRuler = view.ShowRuler ? null : false;
        sheetView.ShowWhiteSpace = view.ShowWhiteSpace ? null : false;
        sheetView.ShowZeros = view.ShowZeros ? null : false;
    }

    private static Pane? SetupPane(SheetView sheetView, XLSheetViewContentManager svcm, XLWorksheet xlWorksheet)
    {
        // XLPaneSettings owns the decision; this method owns only the emission. The streaming
        // writer resolves the same way and maps the result onto raw attribute strings.
        var settings = XLPaneSettings.Resolve(
            xlWorksheet.SheetView.SplitColumn,
            xlWorksheet.SheetView.SplitRow,
            xlWorksheet.SheetView.FreezePanes,
            xlWorksheet.SheetView.PaneTopLeftCellAddress,
            xlWorksheet.ActiveCell);

        // Asking the resolver first removes the dead work in the original, which built a Pane,
        // filled it in, and then removed it again when neither axis was split.
        if (!settings.HasPane)
        {
            sheetView.RemoveAllChildren<Pane>();
            svcm.SetElement(XLSheetViewContents.Pane, null);
            return null;
        }

        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (pane == null)
        {
            pane = new Pane();
            sheetView.InsertAt(pane, 0);
        }

        svcm.SetElement(XLSheetViewContents.Pane, pane);

        pane.HorizontalSplit = settings.SplitColumn;
        pane.VerticalSplit = settings.SplitRow;
        pane.TopLeftCell = settings.TopLeftCell;
        pane.ActivePane = ToOpenXml(settings.ActivePane);
        pane.State = ToOpenXml(settings.State);

        return pane;
    }

    private static PaneValues ToOpenXml(XLPaneCorner corner) => corner switch
    {
        XLPaneCorner.TopRight => PaneValues.TopRight,
        XLPaneCorner.BottomLeft => PaneValues.BottomLeft,
        XLPaneCorner.BottomRight => PaneValues.BottomRight,
        _ => PaneValues.TopLeft,
    };

    /// <remarks>
    /// Total. <see cref="XLPaneState.Frozen"/> and <see cref="XLPaneState.Split"/> both come out of
    /// <see cref="XLPaneSettings.Resolve"/>; <see cref="XLPaneState.FrozenSplit"/> does not, because
    /// the reader normalises it onto a frozen pane. The arm exists so this stays a translation
    /// rather than a coercion if the model ever grows a third state.
    /// </remarks>
    private static PaneStateValues ToOpenXml(XLPaneState state) => state switch
    {
        XLPaneState.FrozenSplit => PaneStateValues.FrozenSplit,
        XLPaneState.Split => PaneStateValues.Split,
        _ => PaneStateValues.Frozen,
    };

    private static void SetTopLeftCell(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        if (!xlWorksheet.SheetView.TopLeftCellAddress.IsValid
            || xlWorksheet.SheetView.TopLeftCellAddress == new XLAddress(1, 1, fixedRow: false, fixedColumn: false))
            sheetView.TopLeftCell = null;
        else
            sheetView.TopLeftCell = xlWorksheet.SheetView.TopLeftCellAddress.ToString();
    }

    private static void SetupSelections(SheetView sheetView, XLSheetViewContentManager svcm,
        XLWorksheet xlWorksheet, Pane? pane, bool deriveActiveCell)
    {
        var firstSelection = deriveActiveCell ? xlWorksheet.SelectedRanges.FirstOrDefault() : null;

        if (pane != null)
        {
            PopulateSelection(new Selection
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

            var seqRef = new List<string>();
            if (selection.ActiveCell is not null)
                seqRef.Add(selection.ActiveCell.Value!);

            seqRef.AddRange(SelectedRanges(xlWorksheet));

            selection.SequenceOfReferences = new ListValue<StringValue>
            { InnerText = string.Join(" ", seqRef.Distinct().ToArray()) };

            sheetView.InsertAfter(selection, svcm.GetPreviousElementFor(XLSheetViewContents.Selection));
            svcm.SetElement(XLSheetViewContents.Selection, selection);
        }
    }

    /// <summary>The selected ranges as a selection's <c>sqref</c> lists them.</summary>
    private static IEnumerable<string> SelectedRanges(XLWorksheet xlWorksheet) =>
        xlWorksheet.SelectedRanges.Select(range =>
            range.RangeAddress.FirstAddress.Equals(range.RangeAddress.LastAddress)
                ? range.RangeAddress.FirstAddress.ToStringRelative(false)
                : range.RangeAddress.ToStringRelative(false));

    /// <summary>The <c>sqref</c> of the selected ranges, with no active cell.</summary>
    private static string SelectedRangesText(XLWorksheet xlWorksheet) =>
        string.Join(" ", SelectedRanges(xlWorksheet).Distinct());

    private static void SetZoomScales(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        sheetView.ZoomScale = xlWorksheet.SheetView.ZoomScale == 100
            ? null
            : (uint)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScale));

        sheetView.ZoomScaleNormal = xlWorksheet.SheetView.ZoomScaleNormal == 100
            ? null
            : (uint)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScaleNormal));

        sheetView.ZoomScalePageLayoutView = xlWorksheet.SheetView.ZoomScalePageLayoutView == 100
            ? null
            : (uint)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScalePageLayoutView));

        sheetView.ZoomScaleSheetLayoutView = xlWorksheet.SheetView.ZoomScaleSheetLayoutView == 100
            ? null
            : (uint)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScaleSheetLayoutView));
    }

    internal static void WriteSheetFormatProperties(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        int maxOutlineColumn,
        int maxOutlineRow)
    {
        worksheet.SheetFormatProperties ??= new SheetFormatProperties();

        cm.SetElement(XLWorksheetContents.SheetFormatProperties,
            worksheet.SheetFormatProperties);

        worksheet.SheetFormatProperties.DefaultRowHeight = xlWorksheet.RowHeight.SaveRound();

        if (xlWorksheet.RowHeightChanged)
            worksheet.SheetFormatProperties.CustomHeight = true;
        else
            worksheet.SheetFormatProperties.CustomHeight = null;

        // ColumnWriter used to be handed this value; it now resolves the sheet default itself
        // through XLColumnSettings, so the width rule has two callers and no stored copy.
        var worksheetColumnWidth = ColumnWriter.GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound();
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
