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

        sheetView.View = SchemaDefault.Enum(sheetView.View, xlWorksheet.SheetView.View.ToOpenXml(),
            SheetViewValues.Normal);

        var loadedPane = sheetView.Elements<Pane>().FirstOrDefault()?.OuterXml;
        var pane = SetupPane(sheetView, svcm, xlWorksheet);

        SetTopLeftCell(sheetView, xlWorksheet);

        var loadedSelections = sheetView.Elements<Selection>().ToList();

        // The selections the file had stay as it wrote them while the pane, the selected ranges and
        // the active cell are the ones the load read from them (#722). Only when the load read the
        // selection of the active pane, where Excel puts the cursor: with others before it, the
        // model cannot tell a cursor moved to the first one from one left where it was.
        if (loadedSelections.Count > 0 && pane?.OuterXml == loadedPane
                                       && IsOfActivePane(loadedSelections[0], pane)
                                       && IsSelectionAsLoaded(loadedSelections[0], xlWorksheet))
        {
            svcm.SetElement(XLSheetViewContents.Selection, loadedSelections[^1]);
            SetZoomScales(sheetView, xlWorksheet);
            return;
        }

        // A selection the file wrote without an active cell is written without one again, while it
        // is the selection the file had. Excel reads the absence as a choice of its own, which an
        // active cell taken from the selection would replace.
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

    /// <summary>
    /// Are the sheet's selected ranges and active cell the ones the load read from
    /// <paramref name="selection"/>, the first selection of the file's view? Ranges are compared by
    /// the cells they cover, so <c>G1:G1048576</c> is the column <c>G:G</c>.
    /// </summary>
    private static bool IsSelectionAsLoaded(Selection selection, XLWorksheet xlWorksheet)
    {
        var activeCell = selection.ActiveCell?.Value;
        if (activeCell is null)
        {
            if (xlWorksheet.ActiveCell is not null)
                return false;
        }
        else if (xlWorksheet.ActiveCell is not { } active || !Point.TryParse(activeCell, out var point) || point != active)
        {
            return false;
        }

        var tokens = selection.SequenceOfReferences?.InnerText?.Split(' ', StringSplitOptions.RemoveEmptyEntries) ?? [];
        var ranges = xlWorksheet.SelectedRanges.ToList();
        if (tokens.Length != ranges.Count)
            return false;

        for (var i = 0; i < tokens.Length; i++)
        {
            if (!Area.TryParse(tokens[i], out var loaded) || !loaded.Equals(Area.FromRangeAddress(ranges[i].RangeAddress)))
                return false;
        }

        return true;
    }

    /// <summary>
    /// Is <paramref name="selection"/> the selection of the view's active pane? Both default to the
    /// top-left pane, which is the only one a view without a pane has.
    /// </summary>
    private static bool IsOfActivePane(Selection selection, Pane? pane)
    {
        var selectionPane = selection.Pane is { HasValue: true } s ? s.Value : PaneValues.TopLeft;
        var activePane = pane?.ActivePane is { HasValue: true } a ? a.Value : PaneValues.TopLeft;
        return selectionPane == activePane;
    }

    private static void SetBooleanViewProperties(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        // A value the file spelled out stays, default or not (#722).
        var view = xlWorksheet.SheetView;
        sheetView.TabSelected = SchemaDefault.Bool(sheetView.TabSelected, view.TabSelected, false);
        sheetView.RightToLeft = SchemaDefault.Bool(sheetView.RightToLeft, view.RightToLeft, false);
        sheetView.ShowFormulas = SchemaDefault.Bool(sheetView.ShowFormulas, view.ShowFormulas, false);
        sheetView.ShowGridLines = SchemaDefault.Bool(sheetView.ShowGridLines, view.ShowGridLines, true);
        sheetView.ShowOutlineSymbols = SchemaDefault.Bool(sheetView.ShowOutlineSymbols, view.ShowOutlineSymbols, true);
        sheetView.ShowRowColHeaders = SchemaDefault.Bool(sheetView.ShowRowColHeaders, view.ShowRowColHeaders, true);
        sheetView.ShowRuler = SchemaDefault.Bool(sheetView.ShowRuler, view.ShowRuler, true);
        sheetView.ShowWhiteSpace = SchemaDefault.Bool(sheetView.ShowWhiteSpace, view.ShowWhiteSpace, true);
        sheetView.ShowZeros = SchemaDefault.Bool(sheetView.ShowZeros, view.ShowZeros, true);
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

    /// <summary>
    /// Writes the top-left cell unless it is A1. A file that spelled out A1 keeps it (#722).
    /// </summary>
    private static void SetTopLeftCell(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        var address = xlWorksheet.SheetView.TopLeftCellAddress;
        if (!address.IsValid)
        {
            sheetView.TopLeftCell = null;
            return;
        }

        var text = address.ToString();
        if (sheetView.TopLeftCell?.Value == text)
            return;

        // Two statements: a conditional expression has the type string, and its null would convert
        // to a StringValue holding no text rather than to no attribute.
        if (address == new XLAddress(1, 1, fixedRow: false, fixedColumn: false))
            sheetView.TopLeftCell = null;
        else
            sheetView.TopLeftCell = text;
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

    /// <summary>
    /// Writes each zoom scale that is not 100. A scale the file spelled out as 100, as Excel often
    /// writes <c>zoomScaleNormal</c>, stays (#722).
    /// </summary>
    private static void SetZoomScales(SheetView sheetView, XLWorksheet xlWorksheet)
    {
        var view = xlWorksheet.SheetView;
        sheetView.ZoomScale = ZoomScale(sheetView.ZoomScale, view.ZoomScale);
        sheetView.ZoomScaleNormal = ZoomScale(sheetView.ZoomScaleNormal, view.ZoomScaleNormal);
        sheetView.ZoomScalePageLayoutView = ZoomScale(sheetView.ZoomScalePageLayoutView, view.ZoomScalePageLayoutView);
        sheetView.ZoomScaleSheetLayoutView = ZoomScale(sheetView.ZoomScaleSheetLayoutView, view.ZoomScaleSheetLayoutView);
    }

    private static UInt32Value? ZoomScale(UInt32Value? loaded, int scale) =>
        SchemaDefault.UInt(loaded, (uint)Math.Clamp(scale, 10, 400), 100);

    /// <summary>
    /// Is the sheet's width the one the load read from <paramref name="loaded"/>? The load converts
    /// <c>defaultColWidth</c> through the width of the font's digits, which the save does not undo,
    /// so the file's value is kept while the width is still the one it loaded as (#722).
    /// </summary>
    private static bool IsWidthLoadedFrom(DoubleValue? loaded, XLWorksheet xlWorksheet) =>
        loaded is { HasValue: true }
        && Math.Abs(XLHelper.ConvertWidthToNoC(loaded.Value, xlWorksheet.Style.Font, xlWorksheet.Workbook)
                    - xlWorksheet.ColumnWidth) < XLHelper.Epsilon;

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

        var sheetFormat = worksheet.SheetFormatProperties;
        sheetFormat.DefaultRowHeight = SchemaDefault.Measure(sheetFormat.DefaultRowHeight, xlWorksheet.RowHeight.SaveRound());

        if (xlWorksheet.RowHeightChanged)
            worksheet.SheetFormatProperties.CustomHeight = true;
        else
            worksheet.SheetFormatProperties.CustomHeight = null;

        // ColumnWriter used to be handed this value; it now resolves the sheet default itself
        // through XLColumnSettings, so the width rule has two callers and no stored copy.
        var worksheetColumnWidth = ColumnWriter.GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound();
        if (xlWorksheet.ColumnWidthChanged)
        {
            if (!IsWidthLoadedFrom(sheetFormat.DefaultColumnWidth, xlWorksheet))
                sheetFormat.DefaultColumnWidth = SchemaDefault.Measure(sheetFormat.DefaultColumnWidth, worksheetColumnWidth);
        }
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
