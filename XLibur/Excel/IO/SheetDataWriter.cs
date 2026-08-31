using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;
using XLibur.Excel.Tables;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class SheetDataWriter
{
    internal static void StreamSheetData(XmlWriter xml, XLWorksheet xlWorksheet, SaveContext context,
        SaveOptions options)
    {
        var maxColumn = GetMaxColumn(xlWorksheet);

        xml.WriteStartElement("sheetData", Main2006SsNs);

        // Evaluating a dirty dynamic-array formula spills into its footprint, which both creates the
        // cells the write loop has to visit and sets the Range that identifies them. Left to the
        // per-cell evaluation below, a spill triggered part-way through the pass would land behind
        // the enumerator: the anchor would claim a ref the file has no cells for. Do it up front, so
        // the footprints and the enumerator both see the final grid.
        if (options.EvaluateFormulasBeforeSaving)
            EvaluateDirtyFormulas(xlWorksheet);

        var tableTotalCells = CollectTableTotalCells(xlWorksheet);
        var cachedResultFormulas = CollectCachedResultFormulas(xlWorksheet);

        // A rather complicated state machine, so rows and cells can be written in a single loop
        var rowState = new RowWriterState();
        var rows = GetSortedRowNumbers(xlWorksheet);
        var cellCtx = new CellWriteContext
        {
            CellsCollection = xlWorksheet.Internals.CellsCollection,
            CellRef = new char[CellXmlWriter.CellRefBufferLength],
            SaveContext = context,
            SaveOptions = options,
            TableTotalCells = tableTotalCells,
            CachedResultFormulas = cachedResultFormulas,
            Use1904DateSystem = xlWorksheet.Workbook.Use1904DateSystem,
        };
        uint rowStyleId = 0;
        XLStyleValue? lastCachedStyle = null;
        uint lastCachedStyleId = 0;
        var enumerator = new XLCellsCollection.SlicesEnumerator(Area.Full, cellCtx.CellsCollection);
        while (enumerator.MoveNext())
        {
            var point = enumerator.Current;
            var currentRowNumber = point.Row;

            WriteIntermediateRows(xml, xlWorksheet, rows, currentRowNumber, maxColumn, context, ref rowState);

            // Resolve the value and its share-string flag once for both the blank-and-empty check
            // and the value write below, avoiding a second ValueSlice traversal per cell.
            var cellValue = cellCtx.CellsCollection.ValueSlice.GetCellValueAndShareString(point, out var shareString);
            if (IsBlankAndEmpty(cellValue, cellCtx.CellsCollection, point))
                continue;

            if (rowState.OpenedRowNumber != currentRowNumber)
            {
                if (rowState.IsRowOpened)
                    xml.WriteEndElement(); // row

                rowStyleId = ResolveRowStyleId(xlWorksheet, currentRowNumber, ref rowState.RowPropIndex, context);

                xlWorksheet.Internals.RowsCollection.TryGetValue(currentRowNumber, out var row);
                WriteStartRow(xml, row, currentRowNumber, maxColumn, context);

                rowState.IsRowOpened = true;
                rowState.OpenedRowNumber = currentRowNumber;
            }

            var cellStyleId =
                ResolveCellStyleId(xlWorksheet, point, ref lastCachedStyle, ref lastCachedStyleId, context);

            WriteCellAtPoint(xml, ref cellCtx, point, rowStyleId, cellStyleId, cellValue, shareString);
        }

        if (rowState.IsRowOpened)
            xml.WriteEndElement(); // row

        WriteTrailingRows(xml, xlWorksheet, rows, rowState.RowPropIndex, context);

        xml.WriteEndElement(); // SheetData
    }

    private static HashSet<Point>? CollectTableTotalCells(XLWorksheet xlWorksheet)
    {
        if (xlWorksheet.Tables.Count == 0)
            return null;

        HashSet<Point>? cells = null;
        foreach (var table in xlWorksheet.Tables)
        {
            if (!table.ShowTotalsRow)
                continue;

            cells ??= [];
            foreach (var cell in table.TotalsRow()!.CellsUsed())
                cells.Add(((XLCell)cell).SheetPoint);
        }

        return cells;
    }

    /// <summary>
    /// Evaluate every dirty formula on the sheet, so the grid the write pass walks is the final one.
    /// </summary>
    /// <remarks>
    /// The formulas are collected before any of them is evaluated: a spill writes into the value
    /// slice and can extend the sheet's used range, which must not happen under an open enumerator.
    /// An array formula is held by every cell of its range, so the same instance is collected more
    /// than once; the second dirty check makes the repeats free, and also skips whatever a fallback
    /// to full recalculation already cleaned.
    /// </remarks>
    private static void EvaluateDirtyFormulas(XLWorksheet xlWorksheet)
    {
        List<(Point Point, XLCellFormula Formula)>? dirty = null;
        using (var enumerator = xlWorksheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full))
        {
            while (enumerator.MoveNext())
            {
                var formula = enumerator.Current;
                if (formula.IsDirty())
                    (dirty ??= []).Add((enumerator.Point, formula));
            }
        }

        if (dirty is null)
            return;

        foreach (var (point, formula) in dirty)
        {
            if (formula.IsDirty())
                EvaluateFormulaForSave(xlWorksheet, formula, point);
        }
    }

    /// <summary>
    /// Every formula on the sheet that stores its results in cells other than the one holding it, or
    /// <c>null</c> when the sheet has none (the common case, so callers pay nothing).
    /// </summary>
    /// <remarks>
    /// A dynamic array and a data table both keep their formula only in the master cell; the rest of
    /// the footprint holds cached results with no <c>&lt;f&gt;</c> of its own. Those cells must still
    /// be written as formula results (<c>t="str"</c> and a <c>&lt;v&gt;</c>), because Excel reads a
    /// shared-string or inline-string cell inside a spill footprint as content occupying the range,
    /// and renders the spill as <c>#VALUE!</c> everywhere below the anchor. Master cells are excluded
    /// implicitly: they hold a formula and so never reach the value-only path. Classic array formulas
    /// need no entry here — the formula slice holds them across the whole range, so every cell of one
    /// already takes the formula path.
    /// <para>
    /// The formulas are held rather than their footprints, so a <see cref="XLCellFormula.Range"/> that
    /// moves after this point is still read correctly. Evaluation is meant to be finished before the
    /// write pass starts, but a footprint snapshotted here would silently go stale if it were not.
    /// </para>
    /// </remarks>
    private static List<XLCellFormula>? CollectCachedResultFormulas(XLWorksheet xlWorksheet)
    {
        List<XLCellFormula>? formulas = null;
        using var enumerator = xlWorksheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            var formula = enumerator.Current;
            if (formula.IsDynamicArray || formula.Type == FormulaType.DataTable)
                (formulas ??= []).Add(formula);
        }

        return formulas;
    }

    private static bool IsCachedResultCell(List<XLCellFormula>? formulas, Point point)
    {
        if (formulas is null)
            return false;

        foreach (var formula in formulas)
        {
            var range = formula.Range;
            if (range != default && range.Contains(point))
                return true;
        }

        return false;
    }

    private static List<int> GetSortedRowNumbers(XLWorksheet xlWorksheet)
    {
        if (xlWorksheet.Internals.RowsCollection.Count <= 0)
            return [];

        var rows = xlWorksheet.Internals.RowsCollection.Keys.ToList();
        rows.Sort();
        return rows;
    }

    private static void WriteIntermediateRows(
        XmlWriter xml, XLWorksheet xlWorksheet, List<int> rows,
        int currentRowNumber, int maxColumn, SaveContext context,
        ref RowWriterState state)
    {
        while (state.RowPropIndex < rows.Count && rows[state.RowPropIndex] < currentRowNumber)
        {
            if (state.IsRowOpened)
            {
                xml.WriteEndElement(); // row
                state.IsRowOpened = false;
            }

            var rowNumber = rows[state.RowPropIndex];
            var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
            if (RowHasCustomProps(xlRow))
            {
                WriteStartRow(xml, xlRow, rowNumber, maxColumn, context);
                state.IsRowOpened = true;
                state.OpenedRowNumber = rowNumber;
            }

            state.RowPropIndex++;
        }
    }

    private static bool RowHasCustomProps(XLRow xlRow)
    {
        return xlRow.HeightChanged ||
               xlRow.IsHidden ||
               xlRow.StyleValue != xlRow.Worksheet.StyleValue ||
               xlRow.Collapsed ||
               xlRow.OutlineLevel > 0;
    }

    private static bool IsBlankAndEmpty(XLCellValue cellValue, XLCellsCollection cellsCollection, Point point)
    {
        if (cellValue.Type != XLDataType.Blank)
            return false;

        var xlCell = cellsCollection.GetCell(point);
        return xlCell.IsEmpty(XLCellsUsedOptions.All
                              & ~XLCellsUsedOptions.ConditionalFormats
                              & ~XLCellsUsedOptions.DataValidation
                              & ~XLCellsUsedOptions.MergedRanges);
    }

    private static uint ResolveRowStyleId(XLWorksheet xlWorksheet, int currentRowNumber,
        ref int rowPropIndex, SaveContext context)
    {
        if (xlWorksheet.Internals.RowsCollection.TryGetValue(currentRowNumber, out var row))
        {
            rowPropIndex++;
            return context.SharedStyles[row.StyleValue].StyleId;
        }

        return 0;
    }

    private static uint ResolveCellStyleId(XLWorksheet xlWorksheet, Point point,
        ref XLStyleValue? lastCachedStyle, ref uint lastCachedStyleId, SaveContext context)
    {
        var cellStyleValue = xlWorksheet.GetStyleValue(point);
        if (ReferenceEquals(cellStyleValue, lastCachedStyle))
            return lastCachedStyleId;

        lastCachedStyle = cellStyleValue;
        lastCachedStyleId = context.SharedStyles[cellStyleValue].StyleId;
        return lastCachedStyleId;
    }

    private static void WriteCellAtPoint(XmlWriter xml, ref CellWriteContext ctx,
        Point point, uint rowStyleId, uint cellStyleId, XLCellValue cellValue, bool shareString)
    {
        var formula = ctx.CellsCollection.FormulaSlice.Get(point);
        if (formula is not null)
        {
            WriteFormulaCellDirect(xml, ref ctx, point, formula, cellStyleId);
            return;
        }

        if (ctx.TableTotalCells is not null && ctx.TableTotalCells.Contains(point))
        {
            WriteTotalLabelCellDirect(xml, ref ctx, point, cellStyleId);
            return;
        }

        // Value and share-string flag were already resolved by the caller (and reused for the
        // blank-and-empty check), so no second ValueSlice traversal is needed here.
        if (cellValue.Type != XLDataType.Blank)
        {
            if (IsCachedResultCell(ctx.CachedResultFormulas, point))
                WriteCachedResultCell(xml, ref ctx, point, cellStyleId, cellValue);
            else
                WriteValueOnlyCell(xml, ref ctx, point, cellStyleId, cellValue, shareString);
        }
        else if (rowStyleId != cellStyleId)
        {
            WriteBlankStyledCell(xml, ctx.CellsCollection, point, ctx.CellRef, cellStyleId);
        }
    }

    /// <summary>
    /// Write a cell that has a formula directly from slice data, without allocating an
    /// <see cref="XLCell"/> wrapper. Mirrors the legacy <c>WriteCellWithFormula</c> +
    /// <c>WriteStartCell</c> path.
    /// </summary>
    private static void WriteFormulaCellDirect(XmlWriter xml, ref CellWriteContext ctx,
        Point point, XLCellFormula formula, uint cellStyleId)
    {
        var cellsCollection = ctx.CellsCollection;
        var xlWorksheet = cellsCollection.Worksheet;
        var saveContext = ctx.SaveContext;

        if (ctx.SaveOptions.EvaluateFormulasBeforeSaving && formula.IsDirty())
            EvaluateFormulaForSave(xlWorksheet, formula, point);

        // Determine cell type from cached value (preserves type round-trip for formulas
        // whose evaluation is unsupported).
        var cachedValue = cellsCollection.ValueSlice.GetCellValue(point);
        var cachedValueType = cachedValue.Type;
        var dataType = cachedValueType != XLDataType.Blank ? CellXmlWriter.GetFormulaCellType(cachedValueType) : null;

        Span<char> cellRefSpan = ctx.CellRef;
        var cellRefLen = point.Format(cellRefSpan);
        ref readonly var misc = ref cellsCollection.MiscSlice[point];

        // Compute "cm" attribute: explicit MiscSlice override, or workbook-wide dynamic-array
        // metadata index for dynamic-array formulas without an explicit override.
        var cmIndex = misc.CellMetaIndex;
        if (cmIndex is null && formula.IsDynamicArray && saveContext.DynamicArrayMetaIndex is not null)
            cmIndex = saveContext.DynamicArrayMetaIndex.Value;

        WriteStartFormulaCellDirect(xml, ctx.CellRef, cellRefLen, dataType, cellStyleId, in misc, cmIndex);

        if (formula.Type == FormulaType.DataTable)
        {
            WriteDataTableFormula(xml, formula);
        }
        else if (formula.IsDynamicArray)
        {
            // A dynamic-array formula lives only in its anchor cell (spilled cells are
            // formula-less and round-trip as plain cached values). Excel serialises it as an
            // array formula whose ref is the spill footprint, paired with the cm dynamic-array
            // metadata (written above). Before the first spill the footprint is unknown, so use
            // the 1x1 anchor.
            xml.WriteStartElement("f", Main2006SsNs);
            xml.WriteAttributeString("t", "array");
            var spillRange = formula.Range == default ? new Area(point) : formula.Range;
            var spillAddress = XLRangeAddress.FromSheetRange(xlWorksheet, spillRange);
            xml.WriteAttributeString("ref", spillAddress.ToStringRelative());
            xml.WriteString(formula.A1);
            xml.WriteEndElement(); // f
        }
        else if (formula.Type == FormulaType.Array)
        {
            var isMasterCell = formula.Range.FirstPoint == point;
            if (isMasterCell)
            {
                xml.WriteStartElement("f", Main2006SsNs);
                xml.WriteAttributeString("t", "array");
                var rangeAddress = XLRangeAddress.FromSheetRange(xlWorksheet, formula.Range);
                xml.WriteAttributeString("ref", rangeAddress.ToStringRelative());
                xml.WriteString(formula.A1);
                xml.WriteEndElement(); // f
            }
        }
        else
        {
            xml.WriteStartElement("f", Main2006SsNs);
            xml.WriteString(formula.A1);
            xml.WriteEndElement(); // f
        }

        // Write cached value if present and the formula isn't dirty. Spilled (non-master)
        // array-formula cells also fall through here so their cached values round-trip.
        if (cachedValueType != XLDataType.Blank && formula.IsClean())
        {
            WriteCachedFormulaValue(xml, cachedValue, ctx.Use1904DateSystem);
        }

        xml.WriteEndElement(); // cell
    }

    private static void EvaluateFormulaForSave(XLWorksheet xlWorksheet, XLCellFormula formula, Point point)
    {
        try
        {
            var workbook = xlWorksheet.Workbook;
            if (!workbook.CalcEngine.TryEvaluateSingleCell(formula, point, xlWorksheet))
                workbook.CalcEngine.Recalculate(workbook, null);
        }
        catch
        {
            // Match XLCell.Evaluate(false) tolerance: unimplemented features should not
            // abort the save. The cell is left with whatever cached value (if any) it
            // already has.
        }
    }

    /// <summary>
    /// Variant of <see cref="WriteStartCellDirect"/> that takes a pre-computed <c>cm</c>
    /// attribute value. Needed for formula cells where the dynamic-array metadata index is
    /// applied as a fallback when <see cref="XLMiscSliceContent.CellMetaIndex"/> is null.
    /// </summary>
    private static void WriteStartFormulaCellDirect(XmlWriter w, char[] reference, int referenceLength,
        string? dataType, uint styleId, in XLMiscSliceContent misc, uint? cmIndex)
    {
        CellXmlWriter.WriteCellStart(w, reference, referenceLength, dataType, styleId);
        CellXmlWriter.WriteCellMetaAttributes(w, misc.HasPhonetic, cmIndex, misc.ValueMetaIndex);
    }

    /// <summary>
    /// Write the cached value of a formula cell. Text is emitted inline; formulas can only
    /// store an inline-string text result, never a shared-string reference.
    /// </summary>
    private static void WriteCachedFormulaValue(XmlWriter w, XLCellValue cellValue, bool use1904DateSystem)
    {
        switch (cellValue.Type)
        {
            case XLDataType.Blank:
                return;
            case XLDataType.Text:
                CellXmlWriter.WriteStringValue(w, cellValue.GetText());
                break;
            default:
                CellXmlWriter.WriteNonTextValue(w, cellValue, use1904DateSystem);
                break;
        }
    }

    /// <summary>
    /// Write a totals-row label cell directly from slice data. The cell is in
    /// <see cref="CellWriteContext.TableTotalCells"/> but has no formula — it carries either
    /// a label (e.g. "Total") or nothing.
    /// </summary>
    private static void WriteTotalLabelCellDirect(XmlWriter xml, ref CellWriteContext ctx,
        Point point, uint cellStyleId)
    {
        var cellsCollection = ctx.CellsCollection;
        var xlWorksheet = cellsCollection.Worksheet;

        XLTable? containingTable = null;
        foreach (var table in xlWorksheet.Tables)
        {
            if (table.Area.Contains(point))
            {
                containingTable = table;
                break;
            }
        }

        XLTableField? field = null;
        if (containingTable is not null)
        {
            foreach (var f in containingTable.Fields)
            {
                if (f.Column.ColumnNumber() == point.Column)
                {
                    field = (XLTableField)f;
                    break;
                }
            }
        }

        if (field is not null && !string.IsNullOrWhiteSpace(field.TotalsRowLabel))
        {
            var memorySstId = cellsCollection.ValueSlice.GetShareStringId(point);
            var sharedStringId = ctx.SaveContext.GetSharedStringId(memorySstId, point);

            Span<char> cellRefSpan = ctx.CellRef;
            var cellRefLen = point.Format(cellRefSpan);
            ref readonly var misc = ref cellsCollection.MiscSlice[point];

            WriteStartCellDirect(xml, ctx.CellRef, cellRefLen, "s", cellStyleId, in misc);
            CellXmlWriter.WriteSharedStringValue(xml, sharedStringId);
            xml.WriteEndElement(); // cell
        }
    }

    /// <summary>
    /// Write a cell that carries only the cached result of a dynamic array spilled into it. It has
    /// no formula of its own, but is typed and serialised like a formula cell so Excel treats it as
    /// part of the spill rather than as content blocking it.
    /// </summary>
    private static void WriteCachedResultCell(XmlWriter xml, ref CellWriteContext ctx,
        Point point, uint cellStyleId, XLCellValue cellValue)
    {
        Span<char> cellRefSpan = ctx.CellRef;
        var cellRefLen = point.Format(cellRefSpan);
        var dataType = CellXmlWriter.GetFormulaCellType(cellValue.Type);
        ref readonly var misc = ref ctx.CellsCollection.MiscSlice[point];

        WriteStartCellDirect(xml, ctx.CellRef, cellRefLen, dataType, cellStyleId, in misc);
        WriteCachedFormulaValue(xml, cellValue, ctx.Use1904DateSystem);
        xml.WriteEndElement(); // cell
    }

    private static void WriteValueOnlyCell(XmlWriter xml, ref CellWriteContext ctx,
        Point point, uint cellStyleId, XLCellValue cellValue, bool shareString)
    {
        Span<char> cellRefSpan = ctx.CellRef;
        var cellRefLen = point.Format(cellRefSpan);
        var dataType = CellXmlWriter.GetValueCellType(cellValue.Type, shareString);
        ref readonly var misc = ref ctx.CellsCollection.MiscSlice[point];

        WriteStartCellDirect(xml, ctx.CellRef, cellRefLen, dataType, cellStyleId, in misc);
        WriteCellValueDirect(xml, cellValue, shareString, point, ctx.CellsCollection, ctx.Use1904DateSystem,
            ctx.SaveContext);
        xml.WriteEndElement(); // cell
    }

    private static void WriteBlankStyledCell(XmlWriter xml, XLCellsCollection cellsCollection,
        Point point, char[] cellRef, uint cellStyleId)
    {
        Span<char> cellRefSpan = cellRef;
        var cellRefLen = point.Format(cellRefSpan);
        ref readonly var misc = ref cellsCollection.MiscSlice[point];

        WriteStartCellDirect(xml, cellRef, cellRefLen, null, cellStyleId, in misc);
        xml.WriteEndElement(); // cell
    }

    private static void WriteTrailingRows(XmlWriter xml, XLWorksheet xlWorksheet,
        List<int> rows, int rowPropIndex, SaveContext context)
    {
        while (rowPropIndex < rows.Count)
        {
            var rowNumber = rows[rowPropIndex];
            var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
            if (RowHasCustomProps(xlRow))
            {
                WriteStartRow(xml, xlRow, rowNumber, 0, context);
                xml.WriteEndElement(); // row
            }

            rowPropIndex++;
        }
    }

    private static void WriteStartRow(XmlWriter w, XLRow? xlRow, int rowNumber, int maxColumn, SaveContext context)
    {
        CellXmlWriter.WriteRowStart(w, rowNumber, maxColumn);

        if (xlRow is null)
            return;

        WriteRowAttributes(w, xlRow, context);
    }

    private static void WriteRowAttributes(XmlWriter w, XLRow xlRow, SaveContext context)
    {
        if (xlRow.HeightChanged)
        {
            var height = xlRow.Height.SaveRound();
            w.WriteStartAttribute("ht");
            w.WriteNumberValue(height);
            w.WriteEndAttribute();

            // Note that dyDescent automatically implies custom height
            w.WriteAttributeString("customHeight", TrueValue);
        }

        if (xlRow.IsHidden)
            w.WriteAttributeString("hidden", TrueValue);

        if (xlRow.StyleValue != xlRow.Worksheet.StyleValue)
        {
            var styleIndex = context.SharedStyles[xlRow.StyleValue].StyleId;
            w.WriteAttribute("s", styleIndex);
            w.WriteAttributeString("customFormat", TrueValue);
        }

        if (xlRow.Collapsed)
            w.WriteAttributeString("collapsed", TrueValue);

        if (xlRow.OutlineLevel > 0)
            w.WriteAttribute("outlineLevel", xlRow.OutlineLevel);

        if (xlRow.ShowPhonetic)
            w.WriteAttributeString("ph", TrueValue);

        if (xlRow.DyDescent is not null)
            w.WriteAttribute("dyDescent", X14Ac2009SsNs, xlRow.DyDescent.Value);

        // thickBot and thickTop attributes are not written, because Excel seems to determine adjustments
        // from cell borders on its own, and it would be rather costly to check each cell in each row.
        // If the row was adjusted when the cell had its border modified, then it would be fine to write
        // the thickBot/thickBot attributes.
    }

    private static void WriteDataTableFormula(XmlWriter xml, XLCellFormula xlFormula)
    {
        xml.WriteStartElement("f", Main2006SsNs);
        xml.WriteAttributeString("t", "dataTable");
        xml.WriteAttributeString("ref", xlFormula.Range.ToString());

        var is2D = xlFormula.Is2DDataTable;
        if (is2D)
            xml.WriteAttributeString("dt2D", TrueValue);

        if (xlFormula.IsRowDataTable)
            xml.WriteAttributeString("dtr", TrueValue);

        xml.WriteAttributeString("r1", xlFormula.Input1.ToString());
        if (xlFormula.Input1Deleted)
            xml.WriteAttributeString("del1", TrueValue);

        if (is2D)
            xml.WriteAttributeString("r2", xlFormula.Input2.ToString());

        if (xlFormula.Input2Deleted)
            xml.WriteAttributeString("del2", TrueValue);

        // Excel doesn't recalculate table formula on a load or on the click of a button or any kind of forced recalculation.
        // It is necessary to mark some precedent formula dirty (e.g., edit cell formula and enter in Excel).
        // By setting the CalculateCell, we ensure that Excel will calculate values of data table formula on load and
        // the user will see correct values.
        xml.WriteAttributeString("ca", TrueValue);

        xml.WriteEndElement(); // f
    }

    private static void WriteStartCellDirect(XmlWriter w, char[] reference, int referenceLength, string? dataType,
        uint styleId, in XLMiscSliceContent misc)
    {
        CellXmlWriter.WriteCellStart(w, reference, referenceLength, dataType, styleId);
        CellXmlWriter.WriteCellMetaAttributes(w, misc.HasPhonetic, misc.CellMetaIndex, misc.ValueMetaIndex);
    }

    private static void WriteCellValueDirect(XmlWriter w, XLCellValue cellValue, bool shareString,
        Point point, XLCellsCollection cellsCollection, bool use1904DateSystem, SaveContext context)
    {
        switch (cellValue.Type)
        {
            case XLDataType.Blank:
                return;
            case XLDataType.Text:
                WriteCellValueDirectText(w, cellValue, shareString, point, cellsCollection, context);
                break;
            default:
                CellXmlWriter.WriteNonTextValue(w, cellValue, use1904DateSystem);
                break;
        }
    }

    private static void WriteCellValueDirectText(XmlWriter w, XLCellValue cellValue, bool shareString,
        Point point, XLCellsCollection cellsCollection, SaveContext context)
    {
        if (shareString)
        {
            var memorySstId = cellsCollection.ValueSlice.GetShareStringId(point);
            var sharedStringId = context.GetSharedStringId(memorySstId, point);
            CellXmlWriter.WriteSharedStringValue(w, sharedStringId);
            return;
        }

        var richText = cellsCollection.ValueSlice.GetRichText(point);
        if (richText is null)
        {
            CellXmlWriter.WriteInlineString(w, cellValue.GetText());
            return;
        }

        w.WriteStartElement("is", Main2006SsNs);
        TextSerializer.WriteRichTextElements(w, richText, context);
        w.WriteEndElement(); // is
    }

    private struct RowWriterState
    {
        public bool IsRowOpened;
        public int OpenedRowNumber;
        public int RowPropIndex;
    }

    private ref struct CellWriteContext
    {
        public XLCellsCollection CellsCollection;
        public char[] CellRef;
        public SaveContext SaveContext;
        public SaveOptions SaveOptions;
        public HashSet<Point>? TableTotalCells;
        public List<XLCellFormula>? CachedResultFormulas;
        public bool Use1904DateSystem;
    }

    internal static int GetMaxColumn(XLWorksheet xlWorksheet)
    {
        var maxColumn = 0;

        if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
        {
            maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
        }

        if (xlWorksheet.Internals.ColumnsCollection.Count <= 0) return maxColumn;
        var maxColCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
        if (maxColCollection > maxColumn)
            maxColumn = maxColCollection;

        return maxColumn;
    }
}
