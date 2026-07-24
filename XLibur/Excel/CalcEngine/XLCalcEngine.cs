using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.CalcEngine.Functions;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// CalcEngine parses strings and returns Expression objects that can
/// be evaluated.
/// </summary>
/// <remarks>
/// <para>This class has three extensibility points:</para>
/// <para>Use the <b>RegisterFunction</b> method to define custom functions.</para>
/// </remarks>
internal sealed class XLCalcEngine : ISheetListener, IWorkbookListener
{
    private readonly CultureInfo _culture;
    private readonly ExpressionCache _cache;               // cache with parsed expressions
    private readonly FormulaParser _parser;
    private readonly CalculationVisitor _visitor;
    private DependencyTree? _dependencyTree;
    private XLCalculationChain? _chain;

    /// <summary>
    /// Set to true after <see cref="TryEvaluateSingleCell"/> clears a formula's dirty flag,
    /// indicating that future MarkDirty calls need a dependency tree to
    /// correctly propagate dirtiness to dependents.
    /// </summary>
    private bool _needsDependencyTree;

    /// <summary>
    /// The spill footprint of every dynamic-array formula, one entry per formula (a rectangle,
    /// not one entry per spilled cell). Spilled cells are formula-less, so this lets a read of one
    /// (during recalc) find the owning anchor and force it to evaluate first — see
    /// <see cref="TryGetDirtySpillOwner"/>. Rebuilt when the dependency tree is (re)built and kept
    /// in sync by <see cref="SpillDynamicArray"/>. Empty for workbooks without dynamic arrays.
    /// </summary>
    private readonly List<SpillFootprint> _spillOwners = new();

    /// <summary>
    /// A dynamic array's spill footprint (including the anchor at <see cref="Range"/>'s first
    /// point) together with the owning formula.
    /// </summary>
    private readonly struct SpillFootprint(uint sheetId, XLSheetRange range, XLCellFormula owner)
    {
        internal readonly uint SheetId = sheetId;
        internal readonly XLSheetRange Range = range;
        internal readonly XLCellFormula Owner = owner;
    }

    /// <summary>
    /// Whether any dynamic-array formula currently has a spill footprint. Used to keep the
    /// spill-owner lookup out of the hot cell-read path for the common no-dynamic-array case.
    /// </summary>
    internal bool HasSpillOwners => _spillOwners.Count > 0;

    public XLCalcEngine(CultureInfo culture)
    {
        _culture = culture;
        _cache = new ExpressionCache(this);
        var funcRegistry = GetFunctionTable();
        _parser = new FormulaParser(funcRegistry);
        _visitor = new CalculationVisitor(funcRegistry);
        _dependencyTree = null;
        _chain = null;
    }

    /// <summary>
    /// Parses a string into an <see cref="Formula"/>.
    /// </summary>
    /// <param name="expression">String to parse.</param>
    /// <returns>A formula that can be evaluated.</returns>
    public Formula Parse(string expression)
    {
        return _parser.GetAst(expression, isA1: true);
    }

    /// <summary>
    /// Add an array formula to the calc engine to manage dirty tracking and evaluation.
    /// </summary>
    internal void AddArrayFormula(XLSheetRange range, XLCellFormula arrayFormula, XLWorksheet sheet)
    {
        if (_chain is not null && _dependencyTree is not null)
        {
            _dependencyTree.AddFormula(new XLBookArea(sheet.Name, range), arrayFormula, sheet.Workbook);
            _chain.AppendArea(sheet.SheetId, range);
        }
    }

    /// <summary>
    /// Add a formula to the calc engine to manage dirty tracking and evaluation.
    /// </summary>
    internal void AddNormalFormula(XLBookPoint point, string sheetName, XLCellFormula formula, XLWorkbook workbook)
    {
        if (_chain is not null && _dependencyTree is not null)
        {
            var pointArea = new XLBookArea(sheetName, new XLSheetRange(point.Point, point.Point));
            _dependencyTree.AddFormula(pointArea, formula, workbook);
            _chain.AddLast(point);
        }
    }

    /// <summary>
    /// Remove formula from dependency tree (=precedents won't mark
    /// it as dirty) and remove <paramref name="point"/> from the chain.
    /// Note that even if formula is used by many cells (e.g. array formula),
    /// it is fully removed from dependency tree, but each cells referencing
    /// the formula must be removed individually from calc chain.
    /// </summary>
    internal void RemoveFormula(XLBookPoint point, XLCellFormula formula)
    {
        if (_chain is not null && _dependencyTree is not null)
        {
            _dependencyTree.RemoveFormula(formula);
            _chain.Remove(point);
        }
    }

    internal void OnAddedSheet(XLWorksheet sheet)
    {
        Purge(sheet.Workbook.WorksheetsInternal);
    }

    internal void OnDeletingSheet(XLWorksheet sheet)
    {
        Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnInsertAreaAndShiftDown(XLWorksheet sheet, XLSheetRange area)
    {
        Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnInsertAreaAndShiftRight(XLWorksheet sheet, XLSheetRange area)
    {
        Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnDeleteAreaAndShiftLeft(XLWorksheet sheet, XLSheetRange deletedRange)
    {
        Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnDeleteAreaAndShiftUp(XLWorksheet sheet, XLSheetRange deletedRange)
    {
        Purge(sheet.Workbook.WorksheetsInternal);
    }

    private void Purge(XLWorksheets sheets)
    {
        _dependencyTree = null;
        _chain = null;
        _needsDependencyTree = false;
        _spillOwners.Clear();

        // Mark everything as dirty, because there can be stale values
        foreach (var sheet in sheets)
        {
            sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(XLSheetRange.Full);
        }
    }

    internal void MarkDirty(XLWorksheet sheet, XLSheetPoint point)
    {
        MarkDirty(sheet, new XLSheetRange(point, point));
    }

    internal void MarkDirty(XLWorksheet sheet, XLSheetRange area)
    {
        if (_dependencyTree is null && _needsDependencyTree)
        {
            _dependencyTree = DependencyTree.CreateFrom(sheet.Workbook);
            RebuildSpillOwners(sheet.Workbook);
        }

        if (_dependencyTree is not null)
        {
            var bookArea = new XLBookArea(sheet.Name, area);
            _dependencyTree.MarkDirty(bookArea);
        }
    }

    /// <summary>
    /// Try to evaluate a single cell's formula directly, without building the full
    /// dependency tree and calculation chain. If the formula references dirty
    /// precedent cells (throws <see cref="GettingDataException"/>), falls back
    /// to full workbook recalculation which handles dependency ordering.
    /// </summary>
    /// <returns><c>true</c> if single-cell eval succeeded, <c>false</c> if full recalculate was used.</returns>
    internal bool TryEvaluateSingleCell(XLCellFormula formula, XLSheetPoint point, XLWorksheet sheet)
    {
        // DataTable formulas need the full chain for correct evaluation.
        if (formula.Type == FormulaType.DataTable)
        {
            Recalculate(sheet.Workbook, null);
            return false;
        }

        try
        {
            var valueSlice = sheet.Internals.CellsCollection.ValueSlice;
            if (formula.IsDynamicArray)
            {
                SpillDynamicArray(formula, point, sheet, recalculateSheetId: null);
            }
            else if (formula.Type == FormulaType.Normal)
            {
                var result = EvaluateFormula(
                    formula.A1,
                    sheet.Workbook,
                    sheet,
                    new XLAddress(sheet, point.Row, point.Column, true, true));
                valueSlice.SetCellValue(point, result.ToCellValue());
            }
            else if (formula.Type == FormulaType.Array)
            {
                var range = formula.Range;
                var leftTopCorner = range.FirstPoint;
                var masterCell = sheet.Cell(leftTopCorner.Row, leftTopCorner.Column);
                var array = EvaluateArrayFormula(formula.A1, masterCell, recalculateSheetId: null);
                var result = array.Broadcast(range.Height, range.Width);

                for (var rowIdx = 0; rowIdx < result.Height; ++rowIdx)
                {
                    for (var colIdx = 0; colIdx < result.Width; ++colIdx)
                    {
                        var cellValue = result[rowIdx, colIdx];
                        var row = range.FirstPoint.Row + rowIdx;
                        var column = range.FirstPoint.Column + colIdx;
                        valueSlice.SetCellValue(new XLSheetPoint(row, column), cellValue.ToCellValue());
                    }
                }
            }

            formula.MarkClean(sheet.Workbook);
            _needsDependencyTree = true;
            return true;
        }
        catch (GettingDataException)
        {
            // Formula depends on a dirty precedent cell — need the full
            // dependency-ordered recalculation to resolve it.
            Recalculate(sheet.Workbook, null);
            return false;
        }
    }

    /// <summary>
    /// Recalculate a workbook or a sheet.
    /// </summary>
    internal void Recalculate(XLWorkbook wb, uint? recalculateSheetId)
    {
        // Lazy, so initialize chain from wb, if it is empty
        if (_chain is null || _dependencyTree is null)
        {
            _chain = XLCalculationChain.CreateFrom(wb);
            _dependencyTree = DependencyTree.CreateFrom(wb);
            RebuildSpillOwners(wb);
        }

        var sheetIdMap = wb.WorksheetsInternal
            .ToDictionary<XLWorksheet, uint, (XLWorksheet Sheet, ValueSlice ValueSlice, FormulaSlice FormulaSlice)>(
                sheet => sheet.SheetId,
                sheet => (sheet, sheet.Internals.CellsCollection.ValueSlice, sheet.Internals.CellsCollection.FormulaSlice));

        // Each outer loop moves chain one cell ahead.
        while (_chain.MoveAhead())
        {
            RecalculateCurrentCell(_chain, sheetIdMap, recalculateSheetId);
        }

        // Super important to clean up the chain for next recalculation.
        // Chain contains shared data and not cleaning it would cause hard
        // to diagnose issues.
        _chain.Reset();
    }

    private void RecalculateCurrentCell(
        XLCalculationChain chain,
        Dictionary<uint, (XLWorksheet Sheet, ValueSlice ValueSlice, FormulaSlice FormulaSlice)> sheetIdMap,
        uint? recalculateSheetId)
    {
        while (true)
        {
            var current = chain.Current;
            var sheetId = current.SheetId;

            if (recalculateSheetId is not null && sheetId != recalculateSheetId.Value)
                break;

            if (!sheetIdMap.TryGetValue(sheetId, out var sheetInfo))
                throw new InvalidOperationException($"Unable to find sheet with sheetId {sheetId} for a point ${current.Point}.");

            if (chain.IsCurrentInCycle)
                throw new InvalidOperationException($"Formula in a cell '${sheetInfo.Sheet.Name}'!${current.Point} is part of a cycle.");

            var cellFormula = sheetInfo.FormulaSlice.Get(current.Point);
            if (cellFormula is null)
                throw new InvalidOperationException($"Calculation chain contains a '${sheetInfo.Sheet.Name}'!${current.Point}, but the cell doesn't contain formula.");

            if (cellFormula.IsClean(sheetInfo.Sheet.Workbook))
                break;

            try
            {
                ApplyFormula(cellFormula, current.Point, sheetInfo.Sheet, sheetInfo.ValueSlice, recalculateSheetId);
                cellFormula.MarkClean(sheetInfo.Sheet.Workbook);
                break;
            }
            catch (GettingDataException ex)
            {
                chain.MoveToCurrent(ex.Point);
            }
        }
    }

    private void ApplyFormula(XLCellFormula formula, XLSheetPoint appliedPoint, XLWorksheet sheet, ValueSlice valueSlice, uint? recalculateSheetId)
    {
        var formulaText = formula.A1;
        if (formula.IsDynamicArray)
        {
            // The formula lives only in the anchor cell (spilled cells are formula-less),
            // so the applied point is always the anchor.
            SpillDynamicArray(formula, appliedPoint, sheet, recalculateSheetId);
        }
        else if (formula.Type == FormulaType.Normal)
        {
            var single = EvaluateFormula(
                formulaText,
                sheet.Workbook,
                sheet,
                new XLAddress(sheet, appliedPoint.Row, appliedPoint.Column, true, true),
                recalculateSheetId: recalculateSheetId);
            valueSlice.SetCellValue(appliedPoint, single.ToCellValue());
        }
        else if (formula.Type == FormulaType.Array)
        {
            // The point can be any point in an array, so we can't use it.
            var range = formula.Range;
            var leftTopCorner = range.FirstPoint;
            var masterCell = sheet.Cell(leftTopCorner.Row, leftTopCorner.Column);
            var array = EvaluateArrayFormula(formulaText, masterCell, recalculateSheetId);

            // The array from formula can be smaller or larger than the
            // range of cells it should fit into. Broadcast it to the size.
            var result = array.Broadcast(range.Height, range.Width);

            // Copy value to the value slice
            for (var rowIdx = 0; rowIdx < result.Height; ++rowIdx)
            {
                for (var colIdx = 0; colIdx < result.Width; ++colIdx)
                {
                    var cellValue = result[rowIdx, colIdx];
                    var row = range.FirstPoint.Row + rowIdx;
                    var column = range.FirstPoint.Column + colIdx;
                    valueSlice.SetCellValue(new XLSheetPoint(row, column), cellValue.ToCellValue());
                }
            }
        }
        else
        {
            throw new NotImplementedException($"Evaluation of formula type '{formula.Type}' is not supported.");
        }
    }

    /// <summary>
    /// Evaluates a normal formula.
    /// </summary>
    /// <param name="expression">Expression to evaluate.</param>
    /// <param name="wb">Workbook where is formula being evaluated.</param>
    /// <param name="ws">Worksheet where is formula being evaluated.</param>
    /// <param name="address">Address of formula.</param>
    /// <param name="recursive">Should the data necessary for this formula (not deeper ones)
    /// be calculated recursively? Used only for non-cell calculations.</param>
    /// <param name="recalculateSheetId">
    /// If set, calculation  will allow dirty reads from other sheets than the passed one.
    /// </param>
    /// <returns>The value of the expression.</returns>
    /// <remarks>
    /// If you are going to evaluate the same expression several times,
    /// it is more efficient to parse it only once using the <see cref="Parse"/>
    /// method and then using the Expression.Evaluate method to evaluate
    /// the parsed expression.
    /// </remarks>
    internal ScalarValue EvaluateFormula(string expression, XLWorkbook? wb = null, XLWorksheet? ws = null, IXLAddress? address = null, bool recursive = false, uint? recalculateSheetId = null)
    {
        var ctx = new CalcContext(this, _culture, wb, ws, address, recursive)
        {
            RecalculateSheetId = recalculateSheetId
        };
        var result = EvaluateFormula(expression, ctx);
        if (CalcContext.UseImplicitIntersection)
        {
            result = result.Match(
                () => AnyValue.Blank,
                logical => logical,
                number => number,
                text => text,
                error => error,
                array => array[0, 0].ToAnyValue(),
                reference => reference);
        }

        return ToCellContentValue(result, ctx);
    }

    private Array EvaluateArrayFormula(string expression, XLCell masterCell, uint? recalculateSheetId)
    {
        var ctx = new CalcContext(this, _culture, masterCell)
        {
            IsArrayCalculation = true,
            RecalculateSheetId = recalculateSheetId
        };
        var result = EvaluateFormula(expression, ctx);
        if (result.TryPickSingleOrMultiValue(out var single, out var multi, ctx))
            return new ScalarArray(single, 1, 1);

        return multi!;
    }

    /// <summary>
    /// Evaluates a dynamic-array formula and spills its result across the anchor's
    /// footprint. Only the anchor holds the <see cref="XLCellFormula"/>; the remaining
    /// footprint cells receive values only. The computed footprint is stored back on
    /// <see cref="XLCellFormula.Range"/> so a later evaluation can clear a stale region.
    /// </summary>
    /// <remarks>
    /// The spill is blocked — the anchor gets <see cref="XLError.SpillRange"/> (<c>#SPILL!</c>)
    /// and nothing is written to the footprint — when the result would run past the sheet
    /// edge, or when any footprint cell (other than the anchor, and other than a cell owned by
    /// this formula's previous footprint) already holds a formula or a non-blank value.
    /// </remarks>
    private void SpillDynamicArray(XLCellFormula formula, XLSheetPoint anchor, XLWorksheet sheet, uint? recalculateSheetId)
    {
        var cells = sheet.Internals.CellsCollection;
        var valueSlice = cells.ValueSlice;
        var formulaSlice = cells.FormulaSlice;

        var masterCell = sheet.Cell(anchor.Row, anchor.Column);
        var array = EvaluateArrayFormula(formula.A1, masterCell, recalculateSheetId);

        var lastRow = anchor.Row + array.Height - 1;
        var lastColumn = anchor.Column + array.Width - 1;

        var previousRange = formula.Range;
        var anchorRange = new XLSheetRange(anchor);

        XLSheetRange newFootprint;
        var outOfBounds = lastRow > XLHelper.MaxRowNumber || lastColumn > XLHelper.MaxColumnNumber;
        if (outOfBounds || HasSpillCollision(anchor, lastRow, lastColumn, previousRange, valueSlice, formulaSlice))
        {
            ClearSpillFootprint(previousRange, anchorRange, valueSlice);
            valueSlice.SetCellValue(anchor, XLError.SpillRange);
            newFootprint = anchorRange;
        }
        else
        {
            newFootprint = new XLSheetRange(anchor, new XLSheetPoint(lastRow, lastColumn));

            // Erase any cell of the previous footprint that the new one no longer covers
            // (the array shrank or moved) before writing the fresh result.
            ClearSpillFootprint(previousRange, newFootprint, valueSlice);

            for (var rowOffset = 0; rowOffset < array.Height; ++rowOffset)
            {
                for (var colOffset = 0; colOffset < array.Width; ++colOffset)
                {
                    var point = new XLSheetPoint(anchor.Row + rowOffset, anchor.Column + colOffset);
                    valueSlice.SetCellValue(point, array[rowOffset, colOffset].ToCellValue());
                }
            }
        }

        formula.Range = newFootprint;

        // Keep the spill-owner lookup in sync so a read of any spilled cell can force this
        // anchor to evaluate first during recalc.
        SetSpillFootprint(sheet.SheetId, formula, newFootprint);

        // Keep the dependency tree's area for this formula in sync with the footprint, so a
        // later change to the array's precedents invalidates dependents of every spilled cell
        // (not just the anchor). Only needed once the tree exists and the footprint changed.
        if (_dependencyTree is not null && newFootprint != previousRange)
        {
            var formulaArea = new XLBookArea(sheet.Name, newFootprint);
            _dependencyTree.UpdateSpillFootprint(formulaArea, formula, sheet.Workbook);
        }
    }

    /// <summary>
    /// Rebuilds <see cref="_spillOwners"/> from every dynamic-array formula's current footprint.
    /// Called whenever the dependency tree is (re)built so the lookup reflects spills produced in
    /// earlier sessions/recalcs or restored from a loaded file.
    /// </summary>
    private void RebuildSpillOwners(XLWorkbook wb)
    {
        _spillOwners.Clear();
        foreach (var sheet in wb.WorksheetsInternal)
        {
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(XLSheetRange.Full);
            while (enumerator.MoveNext())
            {
                var formula = enumerator.Current;
                if (formula.IsDynamicArray && formula.Range != default)
                    _spillOwners.Add(new SpillFootprint(sheet.SheetId, formula.Range, formula));
            }
        }
    }

    /// <summary>
    /// Records a dynamic-array formula's current spill footprint, replacing any prior footprint
    /// for the same formula (footprints never overlap, so one rectangle per formula is enough).
    /// </summary>
    private void SetSpillFootprint(uint sheetId, XLCellFormula formula, XLSheetRange footprint)
    {
        for (var i = 0; i < _spillOwners.Count; i++)
        {
            if (ReferenceEquals(_spillOwners[i].Owner, formula))
            {
                _spillOwners[i] = new SpillFootprint(sheetId, footprint, formula);
                return;
            }
        }

        _spillOwners.Add(new SpillFootprint(sheetId, footprint, formula));
    }

    /// <summary>
    /// If <paramref name="point"/> is a spilled (non-anchor) cell of a dynamic array whose anchor
    /// is dirty, returns the anchor point so the caller can force the anchor to evaluate first.
    /// The anchor cell itself holds a formula, so it never reaches this lookup.
    /// </summary>
    internal bool TryGetDirtySpillOwner(uint sheetId, XLSheetPoint point, XLWorkbook wb, out XLSheetPoint anchor)
    {
        foreach (var footprint in _spillOwners)
        {
            if (footprint.SheetId == sheetId && footprint.Range.Contains(point) && footprint.Owner.IsDirty(wb))
            {
                anchor = footprint.Range.FirstPoint;
                return true;
            }
        }

        anchor = default;
        return false;
    }

    /// <summary>
    /// Returns <c>true</c> if any cell of the prospective footprint blocks the spill.
    /// The anchor itself and any cell within <paramref name="ownedRange"/> (this formula's
    /// previous footprint, which will be overwritten) never block; any other cell holding a
    /// formula or a non-blank value does.
    /// </summary>
    private static bool HasSpillCollision(XLSheetPoint anchor, int lastRow, int lastColumn, XLSheetRange ownedRange, ValueSlice valueSlice, FormulaSlice formulaSlice)
    {
        for (var row = anchor.Row; row <= lastRow; ++row)
        {
            for (var column = anchor.Column; column <= lastColumn; ++column)
            {
                var point = new XLSheetPoint(row, column);
                if (point == anchor || ownedRange.Contains(point))
                    continue;

                if (formulaSlice.Get(point) is not null)
                    return true;

                if (!valueSlice.GetCellValue(point).IsBlank)
                    return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Blanks every cell of <paramref name="previousRange"/> that falls outside
    /// <paramref name="keepRange"/>. Used to erase a stale footprint when a dynamic array
    /// shrinks, moves, or collapses to a <c>#SPILL!</c> anchor. A <c>default</c> previous
    /// range (never spilled) clears nothing.
    /// </summary>
    private static void ClearSpillFootprint(XLSheetRange previousRange, XLSheetRange keepRange, ValueSlice valueSlice)
    {
        if (previousRange == default)
            return;

        foreach (var point in previousRange)
        {
            if (!keepRange.Contains(point))
                valueSlice.SetCellValue(point, Blank.Value);
        }
    }

    internal AnyValue EvaluateName(string nameFormula, XLWorksheet ws)
    {
        var ctx = new CalcContext(this, _culture, ws.Workbook, ws, null);
        return EvaluateFormula(nameFormula, ctx);
    }

    private AnyValue EvaluateFormula(string expression, CalcContext ctx)
    {
        var x = _cache[expression];

        var result = x.AstRoot.Accept(ctx, _visitor);
        return result;
    }

    // build/get static keyword table
    private static FunctionRegistry GetFunctionTable()
    {
        var fr = new FunctionRegistry();

        // register built-in functions (and constants)
        Engineering.Register(fr);
        Information.Register(fr);
        Logical.Register(fr);
        Lookup.Register(fr);
        MathTrig.Register(fr);
        Text.Register(fr);
        Statistical.Register(fr);
        DateAndTime.Register(fr);
        Financial.Register(fr);
        DynamicArray.Register(fr);

        return fr;
    }

    /// <summary>
    /// Convert any kind of formula value to value returned as a content of a cell.
    /// <list type="bullet">
    ///    <item><c>bool</c> - represents a logical value.</item>
    ///    <item><c>double</c> - represents a number and also date/time as serial date-time.</item>
    ///    <item><c>string</c> - represents a text value.</item>
    ///    <item><see cref="XLError" /> - represents a formula calculation error.</item>
    /// </list>
    /// </summary>
    private static ScalarValue ToCellContentValue(AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
            return scalar;

        if (collection.TryPickT0(out var array, out var reference))
        {
            return array[0, 0];
        }

        if (reference.TryGetSingleCellValue(out var cellValue, ctx))
            return cellValue;

        var intersected = reference.ImplicitIntersection(ctx.FormulaAddress);
        if (!intersected.TryPickT0(out var singleCellReference, out var error))
            return error;

        if (!singleCellReference.TryGetSingleCellValue(out var singleCellValue, ctx))
            throw new InvalidOperationException("Got multi cell reference instead of single cell reference.");

        return singleCellValue;
    }

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        if (_dependencyTree is not null)
            _dependencyTree.RenameSheet(oldSheetName, newSheetName);
    }
}

internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue> arg);
