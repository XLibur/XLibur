using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using ClosedXML.Parser;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.CalcEngine.Functions;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

internal sealed class CalcContext
{
    private readonly bool _recursive;

    /// <summary>
    /// Per-evaluation cache for <see cref="GetCellValue"/>'s recursive branch. Lazily
    /// allocated on the first store. Caching is gated on the <c>_recursive</c> flag (set for
    /// <c>worksheet.Evaluate("...")</c>-style entry points) because that branch is the only
    /// one where a miss does real work — it allocates an <see cref="XLCell"/> and triggers a
    /// downstream formula recompute. The non-recursive path is already a couple of slice
    /// indexer reads, so caching there adds Dictionary overhead with no return on the
    /// canonical eval workloads (verified: ~150 MB allocation regression and no time
    /// improvement on <c>LoadAndReadAllCells</c> when the cache covered every read).
    /// </summary>
    private Dictionary<XLBookPoint, ScalarValue>? _recursiveCellValueCache;

    public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLCell cell)
        : this(calcEngine, culture, cell.Worksheet.Workbook, cell.Worksheet, cell.Address)
    {
    }

    public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLWorkbook? workbook, XLWorksheet? worksheet,
        IXLAddress? formulaAddress, bool recursive = false)
    {
        CalcEngine = calcEngine;
        Workbook = workbook;
        Worksheet = worksheet;
        FormulaAddress = formulaAddress;
        _recursive = recursive;
        Culture = culture;
    }

    // LEGACY: Remove once legacy functions are migrated
    internal XLCalcEngine CalcEngine => field ?? throw new MissingContextException();

    /// <summary>
    /// Worksheet of the cell the formula is calculating.
    /// </summary>
    public XLWorkbook Workbook => field ?? throw new MissingContextException();

    /// <summary>
    /// Worksheet of the cell the formula is calculating.
    /// </summary>
    public XLWorksheet Worksheet => field ?? throw new MissingContextException();

    /// <summary>
    /// Address of the calculated formula.
    /// </summary>
    public IXLAddress FormulaAddress => field ?? throw new MissingContextException();

    /// <summary>
    /// A culture used for comparisons and conversions (e.g. text to number).
    /// </summary>
    public CultureInfo Culture { get; }

    /// <summary>
    /// Excel 2016 and earlier doesn't support dynamic array formulas (it used an array formulas instead). As a consequence,
    /// all arguments for scalar functions where passed through implicit intersection before calling the function.
    /// </summary>
    public static bool UseImplicitIntersection => true;

    /// <summary>
    /// Should functions be calculated per item of multi-values argument in the scalar parameters.
    /// </summary>
    public bool IsArrayCalculation { get; set; }

    /// <summary>
    /// Sheet that is being recalculated. If set, formula can read dirty
    /// values from other sheets, but not from this sheetId.
    /// </summary>
    public uint? RecalculateSheetId { get; set; }

    internal XLSheetPoint FormulaSheetPoint => new(FormulaAddress.RowNumber, FormulaAddress.ColumnNumber);

    /// <summary>
    /// What date system should be used in calculation. Either 1900 or 1904.
    /// </summary>
    internal bool Use1904DateSystem { get; init; } = false;

    /// <summary>
    /// An upper limit (exclusive) of used calendar system.
    /// </summary>
    internal double DateSystemUpperLimit =>
        Use1904DateSystem ? XLHelper.Calendar1904UpperLimit : XLHelper.Calendar1900UpperLimit;

    private CancellationToken CancellationToken { get; init; } = CancellationToken.None;

    /// <summary>
    /// A helper method to check is user canceled the calculation in function loops.
    /// </summary>
    internal void ThrowIfCancelled()
    {
        CancellationToken.ThrowIfCancellationRequested();
    }

    internal ScalarValue GetCellValue(XLWorksheet? sheet, int rowNumber, int columnNumber)
    {
        sheet ??= Worksheet;
        var valueSlice = sheet.Internals.CellsCollection.ValueSlice;
        var point = new XLSheetPoint(rowNumber, columnNumber);
        var formula = sheet.Internals.CellsCollection.FormulaSlice.Get(point);

        if (formula is null)
            return valueSlice.GetCellValue(point);

        if (!formula.IsDirty)
            return valueSlice.GetCellValue(point);

        // Used when only one sheet should be recalculated, leaving other sheets with their data.
        if (RecalculateSheetId is not null && sheet.SheetId != RecalculateSheetId.Value)
            return valueSlice.GetCellValue(point);

        // A special branch for functions out of cells (e.g. worksheet.Evaluate("A1+A1*B1")).
        // These are not part of the calculation chain, so reordering a chain for them doesn't
        // make sense — instead the dirty formula is evaluated recursively. Caching here saves
        // a downstream formula recompute when the same cell appears more than once in the
        // expression, not just a slice read.
        if (_recursive)
        {
            var bookPoint = new XLBookPoint(sheet.SheetId, point);
            if (_recursiveCellValueCache is { } cache && cache.TryGetValue(bookPoint, out var cached))
                return cached;

            var cell = sheet.GetCell(point);
            var value = cell?.Value ?? Blank.Value;
            (_recursiveCellValueCache ??= new Dictionary<XLBookPoint, ScalarValue>()).Add(bookPoint, value);
            return value;
        }

        throw new GettingDataException(new XLBookPoint(sheet.SheetId, new XLSheetPoint(rowNumber, columnNumber)));
    }

    /// <summary>
    /// This method goes over slices and returns a value for each non-blank cell. Because it is using
    /// slice iterators, it scales with number of cells, not a size of area in reference (i.e., it works
    /// fine even if reference is <c>A1:XFD1048576</c>). It also works for 3D references.
    /// </summary>
    internal IEnumerable<ScalarValue> GetNonBlankValues(Reference reference)
    {
        foreach (var area in reference)
        {
            var sheet = area.Worksheet ?? Worksheet;
            var range = XLSheetRange.FromRangeAddress(area);

            // A value can be either in a non-empty value slice or an empty cell with a formula.
            var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
            while (enumerator.MoveNext())
            {
                var point = enumerator.Current;
                var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                if (!scalarValue.IsBlank)
                    yield return scalarValue;
            }
        }
    }

    /// <summary>
    /// Return all points in the <paramref name="areaReference" /> that satisfy the <paramref name="criteria" />.
    /// </summary>
    internal IEnumerable<XLSheetPoint> GetCriteriaPoints(XLRangeAddress areaReference, Criteria criteria)
    {
        var sheet = areaReference.Worksheet ?? Worksheet;
        var area = XLSheetRange.FromRangeAddress(areaReference);

        // This is a performance optimization when a user specifies a whole column
        // in the tally function (e.g. SUMIF(A:B, "5", C:D)).
        if (criteria.CanBlankValueMatch)
        {
            // Criteria can match blank cells, thus it's not possible to use optimized
            // used enumerators, and we have to check value of each cell.
            foreach (var point in area)
            {
                var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                if (criteria.Match(scalarValue))
                    yield return point;
            }
        }
        else
        {
            // The criteria can never match blank cells. That means we can skip all blank
            // cells entirely and use optimized used enumerators.
            var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(area);
            while (enumerator.MoveNext())
            {
                var point = enumerator.Current;
                var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                if (criteria.Match(scalarValue))
                    yield return point;
            }
        }
    }

    internal IEnumerable<ScalarValue> GetFilteredNonBlankValues(Reference reference, string function,
        bool skipHiddenRows = false)
    {
        // Allocate one per call, because visitor holds info whether function was found in a formula.
        var visitor = new FunctionVisitor(function);
        foreach (var area in reference)
        {
            var sheet = area.Worksheet ?? Worksheet;
            var range = XLSheetRange.FromRangeAddress(area);
            var hiddenRowTracker = new HiddenRowTracker(sheet);

            // A value can be either in a non-empty value slice or an empty cell with a formula.
            var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
            while (enumerator.MoveNext())
            {
                var point = enumerator.Current;

                if (skipHiddenRows && hiddenRowTracker.IsHidden(point.Row))
                    continue;

                if (CallsFunction(sheet.Internals.CellsCollection.FormulaSlice.Get(point), visitor))
                    continue;

                var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                if (!scalarValue.IsBlank)
                    yield return scalarValue;
            }
        }

        yield break;

        static bool CallsFunction(XLCellFormula? formula, FunctionVisitor visitor)
        {
            if (formula is null)
                return false;

            if (!formula.A1.Contains(visitor.FunctionName, StringComparison.OrdinalIgnoreCase))
                return false;

            FormulaParser<object?, object?, FunctionVisitor>.CellFormulaA1(formula.A1, visitor, visitor);
            if (!visitor.Found)
                return false;

            // To reuse same visitor without allocation, clear the found flag.
            visitor.Clear();
            return true;
        }
    }

    /// <summary>
    /// Tracks whether the current row is hidden, caching the result per row to avoid repeated lookups.
    /// </summary>
    private struct HiddenRowTracker(XLWorksheet sheet)
    {
        private int _currentRow;
        private bool _isHidden = true;

        internal bool IsHidden(int row)
        {
            if (_currentRow != row)
            {
                _currentRow = row;
                _isHidden = sheet.Internals.RowsCollection.TryGetValue(row, out var r) && r.IsHidden;
            }

            return _isHidden;
        }
    }

    /// <summary>
    /// This method should be used mostly for range arguments. If a value is scalar,
    /// return a single value enumerable.
    /// </summary>
    internal IEnumerable<ScalarValue> GetNonBlankValues(AnyValue value)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
        {
            if (scalar.IsBlank)
                return [];

            return new ScalarArray(scalar, 1, 1);
        }

        if (collection.TryPickT0(out var array, out var reference))
            return array.Where(x => !x.IsBlank);

        return GetNonBlankValues(reference);
    }

    internal IEnumerable<ScalarValue> GetAllValues(AnyValue value)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
            return new ScalarArray(scalar, 1, 1);

        if (collection.TryPickT0(out var array, out var reference))
            return array;

        return GetAllCellValues(reference);
    }

    private IEnumerable<ScalarValue> GetAllCellValues(Reference reference)
    {
        foreach (var area in reference)
        {
            var sheet = area.Worksheet;
            foreach (var point in XLSheetRange.FromRangeAddress(area))
            {
                yield return GetCellValue(sheet, point.Row, point.Column);
            }
        }
    }

    private sealed class FunctionVisitor : CollectVisitor<FunctionVisitor>
    {
        public FunctionVisitor(string function)
        {
            FunctionName = function;
        }

        internal string FunctionName { get; }

        public bool Found { get; private set; }

        public void Clear() => Found = false;

        public override object? Function(FunctionVisitor context, SymbolRange range, ReadOnlySpan<char> functionName,
            IReadOnlyList<object?> arguments)
        {
            Found = Found || functionName.Equals(FunctionName.AsSpan(), StringComparison.OrdinalIgnoreCase);
            return null;
        }
    }
}
