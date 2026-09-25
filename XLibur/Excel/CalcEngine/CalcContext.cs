using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using XLibur.Parser;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.CalcEngine.Functions;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

internal sealed class CalcContext : IStructuredReferenceScope
{
    private readonly bool _recursive;
    private readonly XLWorksheet? _worksheet;
    private readonly IXLAddress? _formulaAddress;

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
    private Dictionary<SheetPoint, ScalarValue>? _recursiveCellValueCache;

    public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLCell cell)
        : this(calcEngine, culture, cell.Worksheet.Workbook, cell.Worksheet, cell.Address)
    {
    }

    public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLWorkbook? workbook, XLWorksheet? worksheet,
        IXLAddress? formulaAddress, bool recursive = false)
    {
        CalcEngine = calcEngine;
        Workbook = workbook;
        _worksheet = worksheet;
        _formulaAddress = formulaAddress;
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
    public XLWorksheet Worksheet => _worksheet ?? throw new MissingContextException();

    /// <summary>
    /// Is the formula calculated on a sheet? <see cref="XLWorkbook.Evaluate"/> gives it none. Code
    /// that can do without the sheet asks this first; code that cannot reads <see cref="Worksheet"/>,
    /// which reports the missing context.
    /// </summary>
    internal bool HasWorksheet => _worksheet is not null;

    /// <summary>
    /// Address of the calculated formula.
    /// </summary>
    public IXLAddress FormulaAddress => _formulaAddress ?? throw new MissingContextException();

    /// <summary>
    /// The context a defined name's formula is evaluated in, when the formula this context is
    /// calculating refers to the name.
    /// </summary>
    /// <remarks>
    /// <para>
    /// It keeps this context's sheet and cell (D60), and its choice to calculate a dirty precedent
    /// first, so the engine's pending signal cannot escape a public <c>Evaluate</c> through a name.
    /// Either may be missing: <see cref="XLWorkbook.Evaluate"/> has neither, and the name then fails
    /// only if its own formula asks for one (#491).
    /// It also keeps the sheet a sheet-only recalculation is limited to, so the name reads another
    /// sheet's cells as they stand, exactly as the same formula typed into the cell does. Without it,
    /// a name reading a dirty cell on another sheet asked the chain for a cell the pass then skipped,
    /// and the pass never ended.
    /// </para>
    /// <para>
    /// The rest starts afresh, as it always has: the name is not an array formula because its caller
    /// is, and <see cref="IntersectOperands"/> is off,
    /// so an operator at the top of the name's formula keeps its range operand whole. That is what
    /// Excel does: the owner checked it on 2026-09-14, and a name holding <c>Sheet1!$A$1:$A$3+10</c>
    /// read in row 2 gives 11, with Excel adding <c>@</c>. <c>EvaluationOutcomeTests</c> pins it.
    /// </para>
    /// <para>
    /// A name met again while it is still being evaluated is a circular reference: with the same
    /// sheet and cell, its formula would lead back to itself for ever, and the stack would overflow.
    /// The names being evaluated are the chain of contexts this method built, so the check follows
    /// one formula's names only. A cell a name reads is calculated in a context of its own, and may
    /// use the same name for its own cell without that being a cycle.
    /// </para>
    /// </remarks>
    /// <param name="nameFormula">The formula of the name the formula of this context refers to.</param>
    /// <exception cref="XLCircularReferenceException">
    /// The name is already being evaluated, further out in this chain.
    /// </exception>
    internal CalcContext ForDefinedName(string nameFormula)
    {
        for (var outer = this; outer is not null; outer = outer.NameCaller)
        {
            if (string.Equals(outer.NameFormula, nameFormula, StringComparison.Ordinal))
                throw new XLCircularReferenceException($"A defined name whose formula is '{nameFormula}' depends on its own value.");
        }

        return new(CalcEngine, Culture, Workbook, _worksheet, _formulaAddress, _recursive)
        {
            RecalculateSheetId = RecalculateSheetId,
            NameFormula = nameFormula,
            NameCaller = this,
        };
    }

    /// <summary>
    /// The formula of the defined name this context evaluates, or <c>null</c> if it evaluates
    /// something else.
    /// </summary>
    private string? NameFormula { get; init; }

    /// <summary>
    /// The context whose formula refers to the name this context evaluates, or <c>null</c> if this
    /// context evaluates no name.
    /// </summary>
    private CalcContext? NameCaller { get; init; }

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
    /// <para>
    /// Should a reference operand of an operator be passed through implicit intersection before the
    /// operator sees it? Excel does this for a legacy formula: <c>=A1+B1:B3</c> in <c>C3</c> is
    /// stored as <c>=A1+@B1:B3</c> and answers <c>42 + B3</c>, not <c>42 + B1</c> (D38).
    /// </para>
    /// <para>
    /// This is the per-position flag the formula-level <see cref="IsArrayCalculation"/> cannot be.
    /// <see cref="CalculationVisitor"/> owns it: it is set once for the whole formula and cleared
    /// while a function's arguments are evaluated, because the enclosing function supplies the
    /// context there and several range-accepting functions need the array intact.
    /// </para>
    /// </summary>
    internal bool IntersectOperands { get; set; }

    /// <summary>
    /// Sheet that is being recalculated. If set, formula can read dirty
    /// values from other sheets, but not from this sheetId.
    /// </summary>
    public uint? RecalculateSheetId { get; set; }

    internal Point FormulaSheetPoint => new(FormulaAddress.RowNumber, FormulaAddress.ColumnNumber);

    /// <inheritdoc />
    /// <remarks>
    /// Both this and <see cref="IStructuredReferenceScope.Worksheet"/> throw
    /// <see cref="MissingContextException"/> when the expression has no anchoring cell. That is
    /// deliberate and pre-existing: the resolver reads them only for the forms that genuinely
    /// need a formula location, so <c>Table1[Amount]</c> still resolves from a bare
    /// <c>Evaluate</c> call while <c>[@Amount]</c> cannot.
    /// </remarks>
    Point IStructuredReferenceScope.FormulaPoint => FormulaSheetPoint;

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
        var point = new Point(rowNumber, columnNumber);
        var formula = sheet.Internals.CellsCollection.FormulaSlice.Get(point);

        if (formula is null)
            return GetFormulaLessCellValue(sheet, valueSlice, point);

        // Used when only one sheet should be recalculated, leaving other sheets with their data.
        if (formula.IsClean() || IsOutsideRecalculatedSheet(sheet))
            return valueSlice.GetCellValue(point);

        // A special branch for functions out of cells (e.g. worksheet.Evaluate("A1+A1*B1")).
        // These are not part of the calculation chain, so reordering a chain for them doesn't
        // make sense — instead the dirty formula is evaluated recursively.
        if (_recursive)
            return GetRecursiveCellValue(sheet, point);

        throw new GettingDataException(new SheetPoint(sheet.SheetId, new Point(rowNumber, columnNumber)));
    }

    /// <summary>
    /// Whether <paramref name="sheet"/> is left with its data because only another sheet is being recalculated.
    /// </summary>
    private bool IsOutsideRecalculatedSheet(XLWorksheet sheet)
        => RecalculateSheetId is not null && sheet.SheetId != RecalculateSheetId.Value;

    /// <summary>
    /// A formula-less cell may still be a spilled cell of a dynamic array. If its owning anchor is
    /// dirty, the stored value is stale — force the anchor to evaluate first.
    /// </summary>
    private ScalarValue GetFormulaLessCellValue(XLWorksheet sheet, ValueSlice valueSlice, Point point)
    {
        if (!CalcEngine.HasSpillOwners ||
            !CalcEngine.TryGetDirtySpillOwner(sheet.SheetId, point, out var spillAnchor))
            return valueSlice.GetCellValue(point);

        if (IsOutsideRecalculatedSheet(sheet))
            return valueSlice.GetCellValue(point);

        if (_recursive)
        {
            // Evaluate the anchor recursively so it spills current values into this cell.
            _ = GetCellValue(sheet, spillAnchor.Row, spillAnchor.Column);
            return valueSlice.GetCellValue(point);
        }

        throw new GettingDataException(new SheetPoint(sheet.SheetId, spillAnchor));
    }

    /// <summary>
    /// Evaluates a dirty formula recursively. Caching here saves a downstream formula recompute when
    /// the same cell appears more than once in the expression, not just a slice read.
    /// </summary>
    private ScalarValue GetRecursiveCellValue(XLWorksheet sheet, Point point)
    {
        var bookPoint = new SheetPoint(sheet.SheetId, point);
        if (_recursiveCellValueCache is { } cache && cache.TryGetValue(bookPoint, out var cached))
            return cached;

        var cell = sheet.GetCell(point);
        var value = cell?.Value ?? Blank.Value;
        (_recursiveCellValueCache ??= new Dictionary<SheetPoint, ScalarValue>()).Add(bookPoint, value);
        return value;
    }

    /// <summary>
    /// This method goes over slices and returns a value for each non-blank cell. Because it is using
    /// slice iterators, it scales with number of cells, not a size of area in reference (i.e., it works
    /// fine even if reference is <c>A1:XFD1048576</c>). It also works for 3D references.
    /// </summary>
    /// <remarks>
    /// This is the one way the calc engine reads the values of a reference. Areas are read in
    /// their order in the reference and each area in row-major order (left to right, then top to
    /// bottom), the order functions such as NPV, IRR and MIRR depend on. A cell covered by two
    /// overlapping areas is read once for each.
    /// </remarks>
    internal IEnumerable<ScalarValue> GetNonBlankValues(Reference reference)
    {
        foreach (var area in reference)
        {
            var sheet = area.Worksheet ?? Worksheet;
            var range = Area.FromRangeAddress(area);

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

    /// <summary>
    /// Return all points in the <paramref name="areaReference" /> that satisfy the <paramref name="criteria" />.
    /// </summary>
    internal IEnumerable<Point> GetCriteriaPoints(XLRangeAddress areaReference, Criteria criteria)
    {
        // This is a performance optimization when a user specifies a whole column
        // in the tally function (e.g. SUMIF(A:B, "5", C:D)).
        return criteria.CanBlankValueMatch
            ? GetCriteriaPointsOfEveryCell(areaReference, criteria)
            : GetCriteriaPointsOfUsedCells(areaReference, criteria);
    }

    /// <summary>
    /// Criteria can match blank cells, thus it's not possible to use optimized
    /// used enumerators, and we have to check value of each cell.
    /// </summary>
    private IEnumerable<Point> GetCriteriaPointsOfEveryCell(XLRangeAddress areaReference, Criteria criteria)
    {
        var sheet = areaReference.Worksheet ?? Worksheet;
        var area = Area.FromRangeAddress(areaReference);

        foreach (var point in area)
        {
            var scalarValue = GetCellValue(sheet, point.Row, point.Column);
            if (criteria.Match(scalarValue))
                yield return point;
        }
    }

    /// <summary>
    /// The criteria can never match blank cells. That means we can skip all blank
    /// cells entirely and use optimized used enumerators.
    /// </summary>
    private IEnumerable<Point> GetCriteriaPointsOfUsedCells(XLRangeAddress areaReference, Criteria criteria)
    {
        var sheet = areaReference.Worksheet ?? Worksheet;
        var area = Area.FromRangeAddress(areaReference);

        var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(area);
        while (enumerator.MoveNext())
        {
            var point = enumerator.Current;
            var scalarValue = GetCellValue(sheet, point.Row, point.Column);
            if (criteria.Match(scalarValue))
                yield return point;
        }
    }

    /// <summary>
    /// Values of a reference, skipping blanks and any cell whose own formula calls one of
    /// <paramref name="functions"/> — which is how SUBTOTAL and AGGREGATE avoid counting a nested
    /// subtotal twice. AGGREGATE has to skip both function names, so more than one may be given.
    /// </summary>
    internal IEnumerable<ScalarValue> GetFilteredNonBlankValues(Reference reference, string[] functions,
        bool skipHiddenRows = false)
    {
        // Allocate one per call, because visitor holds info whether function was found in a formula.
        var visitor = new FunctionVisitor(functions);
        foreach (var area in reference)
        {
            var sheet = area.Worksheet ?? Worksheet;
            var range = Area.FromRangeAddress(area);
            var hiddenRowTracker = new HiddenRowTracker(sheet);

            // A value can be either in a non-empty value slice or an empty cell with a formula.
            var enumerator = sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
            while (enumerator.MoveNext())
            {
                var point = enumerator.Current;

                if (IsFilteredOut(sheet, point, skipHiddenRows, ref hiddenRowTracker, visitor))
                    continue;

                var scalarValue = GetCellValue(sheet, point.Row, point.Column);
                if (!scalarValue.IsBlank)
                    yield return scalarValue;
            }
        }

        yield break;
    }

    /// <summary>
    /// The nesting check that <see cref="GetFilteredNonBlankValues"/> applies to each cell, for one
    /// formula on its own: would a SUBTOTAL or AGGREGATE over a cell holding
    /// <paramref name="formulaA1"/> skip that cell?
    /// </summary>
    internal static bool IsSkippedByNestingCheck(string formulaA1, string[] functions)
        => CallsFunction(XLCellFormula.NormalA1(formulaA1), new FunctionVisitor(functions));

    /// <summary>
    /// Whether a cell is left out of <see cref="GetFilteredNonBlankValues"/>: its row is hidden and hidden
    /// rows are skipped, or its own formula calls one of the filtered functions.
    /// </summary>
    private static bool IsFilteredOut(XLWorksheet sheet, Point point, bool skipHiddenRows,
        ref HiddenRowTracker hiddenRowTracker, FunctionVisitor visitor)
    {
        if (skipHiddenRows && hiddenRowTracker.IsHidden(point.Row))
            return true;

        return CallsFunction(sheet.Internals.CellsCollection.FormulaSlice.Get(point), visitor);
    }

    private static bool CallsFunction(XLCellFormula? formula, FunctionVisitor visitor)
    {
        if (formula is null)
            return false;

        if (!visitor.MightBeCalledBy(formula.A1))
            return false;

        // A refused formula does not call SUBTOTAL as far as anyone can tell, so its cell counts; the
        // text is not searched for the name instead. The parse may have stopped after it saw a call,
        // so the flag is cleared on this path too.
        if (!FormulaText.TryWalk(formula.A1, visitor, visitor, FormulaNotation.A1, out _, out _))
        {
            visitor.Clear();
            return false;
        }

        if (!visitor.Found)
            return false;

        // To reuse same visitor without allocation, clear the found flag.
        visitor.Clear();
        return true;
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
            foreach (var point in Area.FromRangeAddress(area))
            {
                yield return GetCellValue(sheet, point.Row, point.Column);
            }
        }
    }

    private sealed class FunctionVisitor : CollectVisitor<FunctionVisitor>
    {
        private readonly string[] _functionNames;

        public FunctionVisitor(string[] functionNames)
        {
            _functionNames = functionNames;
        }

        public bool Found { get; private set; }

        public void Clear() => Found = false;

        /// <summary>
        /// A cheap substring test that rules out most formulas before the parser is involved. A
        /// false positive only costs a parse; a false negative would be wrong, so this must stay
        /// at least as permissive as <see cref="Function"/>.
        /// </summary>
        public bool MightBeCalledBy(string formula)
        {
            foreach (var name in _functionNames)
            {
                if (formula.Contains(name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        public override object? Function(FunctionVisitor context, SymbolRange range, ReadOnlySpan<char> functionName,
            IReadOnlyList<object?> arguments)
        {
            var bareName = StripNameSpace(functionName);
            foreach (var name in _functionNames)
                Found = Found || bareName.Equals(name.AsSpan(), StringComparison.OrdinalIgnoreCase);

            return null;
        }

        /// <summary>
        /// Drop the namespace a post-2007 function is stored under. AGGREGATE is one of them, so a
        /// cell holding it reads back as <c>_xlfn.AGGREGATE(…)</c> and would not match its own name.
        /// </summary>
        private static ReadOnlySpan<char> StripNameSpace(ReadOnlySpan<char> functionName)
        {
            const string futureNameSpace = "_xlfn.";
            const string worksheetNameSpace = "_xlws.";

            if (functionName.StartsWith(futureNameSpace, StringComparison.OrdinalIgnoreCase))
                functionName = functionName[futureNameSpace.Length..];

            if (functionName.StartsWith(worksheetNameSpace, StringComparison.OrdinalIgnoreCase))
                functionName = functionName[worksheetNameSpace.Length..];

            return functionName;
        }
    }
}
