using ClosedXML.Parser;

namespace XLibur.Excel.CalcEngine.Visitors;

/// <summary>
/// How far down and to the right a formula's <em>shiftable</em> references reach.
/// <para>
/// A row or column shift can only rewrite a reference whose extent reaches into the shifted region,
/// so a formula whose furthest reference stops short of the shift start cannot be edited by it. That
/// makes this extent a sound pre-filter for the shift pass, which otherwise parses every formula in
/// the workbook on every structural edit just to discover it had nothing to do.
/// </para>
/// </summary>
/// <remarks>
/// Only the reference kinds the shifter rewrites are measured — plain <c>A1</c> and single-sheet
/// <c>Sheet1!A1</c> — matching <c>XLCellFormulaShifter</c>'s collector exactly. 3D, bang, external,
/// defined-name and structured references are never rewritten, so they widen nothing.
/// <para>
/// The extent deliberately ignores which sheet a reference names. A formula on one sheet can refer to
/// the sheet being shifted, and resolving that here would mean caching a per-sheet answer that a sheet
/// rename invalidates. Taking the maximum across all sheets over-approximates, which only ever costs a
/// parse that turns out to be unnecessary — it can never skip a formula that needed shifting.
/// </para>
/// </remarks>
internal readonly struct FormulaExtent
{
    /// <summary>The largest row any shiftable reference reaches.</summary>
    internal int MaxRow { get; private init; }

    /// <summary>The largest column any shiftable reference reaches.</summary>
    internal int MaxColumn { get; private init; }

    /// <summary>
    /// Measures <paramref name="formulaA1"/>. A formula the parser rejects — an external workbook
    /// reference, say — reports the whole sheet, so it always reaches the shifter and its legacy
    /// fallback rather than being filtered out on a failed parse.
    /// </summary>
    internal static FormulaExtent Of(string formulaA1)
    {
        if (string.IsNullOrWhiteSpace(formulaA1))
            return new FormulaExtent { MaxRow = 0, MaxColumn = 0 };

        var collector = new ExtentCollector();
        try
        {
            FormulaParser<object?, object?, ExtentCollector>.CellFormulaA1(
                FormulaTransformation.ProtectStructuredRefColons(formulaA1, out _),
                collector,
                ExtentVisitor.Instance);
        }
        catch
        {
            return Unbounded;
        }

        return new FormulaExtent { MaxRow = collector.MaxRow, MaxColumn = collector.MaxColumn };
    }

    private static FormulaExtent Unbounded => new()
    {
        MaxRow = XLHelper.MaxRowNumber,
        MaxColumn = XLHelper.MaxColumnNumber,
    };

    /// <summary>
    /// Mutable accumulator threaded through the parse as the visitor's context.
    /// </summary>
    private sealed class ExtentCollector
    {
        internal int MaxRow { get; private set; }

        internal int MaxColumn { get; private set; }

        internal void Widen(ReferenceArea reference)
        {
            var first = reference.First;
            var second = reference.Second;

            // An axis the reference does not name spans the whole sheet: the rows of B:D are every
            // row, which is exactly why a whole-row shift reaches it.
            MaxRow = first.RowType == ReferenceAxisType.None
                ? XLHelper.MaxRowNumber
                : Max(MaxRow, first.RowValue, second.RowValue);

            MaxColumn = first.ColumnType == ReferenceAxisType.None
                ? XLHelper.MaxColumnNumber
                : Max(MaxColumn, first.ColumnValue, second.ColumnValue);
        }

        private static int Max(int current, int a, int b)
        {
            if (a > current)
                current = a;

            return b > current ? b : current;
        }
    }

    private sealed class ExtentVisitor : CollectVisitor<ExtentCollector>
    {
        internal static readonly ExtentVisitor Instance = new();

        public override object? Reference(ExtentCollector context, SymbolRange range, ReferenceArea reference)
        {
            context.Widen(reference);
            return null;
        }

        public override object? SheetReference(ExtentCollector context, SymbolRange range, string sheet,
            ReferenceArea reference)
        {
            context.Widen(reference);
            return null;
        }
    }
}
