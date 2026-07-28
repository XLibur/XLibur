using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Parser;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Repoints the references inside a formula when rows or columns are inserted or deleted.
/// <para>
/// References are located with <see cref="ClosedXML.Parser"/> rather than a regex. The parser knows
/// which spans of the formula are references and which are string literals, function names or
/// structured references, and it hands back each reference already decomposed into row/column values
/// and relative/absolute markers. That removes two whole classes of work the regex path needed: the
/// quote-parity scan that decided whether a match sat inside a string literal, and — far more
/// expensive — re-parsing every matched address by materialising a live <c>XLRange</c> through the
/// worksheet's range repository, once per reference, per formula, per shift.
/// </para>
/// <para>
/// Formulas the parser rejects (external workbook references such as <c>'[file.xlsx]Sheet'!A1</c>)
/// fall back to the regex implementation in <c>XLCellFormulaShifter.Legacy.cs</c>, so no formula that
/// shifted before stops shifting now.
/// </para>
/// </summary>
internal static partial class XLCellFormulaShifter
{
    internal static string ShiftFormulaRows(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
        int rowsShifted)
        => Shift(formulaA1, worksheetInAction, shiftedRange, rowsShifted, ShiftAxis.Row);

    internal static string ShiftFormulaColumns(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
        int columnsShifted)
        => Shift(formulaA1, worksheetInAction, shiftedRange, columnsShifted, ShiftAxis.Column);

    private static string Shift(string formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange, int shift,
        ShiftAxis axis)
    {
        if (string.IsNullOrWhiteSpace(formulaA1))
            return string.Empty;

        if (shift == 0 || !shiftedRange.RangeAddress.IsValid)
            return formulaA1;

        // Colons inside single-bracket structured reference column names would be read as range
        // operators. The placeholder is the same width as the colon it replaces, so every SymbolRange
        // the parser reports still indexes correctly into the original formula.
        var parseable = FormulaTransformation.ProtectStructuredRefColons(formulaA1, out var wasProtected);

        // Every worksheet's formulas are visited, not just the shifted sheet's, so that a formula on
        // one sheet referring to the shifted sheet is repointed too. Which references those are depends
        // on both names: a qualified reference names its sheet outright, while an unqualified one means
        // the sheet its formula lives on — and that only matches when the formula is *on* the sheet
        // being shifted.
        var plan = new ShiftPlan(
            shiftedSheetName: shiftedRange.Worksheet.Name,
            formulaSheetName: worksheetInAction.Name,
            parseable,
            shiftedRange,
            shift,
            axis);
        try
        {
            FormulaParser<object?, object?, ShiftPlan>.CellFormulaA1(parseable, plan, ShiftCollector.Instance);
        }
        catch (Exception)
        {
            return axis == ShiftAxis.Row
                ? ShiftFormulaRowsLegacy(formulaA1, worksheetInAction, shiftedRange, shift)
                : ShiftFormulaColumnsLegacy(formulaA1, worksheetInAction, shiftedRange, shift);
        }

        // The common case by a wide margin: the shift cannot reach anything this formula refers to.
        // Returning the original instance means an unaffected formula costs a parse and nothing else.
        if (!plan.HasEdits)
            return formulaA1;

        var shifted = plan.Apply(parseable);
        return wasProtected ? shifted.Replace(FormulaTransformation.ColonPlaceholder, ':') : shifted;
    }

    private enum ShiftAxis
    {
        Row,
        Column,
    }

    /// <summary>
    /// Collects the reference spans a shift has to rewrite. Only plain and single-sheet references are
    /// considered, matching what the regex path matched: a 3D reference (<c>Sheet1:Sheet3!A1</c>), a
    /// workbook-scoped <c>!A1</c> and any external-workbook reference are all left alone.
    /// </summary>
    private sealed class ShiftCollector : CollectVisitor<ShiftPlan>
    {
        internal static readonly ShiftCollector Instance = new();

        public override object? Reference(ShiftPlan plan, SymbolRange range, ReferenceArea reference)
        {
            plan.Visit(range, reference, sheet: null);
            return null;
        }

        public override object? SheetReference(ShiftPlan plan, SymbolRange range, string sheet, ReferenceArea reference)
        {
            plan.Visit(range, reference, sheet);
            return null;
        }
    }

    /// <summary>
    /// The shift being applied, plus the replacement spans collected for it.
    /// </summary>
    private sealed class ShiftPlan(
        string shiftedSheetName,
        string formulaSheetName,
        string formula,
        XLRange shiftedRange,
        int shift,
        ShiftAxis axis)
    {
        private const string RefErrorText = "#REF!";

        private readonly int _shiftFirstRow = shiftedRange.RangeAddress.FirstAddress.RowNumber;
        private readonly int _shiftLastRow = shiftedRange.RangeAddress.LastAddress.RowNumber;
        private readonly int _shiftFirstColumn = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
        private readonly int _shiftLastColumn = shiftedRange.RangeAddress.LastAddress.ColumnNumber;

        private List<Edit>? _edits;

        internal bool HasEdits => _edits is { Count: > 0 };

        internal void Visit(SymbolRange range, ReferenceArea reference, string? sheet)
        {
            // An unqualified reference resolves against the sheet its formula sits on.
            var referencedSheet = sheet ?? formulaSheetName;
            if (!XLHelper.SheetComparer.Equals(referencedSheet, shiftedSheetName))
                return;

            if (!TryShiftReference(reference, out var shifted, out var destroyed))
                return;

            _edits ??= [];
            _edits.Add(new Edit(range.Start, range.Length, Render(range, shifted, sheet, destroyed)));
        }

        /// <summary>
        /// Applies the collected replacements to <paramref name="formula"/>. The parser reports spans in
        /// source order and reference spans never nest, so a single forward pass suffices.
        /// </summary>
        internal string Apply(string formula)
        {
            var edits = _edits!;
            var builder = new StringBuilder(formula.Length + 8);
            var copiedTo = 0;

            foreach (var edit in edits)
            {
                builder.Append(formula, copiedTo, edit.Start - copiedTo);
                builder.Append(edit.Replacement);
                copiedTo = edit.Start + edit.Length;
            }

            builder.Append(formula, copiedTo, formula.Length - copiedTo);
            return builder.ToString();
        }

        /// <summary>
        /// Computes the reference's new extent on the shift axis.
        /// </summary>
        /// <returns><c>false</c> when the shift leaves the reference untouched, in which case the
        /// original text is kept verbatim rather than re-rendered.</returns>
        private bool TryShiftReference(ReferenceArea reference, out ReferenceArea shifted, out bool destroyed)
        {
            shifted = reference;
            destroyed = false;

            var first = reference.First;
            var second = reference.Second;

            // A reference confined to the other axis (3:5 for a column shift, B:D for a row shift) has
            // no extent to move. Excel leaves those alone and so did the regex path.
            if (IsConfinedToOtherAxis(first, second))
                return false;

            var (refFirst, refLast) = Extent(first, second, axis);
            var (crossFirst, crossLast) = Extent(first, second, Other(axis));
            var (shiftCrossFirst, shiftCrossLast) = axis == ShiftAxis.Row
                ? (_shiftFirstColumn, _shiftLastColumn)
                : (_shiftFirstRow, _shiftLastRow);

            // The shifted range must span the reference on the *other* axis, otherwise the insert or
            // delete cuts across it rather than moving it, and Excel leaves the reference where it is.
            if (crossFirst < shiftCrossFirst || crossLast > shiftCrossLast)
                return false;

            var shiftStart = axis == ShiftAxis.Row ? _shiftFirstRow : _shiftFirstColumn;
            if (refLast < shiftStart)
                return false;

            int newFirst, newLast;
            if (shift > 0)
            {
                newFirst = refFirst >= shiftStart ? refFirst + shift : refFirst;
                newLast = refLast + shift;
            }
            else
            {
                var deleted = -shift;
                var shiftEnd = shiftStart + deleted - 1;

                // Rows above the deletion keep their number; rows below move up by the deleted count;
                // rows inside it are gone. A boundary that lands inside the deleted block collapses onto
                // the edge of the block rather than being moved by the full count — clamping the bottom
                // is what stops a deletion that swallows a reference's tail from producing an inverted
                // range such as A2:A8 -> A2:A3 (row 4 survives; the answer is A2:A4).
                newFirst = refFirst < shiftStart ? refFirst : Math.Max(shiftStart, refFirst - deleted);
                newLast = refLast <= shiftEnd ? Math.Min(refLast, shiftStart - 1) : refLast - deleted;

                if (newLast < newFirst)
                {
                    destroyed = true;
                    return true;
                }
            }

            var max = axis == ShiftAxis.Row ? XLHelper.MaxRowNumber : XLHelper.MaxColumnNumber;
            newFirst = Math.Clamp(newFirst, 1, max);
            newLast = Math.Clamp(newLast, 1, max);

            if (newFirst == refFirst && newLast == refLast)
                return false;

            shifted = new ReferenceArea(
                WithExtent(first, newFirst, axis),
                WithExtent(second, newLast, axis));
            return true;
        }

        /// <summary>
        /// A reference that names only the axis we are not shifting: <c>3:5</c> during a column shift,
        /// <c>B:D</c> during a row shift. Its extent on the shift axis is the whole sheet, so it neither
        /// moves nor collapses.
        /// </summary>
        private bool IsConfinedToOtherAxis(RowCol first, RowCol second)
        {
            return axis == ShiftAxis.Row
                ? first.RowType == ReferenceAxisType.None && second.RowType == ReferenceAxisType.None
                : first.ColumnType == ReferenceAxisType.None && second.ColumnType == ReferenceAxisType.None;
        }

        /// <summary>
        /// The reference's span on <paramref name="on"/>. An axis the reference does not name (the rows
        /// of <c>B:D</c>) spans the whole sheet, which is what makes a whole-row insert reach it.
        /// </summary>
        private static (int First, int Last) Extent(RowCol first, RowCol second, ShiftAxis on)
        {
            if (on == ShiftAxis.Row)
            {
                return first.RowType == ReferenceAxisType.None
                    ? (1, XLHelper.MaxRowNumber)
                    : (first.RowValue, second.RowValue);
            }

            return first.ColumnType == ReferenceAxisType.None
                ? (1, XLHelper.MaxColumnNumber)
                : (first.ColumnValue, second.ColumnValue);
        }

        private static RowCol WithExtent(RowCol source, int value, ShiftAxis on)
        {
            return on == ShiftAxis.Row
                ? new RowCol(source.RowType, value, source.ColumnType, source.ColumnValue, source.Style)
                : new RowCol(source.RowType, source.RowValue, source.ColumnType, value, source.Style);
        }

        private static ShiftAxis Other(ShiftAxis value) => value == ShiftAxis.Row ? ShiftAxis.Column : ShiftAxis.Row;

        /// <summary>
        /// Re-renders a reference, restoring the sheet prefix the span covered and keeping the source's
        /// own shape.
        /// </summary>
        /// <remarks>
        /// The endpoints are joined by hand rather than through <see cref="ReferenceArea.GetDisplayStringA1"/>,
        /// which collapses a degenerate area to a single address: a range that a deletion narrows to one
        /// cell would come back as <c>A5</c> where the author wrote <c>A5:A10</c>, silently rewriting a
        /// range reference into a cell reference.
        /// </remarks>
        private string Render(SymbolRange range, ReferenceArea shifted, string? sheet, bool destroyed)
        {
            var body = destroyed
                ? RefErrorText
                : SourceIsRange(range)
                    ? string.Concat(shifted.First.GetDisplayStringA1(), ":", shifted.Second.GetDisplayStringA1())
                    : shifted.First.GetDisplayStringA1();

            if (sheet is null)
                return body;

            var builder = new StringBuilder(sheet.Length + body.Length + 3);
            builder.Append(sheet.EscapeSheetName());
            builder.Append('!');
            builder.Append(body);
            return builder.ToString();
        }

        /// <summary>
        /// Whether the source text wrote two addresses rather than one. Equal endpoints do not settle it
        /// — <c>A5:A5</c> is a legal one-cell range and round-trips as written — so only the presence of
        /// a range operator does. The scan starts past the sheet separator, since a quoted sheet name is
        /// allowed to contain a colon (<c>'a:b'!A5</c>).
        /// </summary>
        private bool SourceIsRange(SymbolRange range)
        {
            var text = formula.AsSpan(range.Start, range.Length);
            var separator = text.LastIndexOf('!');
            if (separator >= 0)
                text = text[(separator + 1)..];

            return text.Contains(':');
        }

        private readonly record struct Edit(int Start, int Length, string Replacement);
    }
}
