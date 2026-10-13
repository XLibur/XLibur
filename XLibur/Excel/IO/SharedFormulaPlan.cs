using System;
using System.Collections.Generic;
using System.Globalization;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.IO;

/// <summary>
/// How a cell's formula is written: as a formula of its own, or as part of a shared formula.
/// </summary>
internal enum SharedFormulaRole
{
    /// <summary>Not part of a shared formula. The cell's <c>f</c> holds its own text.</summary>
    None,

    /// <summary>The first cell of a shared formula. Its <c>f</c> holds the text, the <c>ref</c> and the <c>si</c>.</summary>
    Master,

    /// <summary>A later cell of a shared formula. Its <c>f</c> holds only <c>t="shared"</c> and the <c>si</c>.</summary>
    Follower,
}

/// <summary>
/// The shared formulas that a save writes back for one sheet: each shared formula that a load brought
/// in, limited to the cells that still hold it (ADR 0003, #692).
/// </summary>
/// <remarks>
/// <para>
/// A cell keeps its <see cref="SharedFormulaGroup"/> only while its A1 text and its position are the
/// ones the loader gave it (<see cref="XLCellFormula.TryGetShared"/>). R1C1 text depends on nothing
/// else, so every cell that still has its group holds the group's formula, and no formula is parsed or
/// converted here. An edit, a move, a shift, a copy and a rewrite by a sheet rename each drop the
/// group, and those cells are written as formulas of their own.
/// </para>
/// <para>
/// The rules come from ECMA-376 Part 1 §18.3.1.40 and from what Excel opens, measured over COM:
/// </para>
/// <list type="bullet">
/// <item>The master must come before every other cell of the formula in the sheet, and every cell
/// must be inside the master's <c>ref</c>. Excel refuses a file that breaks either rule. The master
/// here is the first cell in row-major order, and a cell left of its column is written as a formula of
/// its own, so the master is also the top-left cell of <c>ref</c>, as Excel writes it.</item>
/// <item>A <c>ref</c> may hold cells that are not in the formula; those keep their own formula.
/// <c>si</c> is numbered from 0 in the order the masters are written, as Excel numbers it, although
/// Excel opens any numbers.</item>
/// <item>Excel derives each later cell from the master's text, moved by the offset between the two
/// cells, and wraps a reference that runs off the sheet. XLibur writes <c>#REF!</c> there instead, so
/// a cell whose text has <c>#REF!</c> is written as a formula of its own.</item>
/// <item>Excel refuses a shared formula that has a reference to a sheet (also 3D and external) or a
/// structured reference, and never writes one. A formula whose text has <c>!</c> or <c>[</c> is
/// written as a formula of its own in every cell.</item>
/// </list>
/// </remarks>
internal sealed class SharedFormulaPlan
{
    private readonly Dictionary<SharedFormulaGroup, Group> _groups;
    private int _nextIndex;

    private SharedFormulaPlan(Dictionary<SharedFormulaGroup, Group> groups)
    {
        _groups = groups;
    }

    /// <summary>
    /// Find the shared formulas of a sheet, or <c>null</c> when no cell still holds one.
    /// </summary>
    /// <remarks>
    /// One walk over the formulas of the sheet. It allocates only when a cell holds a shared formula:
    /// one dictionary for the sheet and one object for each shared formula.
    /// </remarks>
    internal static SharedFormulaPlan? Create(XLWorksheet worksheet)
    {
        Dictionary<SharedFormulaGroup, Group>? groups = null;
        using var enumerator = worksheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            var point = enumerator.Point;
            if (!TryGetCandidate(enumerator.Current, point, out var sharedGroup))
                continue;

            groups ??= new Dictionary<SharedFormulaGroup, Group>();
            if (groups.TryGetValue(sharedGroup, out var group))
                group.Add(point);
            else
                groups.Add(sharedGroup, new Group(point, IsShareable(enumerator.Current.A1)));
        }

        return groups is null ? null : new SharedFormulaPlan(groups);
    }

    /// <summary>
    /// How to write the formula of the cell at <paramref name="point"/>. Cells must be asked in the
    /// order they are written, because the master's call numbers the formula.
    /// </summary>
    /// <param name="formula">The formula of the cell.</param>
    /// <param name="point">The cell.</param>
    /// <param name="index">The <c>si</c> to write, when the cell is part of a shared formula.</param>
    /// <param name="range">The <c>ref</c> to write, when the cell is the master.</param>
    internal SharedFormulaRole GetRole(XLCellFormula formula, Point point, out string index, out Area range)
    {
        index = string.Empty;
        range = default;
        if (!TryGetCandidate(formula, point, out var sharedGroup)
            || !_groups.TryGetValue(sharedGroup, out var group)
            || !group.IsWritten)
            return SharedFormulaRole.None;

        if (point == group.Master)
        {
            group.Index = _nextIndex.ToString(CultureInfo.InvariantCulture);
            _nextIndex++;
            index = group.Index;
            range = group.Range;
            return SharedFormulaRole.Master;
        }

        // A cell outside ref, or one written before its master, would make Excel refuse the file.
        // Neither can happen to a cell the plan saw, but the check keeps it that way by construction.
        if (group.Index is null || !group.Range.Contains(point))
            return SharedFormulaRole.None;

        index = group.Index;
        return SharedFormulaRole.Follower;
    }

    /// <summary>
    /// Get the shared formula of a cell that may be written as part of it.
    /// </summary>
    private static bool TryGetCandidate(XLCellFormula formula, Point point, out SharedFormulaGroup sharedGroup)
    {
        sharedGroup = null!;
        return formula.Type == FormulaType.Normal
               && !formula.IsDynamicArray
               && formula.TryGetShared(point, out sharedGroup!)
               && !formula.A1.Contains("#REF!", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Whether Excel opens this text as a shared formula. Excel refuses a reference to a sheet, which
    /// a 3D and an external reference also have, and a structured reference. The test is on characters,
    /// so a string literal with either character is written as a formula of its own in every cell too.
    /// </summary>
    private static bool IsShareable(string formulaA1) => formulaA1.AsSpan().IndexOfAny('!', '[') < 0;

    /// <summary>
    /// One shared formula of the sheet: its master, the area its cells span, and its <c>si</c>
    /// once the master is written.
    /// </summary>
    private sealed class Group(Point master, bool shareable)
    {
        private int _lastRow = master.Row;
        private int _lastColumn = master.Column;
        private int _count = 1;

        internal Point Master { get; } = master;

        /// <summary>The <c>si</c>, set when the master is written.</summary>
        internal string? Index { get; set; }

        internal Area Range => new(Master, new Point(_lastRow, _lastColumn));

        /// <summary>
        /// Whether the formula is written as a shared formula. A formula of one cell is written as a
        /// formula of its own: sharing it saves nothing.
        /// </summary>
        internal bool IsWritten => shareable && _count > 1;

        internal void Add(Point point)
        {
            // Left of the master, the cell would be outside a ref whose top-left cell is the master.
            if (point.Column < Master.Column)
                return;

            _lastRow = Math.Max(_lastRow, point.Row);
            _lastColumn = Math.Max(_lastColumn, point.Column);
            _count++;
        }
    }
}
