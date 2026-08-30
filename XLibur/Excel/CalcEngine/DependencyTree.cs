using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using RBush;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// <para>
/// A dependency tree structure to hold all formulas of the workbook and reference
/// objects they depend on. The key feature of dependency tree is to propagate
/// dirty flag across formulas.
/// </para>
/// <para>
/// When a data in a cell changes, all formulas that depend on it should be marked
/// as dirty, but it is hard to find which cells are affected - that is what
/// dependency tree does.
/// </para>
/// <para>
/// Dependency tree must be updated, when structure of a workbook is updated:
/// <list type="bullet">
///   <item>Sheet is added, renamed or deleted.</item>
///   <item>Name is added or deleted.</item>
///   <item>Table is resized, renamed, added or deleted.</item>
/// </list>
/// Any such action changes what cells formula depends on and
/// the formula dependencies must be updated.
/// </para>
/// </summary>
internal sealed class DependencyTree
{
    /// <summary>
    /// The source of the truth, a storage of formula dependencies. The dependency tree is
    /// constructed from this collection.
    /// </summary>
    private readonly Dictionary<XLCellFormula, FormulaDependencies> _dependencies = new();

    /// <summary>
    /// Visitor to extract precedents of formulas.
    /// </summary>
    private readonly DependenciesVisitor _visitor;

    /// <summary>
    /// A dependency tree for each sheet (key is sheet name).
    /// </summary>
    private readonly Dictionary<string, SheetDependencyTree> _sheetTrees = new(XLHelper.SheetComparer);

    public DependencyTree()
    {
        _visitor = new DependenciesVisitor();
    }

    internal bool IsEmpty => _sheetTrees.All(sheetTree => sheetTree.Value.IsEmpty) && _dependencies.Count == 0;

#pragma warning disable S3776 // One branch per formula kind, each documented; splitting would separate them from the walk
    internal static DependencyTree CreateFrom(XLWorkbook workbook)
    {
        var tree = new DependencyTree();

        // Add tree before adding formulas, because formula can reference any sheet.
        foreach (var sheet in workbook.WorksheetsInternal)
            tree.AddSheetTree(sheet);

        foreach (var sheet in workbook.WorksheetsInternal)
        {
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
            {
                var formula = enumerator.Current;
                var point = enumerator.Point;
                if (formula.IsDynamicArray)
                {
                    // A dynamic-array formula lives only in its anchor cell (spilled cells are
                    // formula-less), so it appears exactly once. Register the whole spill
                    // footprint so a change to the array's precedents invalidates dependents of
                    // ANY spilled cell, not just the anchor. Before the first spill the footprint
                    // is unknown (default) — register the 1x1 anchor; the spill re-registers the
                    // formula once its size is known (see XLCalcEngine.SpillDynamicArray).
                    var footprint = formula.Range == default ? new Area(point, point) : formula.Range;
                    var bookArea = new SheetArea(sheet.Name, footprint);
                    tree.AddFormula(bookArea, formula, workbook);
                }
                else if (formula.Type == FormulaType.Normal)
                {
                    var bookArea = new SheetArea(sheet.Name, new Area(point, point));
                    tree.AddFormula(bookArea, formula, workbook);
                }
                else if (formula.Type == FormulaType.Array)
                {
                    // Ignore all non-master cells
                    var isMasterCell = formula.Range.FirstPoint == point;
                    if (isMasterCell)
                    {
                        var bookArea = new SheetArea(sheet.Name, formula.Range);
                        tree.AddFormula(bookArea, formula, workbook);
                    }
                }
                // Data-table formulas are skipped deliberately, and cannot simply be added to
                // the chain above. AddFormula derives precedents by parsing the formula text,
                // and a data table's text is the placeholder "{TABLE(A1,}" — not valid formula
                // syntax, so parsing it throws ExpressionParseException. Registering them needs
                // precedents built from Input1/Input2 and the table's header formulas instead of
                // from an AST. XLibur does not evaluate data tables either (there is no TABLE
                // function), so the only gain would be dropping the full-recalculation trigger in
                // XLCalcEngine.TryEvaluateSingleCell.
                // FormulaType.Shared is never produced, so it needs no handling here.
            }
        }

        return tree;
    }
#pragma warning restore S3776

    /// <summary>
    /// Add a formula to the dependency tree.
    /// </summary>
    /// <param name="formulaArea">Area of a formula, for normal cells 1x1, for array can be larger.</param>
    /// <param name="formula">The cell formula.</param>
    /// <param name="workbook">Workbook that is used to find precedents (names ect.).</param>
    /// <returns>Added cell formula dependencies.</returns>
    /// <exception cref="ArgumentException">Formula already is in the tree.</exception>
    internal FormulaDependencies AddFormula(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook)
    {
        var precedents = GetFormulaPrecedents(formulaArea, formula, workbook);

        _dependencies.Add(formula, precedents);

        foreach (var precedentArea in precedents.Areas)
        {
            // Add dependency to its sheet dependency tree. The formula might contain
            // a dependency for a sheet that doesn't exist in a workbook. Such dependencies
            // are ignored, until sheet is added.
            if (_sheetTrees.TryGetValue(precedentArea.Name, out var sheetTree))
            {
                // Dependent worksheet exists
                var dependent = new Dependent(formulaArea, formula);
                sheetTree.AddDependent(precedentArea.Area, dependent);
            }
        }

        return precedents;
    }

    /// <summary>
    /// Re-register a dynamic-array formula under a new spill footprint. Called after a spill
    /// grows, shrinks, or collapses to a <c>#SPILL!</c> anchor so the tree keeps invalidating
    /// dependents of every currently-spilled cell. Safe to call whether or not the formula is
    /// already in the tree.
    /// </summary>
    internal void UpdateSpillFootprint(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook)
    {
        RemoveFormula(formula);
        AddFormula(formulaArea, formula, workbook);
    }

    /// <summary>
    /// Remove formula from the dependency tree.
    /// </summary>
    /// <param name="formula">Formula to remove.</param>
    internal void RemoveFormula(XLCellFormula formula)
    {
        if (!_dependencies.Remove(formula, out var dependencies))
            return;

        foreach (var precedentArea in dependencies.Areas)
        {
            if (!_sheetTrees.TryGetValue(precedentArea.Name, out var sheetTree))
                throw new InvalidOperationException($"Dependency tree for sheet '{precedentArea.Name}' not found.");

            sheetTree.RemoveDependent(precedentArea.Area, formula);
        }
    }

    internal void AddSheetTree(IXLWorksheet sheet)
    {
        _sheetTrees.Add(sheet.Name, new SheetDependencyTree());
    }

    internal void RenameSheet(string oldSheetName, string newSheetName)
    {
        foreach (var formulaDependencies in _dependencies.Values)
            formulaDependencies.RenameSheet(oldSheetName, newSheetName);

        var renamedSheetTree = _sheetTrees[oldSheetName];
        _sheetTrees.Remove(oldSheetName);
        _sheetTrees.Add(newSheetName, renamedSheetTree);

        foreach (var sheetTree in _sheetTrees.Values)
            sheetTree.RenameSheet(oldSheetName, newSheetName);
    }

    /// <summary>
    /// <para>
    /// Monotonically increasing walk id, handed out one per <see cref="MarkDirty"/> call and
    /// stamped onto each formula the walk enqueues (see <see cref="XLCellFormula.TryVisit"/>).
    /// Distinguishing "enqueued by walk N" from "dirty for any other reason" this way costs one
    /// field compare-and-set per node instead of a collection allocated per call; a HashSet-based
    /// visited set was measured first and cost roughly 7x the allocation and 3x the wall time on
    /// a bulk-edit workload with real dependents (see XLibur.Benchmarks.BulkEditDirtyWalkProfile,
    /// "bulkedit" profile mode) before this replaced it.
    /// </para>
    /// <para>
    /// The counter is process-wide rather than per-tree because the stamps outlive any single
    /// tree: <c>XLCalcEngine.Purge</c> discards and rebuilds the whole dependency tree on a sheet
    /// add or rename and on every row/column insert or delete, but the <see cref="XLCellFormula"/>
    /// objects holding the stamps are not recreated. A per-tree counter would restart at zero and
    /// hand a surviving formula an id it is already stamped with, so the first walk after each
    /// rebuild would prune at its first hop. One interlocked increment per walk — not per node —
    /// buys ids that are never reused for the life of the process.
    /// </para>
    /// </summary>
    private static long _walkGeneration;

    /// <summary>
    /// Queue reused across <see cref="MarkDirty"/> calls. The walk never re-enters itself (it only
    /// reads the sheet trees and sets a flag per formula), so one queue per tree is enough, and it
    /// keeps a bulk edit from allocating and regrowing a queue per written cell. Like the rest of
    /// the tree — and the workbook it belongs to — this assumes a single thread at a time.
    /// </summary>
    private readonly Queue<SheetArea> _walkQueue = new();

    /// <summary>
    /// Mark all formulas that depend (directly or transitively) on the area as dirty.
    /// </summary>
    /// <remarks>
    /// The walk tracks which formulas it has already enqueued itself, instead of asking whether a
    /// formula is already dirty. A formula can be dirty for reasons that have nothing to do with
    /// this walk — <see cref="XLCellFormula.MarkExplicitlyDirty"/> is also called by the public
    /// <c>InvalidateFormula</c>, a sheet rename, a reference shift and a range move — and treating
    /// "already dirty" as "already visited" stopped the walk at such a node and pruned everything
    /// downstream of it. Marking stays idempotent only for a node this same walk has already
    /// enqueued; a node dirtied by anything else is still traversed.
    /// </remarks>
    internal void MarkDirty(SheetArea dirtyArea)
    {
        var walkId = Interlocked.Increment(ref _walkGeneration);

        // BFS vs DFS: Although the longest chain found in the wild is 1000
        // formulas long, attacker could supply malicious excel with recursion
        // leading to stack overflow => use queue even with extra allocation cost.
        var queue = _walkQueue;
        queue.Clear();
        queue.Enqueue(dirtyArea);
        while (queue.Count > 0)
        {
            var affectedArea = queue.Dequeue();
            var sheetTree = _sheetTrees[affectedArea.Name];
            foreach (var area in sheetTree.FindDependentsAreas(affectedArea.Area))
            {
                foreach (var dependent in area.Dependents)
                {
                    // Ensure we don't end up in an infinite cycle: a formula already enqueued by
                    // this walk is not enqueued again, regardless of its dirty state.
                    if (!dependent.Formula.TryVisit(walkId))
                        continue;

                    dependent.MarkDirty();
                    queue.Enqueue(dependent.FormulaArea);
                }
            }
        }
    }

    private FormulaDependencies GetFormulaPrecedents(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook)
    {
        var ast = formula.GetAst(workbook.CalcEngine);
        var context = new DependenciesContext(formulaArea, workbook);
        var rootReference = ast.AstRoot.Accept(context, _visitor);

        // If formula references are propagated to the root, make sure to add them.
        if (rootReference is not null)
            context.AddAreas(rootReference);

        return context.Dependencies;
    }

    /// <summary>
    /// An area that is referred by formulas in different cells, i.e. it
    /// contains precedent cells for a formula. If anything in the area
    /// potentially changes, all dependents might also change.
    /// </summary>
    private sealed class AreaDependents : ISpatialData
    {
        /// <summary>
        /// An area in a sheet that is used by formulas, converted to RBush envelope.
        /// All RBush <c>double</c> coordinates are whole numbers.
        /// </summary>
        private readonly Envelope _area;

        private readonly List<Dependent> _dependents;

        internal AreaDependents(in Envelope area, Dependent firstDependent)
        {
            _area = area;
            _dependents = [firstDependent];
        }

        /// <summary>
        /// The area in a sheet on which some formulas depend on.
        /// </summary>
        /// <example><c>SIN(A4)</c> depends on <c>A4:A4</c> area.</example>.
        public ref readonly Envelope Envelope => ref _area;

        /// <summary>
        /// List of formulas that depend on the range, always at least one.
        /// </summary>
        internal List<Dependent> Dependents => _dependents;

        internal void AddDependent(Dependent dependent)
        {
            _dependents.Add(dependent);
        }

        internal void RemoveDependent(XLCellFormula formula)
        {
            for (var i = 0; i < _dependents.Count; ++i)
            {
                var dependent = _dependents[i];

                // several different formulas can depend on same area,
                // remove only dependent of the formula.
                if (dependent.Formula != formula)
                    continue;

                // Remove from list by moving the last element to the removed
                // element place and decrease capacity.
                _dependents[i] = _dependents[^1];

                // Remove last item, capacity is unchanged, only list size is updated.
                _dependents.RemoveAt(_dependents.Count - 1);
            }
        }

        internal void RenameSheet(string oldSheetName, string newSheetName)
        {
            for (var i = 0; i < _dependents.Count; ++i)
            {
                var dependent = _dependents[i];
                if (XLHelper.SheetComparer.Equals(dependent.FormulaArea.Name, oldSheetName))
                {
                    var renamedArea = new SheetArea(newSheetName, dependent.FormulaArea.Area);
                    _dependents[i] = new Dependent(renamedArea, dependent.Formula);
                }
            }
        }
    }

    /// <summary>
    /// A dependent on a precedent area. If the precedent area changes,
    /// the dependent might also now be invalid.
    /// </summary>
    private readonly struct Dependent
    {
        /// <summary>
        /// Area that is invalidated, when precedent area is marked as
        /// dirty. Generally, it is an area of formula (1x1 for normal
        /// formulas), larger for array formulas. Cell formula by itself
        /// doesn't contain it's address to make it easier add/delete
        /// rows/cols.
        /// </summary>
        internal readonly SheetArea FormulaArea;

        internal Dependent(SheetArea formulaArea, XLCellFormula formula)
        {
            FormulaArea = formulaArea;
            Formula = formula;
        }

        /// <summary>
        /// The formula that is affected by changes in precedent area.
        /// </summary>
        internal XLCellFormula Formula { get; }

        internal void MarkDirty() => Formula.MarkExplicitlyDirty();
    }

    /// <summary>
    /// A dependency tree for a single worksheet.
    /// </summary>
    private sealed class SheetDependencyTree
    {
        /// <summary>
        /// The precedent areas are not duplicated, though two areas might overlap.
        /// </summary>
        private readonly RBush<AreaDependents> _tree;

        /// <summary>
        /// All precedent areas in the sheet for all formulas in the workbook.
        /// </summary>
        /// <remarks>
        /// Not sure extra memory (at least 32 bytes per formula) is worth less CPU: O(1) vs O(log N)....
        /// </remarks>
        private readonly Dictionary<Area, AreaDependents> _precedentAreas;

        internal SheetDependencyTree()
        {
            _tree = new RBush<AreaDependents>();
            _precedentAreas = new Dictionary<Area, AreaDependents>();
        }

        internal bool IsEmpty => _tree.Count == 0;

        internal void AddDependent(Area precedentRange, Dependent dependent)
        {
            if (!_precedentAreas.TryGetValue(precedentRange, out var precedentArea))
            {
                precedentArea = new AreaDependents(ToEnvelope(precedentRange), dependent);
                _precedentAreas.Add(precedentRange, precedentArea);
                _tree.Insert(precedentArea);
            }
            else
            {
                precedentArea.AddDependent(dependent);
            }
        }

        internal IReadOnlyList<AreaDependents> FindDependentsAreas(Area dirtyRange)
        {
            return _tree.Search(ToEnvelope(dirtyRange));
        }

        /// <summary>
        /// Remove a dependency of <paramref name="formula"/> on a
        /// <paramref name="precedentRange"/> from the sheet dependency tree.
        /// </summary>
        /// <param name="precedentRange">A precedent area in the sheet.</param>
        /// <param name="formula">Formula depending on the <paramref name="precedentRange"/>.</param>
        internal void RemoveDependent(Area precedentRange, XLCellFormula formula)
        {
            if (!_precedentAreas.TryGetValue(precedentRange, out var precedentArea))
                return;

            precedentArea.RemoveDependent(formula);
            if (precedentArea.Dependents.Count == 0)
            {
                _tree.Delete(precedentArea);
                _precedentAreas.Remove(precedentRange);
            }
        }

        internal void RenameSheet(string oldSheetName, string newSheetName)
        {
            // Area dependents instances are shared among _precedentAreas and _tree, so it is
            // enough to change _precedentAreas.
            foreach (var areaDependents in _precedentAreas.Values)
                areaDependents.RenameSheet(oldSheetName, newSheetName);
        }

        private static Envelope ToEnvelope(Area range)
        {
            return new Envelope(range.LeftColumn, range.TopRow, range.RightColumn, range.BottomRow);
        }
    }
}
