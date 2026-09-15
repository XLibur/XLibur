using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
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
    /// <remarks>
    /// Each formula keeps its precedent areas in an array of the exact size, not a
    /// <see cref="FormulaDependencies"/> with two sets, because the tree keeps one for every formula
    /// in the workbook (#513).
    /// </remarks>
    private readonly Dictionary<XLCellFormula, SheetArea[]> _dependencies = new();

    /// <summary>
    /// Collects the precedents of the formula that <see cref="AddFormula"/> adds. One instance serves
    /// every formula: the tree is used from one thread at a time, and adding a formula never adds
    /// another one.
    /// </summary>
    private readonly FormulaDependencies _scratch = new();

    /// <summary>
    /// The context of the visit, kept for every formula for the same reason as <see cref="_scratch"/>.
    /// Created by the first visit, because only a visit knows the workbook.
    /// </summary>
    private DependenciesContext? _context;

    /// <summary>
    /// During <see cref="CreateFrom"/>, the ASTs of the shared formulas parsed so far, keyed by their
    /// R1C1 text, with <c>null</c> for text that the parser refused. <c>null</c> outside a build, so
    /// that the tree does not keep the ASTs.
    /// </summary>
    private Dictionary<string, Formula?>? _sharedAsts;

    /// <summary>
    /// Visitor to extract precedents of formulas.
    /// </summary>
    private readonly DependenciesVisitor _visitor;

    /// <summary>
    /// Collects the precedents of a formula while the parser reads it, with no AST (#513).
    /// </summary>
    private readonly PrecedentsFactory _factory = new();

    /// <summary>
    /// A dependency tree for each sheet (key is sheet name).
    /// </summary>
    private readonly Dictionary<string, SheetDependencyTree> _sheetTrees = new(XLHelper.SheetComparer);

    /// <summary>
    /// Formulas whose precedents cannot be known, each with its area. Any change marks them dirty,
    /// with whatever depends on them (see <see cref="MarkDirty"/>). Empty in almost every workbook.
    /// </summary>
    private readonly Dictionary<XLCellFormula, SheetArea> _unknownPrecedents = new();

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

        // Sized once, so that the dictionary does not grow and copy itself about 18 times on a
        // workbook of 200,000 formulas.
        tree._dependencies.EnsureCapacity(CountFormulaCells(workbook));

        // Each sheet tree is empty, so it can take all its areas in one bulk load at the end. One
        // RBush insert per area allocated enumerators, LINQ iterators and node arrays for each insert
        // and each node split, and areas that arrive row by row split many nodes (#513).
        foreach (var sheetTree in tree._sheetTrees.Values)
            sheetTree.BeginBulkLoad();

        // A formula filled down a column is one shared formula in a file that Excel saved. Its cells
        // share one parse of the R1C1 text, instead of one parse of each cell's A1 text (#513).
        tree._sharedAsts = new Dictionary<string, Formula?>(StringComparer.Ordinal);

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
                // syntax, so the parser refuses it and it would be taken to depend on every cell,
                // which is not what its inputs are. Registering them needs
                // precedents built from Input1/Input2 and the table's header formulas instead of
                // from an AST. XLibur does not evaluate data tables either (there is no TABLE
                // function), so the only gain would be dropping the full-recalculation trigger in
                // XLCalcEngine.TryEvaluateSingleCell.
                // FormulaType.Shared is never produced, so it needs no handling here.
            }
        }

        foreach (var sheetTree in tree._sheetTrees.Values)
            sheetTree.EndBulkLoad();

        tree._sharedAsts = null;
        return tree;
    }
#pragma warning restore S3776

    /// <summary>
    /// The number of cells that hold a formula. An array formula is counted once for each of its
    /// cells, but the tree adds it once, so the count can be larger than the number of formulas the
    /// tree adds.
    /// </summary>
    private static int CountFormulaCells(XLWorkbook workbook)
    {
        var count = 0;
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
                count++;
        }

        return count;
    }

    /// <summary>
    /// Add a formula to the dependency tree.
    /// </summary>
    /// <param name="formulaArea">Area of a formula, for normal cells 1x1, for array can be larger.</param>
    /// <param name="formula">The cell formula.</param>
    /// <param name="workbook">Workbook that is used to find precedents (names ect.).</param>
    /// <exception cref="ArgumentException">Formula already is in the tree.</exception>
    internal void AddFormula(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook)
    {
        var precedents = _scratch;
        precedents.Clear();
        CollectPrecedents(formulaArea, formula, workbook, precedents);

        var precedentAreas = precedents.ToAreaArray();
        _dependencies.Add(formula, precedentAreas);

        if (precedents.HasUnknownPrecedents)
            _unknownPrecedents[formula] = formulaArea;

        foreach (var precedentArea in precedentAreas)
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
    }

    /// <summary>
    /// The precedents of a formula, as <see cref="AddFormula"/> finds them, without adding the
    /// formula to the tree. Tests use it to read the names as well, which the tree does not keep.
    /// </summary>
    internal FormulaDependencies GetPrecedents(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook)
    {
        var precedents = new FormulaDependencies();
        CollectPrecedents(formulaArea, formula, workbook, precedents);
        return precedents;
    }

    /// <summary>
    /// The precedent areas that the tree keeps for a formula in it. For tests.
    /// </summary>
    internal IReadOnlyList<SheetArea> GetKeptPrecedents(XLCellFormula formula) => _dependencies[formula];

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
        if (!_dependencies.Remove(formula, out var precedentAreas))
            return;

        _unknownPrecedents.Remove(formula);

        foreach (var precedentArea in precedentAreas)
        {
            // An area on a sheet the workbook does not have was never added to a sheet tree (see
            // AddFormula), so there is nothing to take out. This threw, so a formula such as
            // Missing!A1 could not be replaced once the tree was built.
            if (_sheetTrees.TryGetValue(precedentArea.Name, out var sheetTree))
                sheetTree.RemoveDependent(precedentArea.Area, formula);
        }
    }

    internal void AddSheetTree(IXLWorksheet sheet)
    {
        _sheetTrees.Add(sheet.Name, new SheetDependencyTree());
    }

    internal void RenameSheet(string oldSheetName, string newSheetName)
    {
        // In place, because each array belongs to the tree. A formula that reads both the old name
        // and a missing sheet that has the new name then holds the same area twice. RemoveFormula
        // finds nothing to remove for the second copy, so that does no harm.
        foreach (var precedentAreas in _dependencies.Values)
        {
            for (var i = 0; i < precedentAreas.Length; ++i)
            {
                var precedentArea = precedentAreas[i];
                if (XLHelper.SheetComparer.Equals(precedentArea.Name, oldSheetName))
                    precedentAreas[i] = new SheetArea(newSheetName, precedentArea.Area);
            }
        }

        foreach (var formula in _unknownPrecedents.Keys.ToList())
        {
            var formulaArea = _unknownPrecedents[formula];
            if (XLHelper.SheetComparer.Equals(formulaArea.Name, oldSheetName))
                _unknownPrecedents[formula] = new SheetArea(newSheetName, formulaArea.Area);
        }

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
    /// The dependents that one step of <see cref="MarkDirty"/> finds, reused for the same reason as
    /// <see cref="_walkQueue"/>.
    /// </summary>
    private readonly List<Dependents> _found = new();

    /// <summary>
    /// Capacity above which <see cref="MarkDirty"/> releases the reused queue's backing array
    /// instead of holding it for the tree's lifetime. One unusually wide walk on a long-lived
    /// workbook would otherwise retain a slot per node it visited, along with a sheet-name
    /// reference each, long after the small walks that follow could use them. The threshold is
    /// well above any ordinary closure, so the common path never trims.
    /// </summary>
    private const int WalkQueueRetainedCapacity = 4096;

    /// <summary>
    /// Guards the assumption that <see cref="MarkDirty"/> is never re-entered, which is what makes
    /// a single reused <see cref="_walkQueue"/> safe. Nothing the walk calls evaluates a formula or
    /// re-enters the tree today; a future change that did would silently corrupt the outer walk's
    /// queue rather than fail, so it is asserted in debug builds instead of left to comments.
    /// </summary>
    private bool _walkInProgress;

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
        Debug.Assert(!_walkInProgress, "MarkDirty is not re-entrant: it reuses a single walk queue.");
        _walkInProgress = true;

        var walkId = Interlocked.Increment(ref _walkGeneration);

        // BFS vs DFS: Although the longest chain found in the wild is 1000
        // formulas long, attacker could supply malicious excel with recursion
        // leading to stack overflow => use queue even with extra allocation cost.
        var queue = _walkQueue;
        var found = _found;
        try
        {
            queue.Enqueue(dirtyArea);

            // A formula whose precedents are unknown may read the changed area, so it is taken to read
            // every cell: any change marks it dirty, together with whatever depends on it. If an
            // earlier walk did that and the formula has stayed dirty since, this walk has nothing to
            // mark (see XLCellFormula.DependentsMarkedDirty). A refused formula never becomes clean,
            // so without the skip every edit re-marked the same closure. It is visited either way, so
            // a skipped formula is not reached again through a precedent it does know.
            foreach (var (formula, formulaArea) in _unknownPrecedents)
            {
                if (!formula.TryVisit(walkId) || formula.DependentsMarkedDirty)
                    continue;

                formula.MarkDirtyWithDependents();
                queue.Enqueue(formulaArea);
            }

            while (queue.Count > 0)
            {
                var affectedArea = queue.Dequeue();
                var sheetTree = _sheetTrees[affectedArea.Name];
                found.Clear();
                sheetTree.FindDependents(affectedArea.Area, found);
                foreach (var precedent in found)
                {
                    for (var i = 0; i < precedent.Count; ++i)
                    {
                        var dependent = precedent[i];

                        // Ensure we don't end up in an infinite cycle: a formula already enqueued
                        // by this walk is not enqueued again, regardless of its dirty state.
                        if (!dependent.Formula.TryVisit(walkId))
                            continue;

                        dependent.MarkDirty();
                        queue.Enqueue(dependent.FormulaArea);
                    }
                }
            }
        }
        finally
        {
            // Clear on the way out, not on the way in, so a walk that threw does not leave its
            // entries — and their sheet-name references — reachable until the next MarkDirty.
            queue.Clear();

            // EnsureCapacity(0) grows nothing and returns the capacity, which Queue<T> does not
            // otherwise expose on every target framework.
            if (queue.EnsureCapacity(0) > WalkQueueRetainedCapacity)
                queue.TrimExcess();

            // A dirty area such as a whole column can find thousands of precedent cells in one step.
            found.Clear();
            if (found.Capacity > WalkQueueRetainedCapacity)
                found.TrimExcess();

            _walkInProgress = false;
        }
    }

    private void CollectPrecedents(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook,
        FormulaDependencies precedents)
    {
        var context = _context;
        if (context is null)
            _context = context = new DependenciesContext(formulaArea, workbook, precedents);
        else
            context.Reset(formulaArea, workbook, precedents);

        // A shared formula from a file that Excel saved: one AST serves every cell of its group.
        if (TryGetSharedAst(formulaArea, formula, workbook, out var sharedAst))
        {
            VisitAst(context, sharedAst);
            return;
        }

        // Any other formula: the precedents are collected while the parser reads the text, and no
        // AST is built (#513).
        if (_factory.TryCollect(formula.A1, context))
        {
            if (!context.NeedsAst)
                return;

            // Rare: the walk added precedents that the visitor does not add (see PrecedentsFactory).
            precedents.Clear();
            if (formula.TryGetAst(workbook.CalcEngine, out var ast))
            {
                VisitAst(context, ast);
                return;
            }
        }

        // A refused formula's precedents cannot be known: the parser could not read its references
        // (ADR 0002), or the calc engine refused a part of it, such as a function with the wrong
        // number of arguments (#543). Before, the parse threw, and one such formula stopped every
        // write to the workbook and every recalculation from building the tree (#489). It is now
        // taken to depend on every cell, so any change marks it dirty (see MarkDirty). That matters
        // for a formula a load gave a cached value: it is clean, and would otherwise keep that value
        // after an edit it may read. The cell fails when it is evaluated. The walk may have added
        // precedents before the text was refused, and they go.
        precedents.Clear();
        precedents.MarkPrecedentsUnknown();
    }

    private void VisitAst(DependenciesContext context, Formula ast)
    {
        var rootReference = ast.AstRoot.Accept(context, _visitor);

        // If formula references are propagated to the root, make sure to add them.
        if (rootReference.IsReference)
            context.AddAreas(rootReference);
    }

    /// <summary>
    /// Get the AST of a shared formula. During <see cref="CreateFrom"/>, a formula that the loader read
    /// from a shared formula uses the one AST of its group, parsed from the R1C1 text. The visitor
    /// resolves each relative reference against the anchor of the formula, so one AST serves every
    /// cell of the group (#513).
    /// </summary>
    /// <remarks>
    /// The loader made the A1 text of each cell from the R1C1 text, so both give the same references.
    /// If the parser refuses the R1C1 text, the cell reads its own A1 text instead, so that a refusal
    /// is always the one that the A1 text gets.
    /// </remarks>
    private bool TryGetSharedAst(SheetArea formulaArea, XLCellFormula formula, XLWorkbook workbook,
        [NotNullWhen(true)] out Formula? ast)
    {
        var sharedAsts = _sharedAsts;
        if (sharedAsts is not null && formula.TryGetSharedR1C1(formulaArea.Area.FirstPoint, out var r1c1))
        {
            if (!sharedAsts.TryGetValue(r1c1, out var sharedAst))
            {
                workbook.CalcEngine.TryParseR1C1(r1c1, out sharedAst);
                sharedAsts.Add(r1c1, sharedAst);
            }

            if (sharedAst is not null)
            {
                ast = sharedAst;
                return true;
            }
        }

        ast = null;
        return false;
    }

    /// <summary>
    /// The formulas that depend on one precedent: a cell, or an area of more than one cell (see
    /// <see cref="AreaDependents"/>). If anything in the precedent potentially changes, all
    /// dependents might also change.
    /// </summary>
    private class Dependents
    {
        /// <summary>
        /// The first formula that depends on the precedent. It is held in a field, because most
        /// precedents have one dependent: a list for each precedent cost the list and its array (#513).
        /// </summary>
        private Dependent _first;

        /// <summary>
        /// The formulas after the first that depend on the precedent, or <c>null</c> until a second
        /// one does.
        /// </summary>
        private List<Dependent>? _others;

        internal Dependents(Dependent firstDependent)
        {
            _first = firstDependent;
        }

        /// <summary>
        /// The number of formulas that depend on the precedent, always at least one.
        /// </summary>
        internal int Count => 1 + (_others?.Count ?? 0);

        /// <summary>
        /// A formula that depends on the precedent, from <c>0</c> to <see cref="Count"/> - 1.
        /// </summary>
        internal Dependent this[int index] => index == 0 ? _first : _others![index - 1];

        internal void AddDependent(Dependent dependent)
        {
            // Capacity 1, because a second dependent is often the last: an empty list grows to 4 on
            // its first add, 96 B more than one slot for each such area.
            (_others ??= new List<Dependent>(1)).Add(dependent);
        }

        /// <summary>
        /// Remove the dependent of <paramref name="formula"/>. Several different formulas can depend
        /// on the same precedent, so only the dependent of this formula goes.
        /// </summary>
        /// <returns>
        /// <c>true</c> when no formula depends on the precedent any more, and the caller must discard it.
        /// </returns>
        internal bool RemoveDependent(XLCellFormula formula)
        {
            if (_others is not null)
            {
                // Move the last element into the place of the removed one. The loop goes backwards,
                // so the element it moves has already been compared.
                for (var i = _others.Count - 1; i >= 0; --i)
                {
                    if (_others[i].Formula != formula)
                        continue;

                    _others[i] = _others[^1];
                    _others.RemoveAt(_others.Count - 1);
                }
            }

            if (_first.Formula != formula)
                return false;

            if (_others is null || _others.Count == 0)
                return true;

            _first = _others[^1];
            _others.RemoveAt(_others.Count - 1);
            return false;
        }

        internal void RenameSheet(string oldSheetName, string newSheetName)
        {
            _first = Renamed(_first, oldSheetName, newSheetName);
            if (_others is null)
                return;

            for (var i = 0; i < _others.Count; ++i)
                _others[i] = Renamed(_others[i], oldSheetName, newSheetName);
        }

        private static Dependent Renamed(Dependent dependent, string oldSheetName, string newSheetName)
        {
            if (!XLHelper.SheetComparer.Equals(dependent.FormulaArea.Name, oldSheetName))
                return dependent;

            var renamedArea = new SheetArea(newSheetName, dependent.FormulaArea.Area);
            return new Dependent(renamedArea, dependent.Formula);
        }
    }

    /// <summary>
    /// The formulas that depend on an area of more than one cell, kept in the R-tree of its sheet.
    /// </summary>
    private sealed class AreaDependents : Dependents, ISpatialData
    {
        /// <summary>
        /// An area in a sheet that is used by formulas, converted to RBush envelope.
        /// All RBush <c>double</c> coordinates are whole numbers.
        /// </summary>
        private readonly Envelope _area;

        internal AreaDependents(in Envelope area, Dependent firstDependent)
            : base(firstDependent)
        {
            _area = area;
        }

        /// <summary>
        /// The area in a sheet on which some formulas depend on.
        /// </summary>
        /// <example><c>SUM(A4:A6)</c> depends on <c>A4:A6</c> area.</example>.
        public ref readonly Envelope Envelope => ref _area;
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
        /// The precedent areas of more than one cell. The areas are not duplicated, though two areas
        /// might overlap.
        /// </summary>
        private readonly RBush<AreaDependents> _tree;

        /// <summary>
        /// The precedent areas of more than one cell in the sheet, for all formulas in the workbook.
        /// They are the objects in <see cref="_tree"/>, found here by their area.
        /// </summary>
        private readonly Dictionary<Area, AreaDependents> _precedentAreas;

        /// <summary>
        /// The precedent cells in the sheet, for all formulas in the workbook. A cell is not in
        /// <see cref="_tree"/>. Most precedents are one cell, and each cost an envelope, a place in an
        /// R-tree node, and a place in the sort of the bulk load (#513).
        /// </summary>
        private readonly Dictionary<Point, Dependents> _precedentCells;

        internal SheetDependencyTree()
        {
            _tree = new RBush<AreaDependents>();
            _precedentAreas = new Dictionary<Area, AreaDependents>();
            _precedentCells = new Dictionary<Point, Dependents>();
        }

        /// <summary>
        /// Set between <see cref="BeginBulkLoad"/> and <see cref="EndBulkLoad"/>. A new area then
        /// goes only into <see cref="_precedentAreas"/>, and <see cref="EndBulkLoad"/> loads all of
        /// them into <see cref="_tree"/> at once.
        /// </summary>
        private bool _bulkLoading;

        internal bool IsEmpty => _tree.Count == 0 && _precedentCells.Count == 0;

        /// <summary>
        /// Start to collect areas for one bulk load. Only <see cref="AddDependent"/> is valid until
        /// <see cref="EndBulkLoad"/>.
        /// </summary>
        internal void BeginBulkLoad()
        {
            Debug.Assert(_precedentAreas.Count == 0 && _precedentCells.Count == 0, "A bulk load fills an empty sheet tree.");
            _bulkLoading = true;
        }

        /// <summary>
        /// Load every area collected since <see cref="BeginBulkLoad"/> into the R-tree.
        /// </summary>
        internal void EndBulkLoad()
        {
            _bulkLoading = false;
            if (_precedentAreas.Count > 0)
                _tree.BulkLoad(_precedentAreas.Values);
        }

        internal void AddDependent(Area precedentRange, Dependent dependent)
        {
            if (IsCell(precedentRange))
            {
                var cell = precedentRange.FirstPoint;
                if (_precedentCells.TryGetValue(cell, out var precedentCell))
                    precedentCell.AddDependent(dependent);
                else
                    _precedentCells.Add(cell, new Dependents(dependent));

                return;
            }

            if (!_precedentAreas.TryGetValue(precedentRange, out var precedentArea))
            {
                precedentArea = new AreaDependents(ToEnvelope(precedentRange), dependent);
                _precedentAreas.Add(precedentRange, precedentArea);
                if (!_bulkLoading)
                    _tree.Insert(precedentArea);
            }
            else
            {
                precedentArea.AddDependent(dependent);
            }
        }

        /// <summary>
        /// Add to <paramref name="found"/> the dependents of each precedent cell and area that
        /// <paramref name="dirtyRange"/> intersects.
        /// </summary>
        internal void FindDependents(Area dirtyRange, List<Dependents> found)
        {
            Debug.Assert(!_bulkLoading, "The R-tree is incomplete during a bulk load.");
            if (_tree.Count > 0)
            {
                var areas = _tree.Search(ToEnvelope(dirtyRange));
                for (var i = 0; i < areas.Count; ++i)
                    found.Add(areas[i]);
            }

            if (_precedentCells.Count == 0)
                return;

            // A small range looks up each of its cells. A large one, such as a whole column, tests
            // each precedent cell instead. Either way, the cost is the smaller of the two counts.
            if ((long)dirtyRange.Width * dirtyRange.Height <= _precedentCells.Count)
            {
                for (var row = dirtyRange.TopRow; row <= dirtyRange.BottomRow; ++row)
                {
                    for (var column = dirtyRange.LeftColumn; column <= dirtyRange.RightColumn; ++column)
                    {
                        if (_precedentCells.TryGetValue(new Point(row, column), out var precedentCell))
                            found.Add(precedentCell);
                    }
                }
            }
            else
            {
                foreach (var (cell, precedentCell) in _precedentCells)
                {
                    if (dirtyRange.Contains(cell))
                        found.Add(precedentCell);
                }
            }
        }

        /// <summary>
        /// Remove a dependency of <paramref name="formula"/> on a
        /// <paramref name="precedentRange"/> from the sheet dependency tree.
        /// </summary>
        /// <param name="precedentRange">A precedent area in the sheet.</param>
        /// <param name="formula">Formula depending on the <paramref name="precedentRange"/>.</param>
        internal void RemoveDependent(Area precedentRange, XLCellFormula formula)
        {
            Debug.Assert(!_bulkLoading, "The R-tree is incomplete during a bulk load.");
            if (IsCell(precedentRange))
            {
                var cell = precedentRange.FirstPoint;
                if (_precedentCells.TryGetValue(cell, out var precedentCell) && precedentCell.RemoveDependent(formula))
                    _precedentCells.Remove(cell);

                return;
            }

            if (!_precedentAreas.TryGetValue(precedentRange, out var precedentArea))
                return;

            if (precedentArea.RemoveDependent(formula))
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

            foreach (var cellDependents in _precedentCells.Values)
                cellDependents.RenameSheet(oldSheetName, newSheetName);
        }

        private static bool IsCell(Area range) => range.FirstPoint == range.LastPoint;

        private static Envelope ToEnvelope(Area range)
        {
            return new Envelope(range.LeftColumn, range.TopRow, range.RightColumn, range.BottomRow);
        }
    }
}
