using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// A list of objects a cell formula depends on. If one of them changes,
/// the formula value might no longer be accurate and needs to be recalculated.
/// </summary>
/// <remarks>
/// The dependency tree does not keep this object. It collects each formula into one reused
/// instance, and keeps only the areas, as an array of the exact size (see <see cref="ToAreaArray"/>).
/// One instance per formula, with two sets, was about 330 B of the ~770 B that the tree kept for
/// each formula (#513).
/// </remarks>
internal sealed class FormulaDependencies
{
    private readonly HashSet<SheetArea> _areas = [];
    private readonly HashSet<XLName> _names = [];

    /// <summary>
    /// List of areas the formula depends on. It is likely a superset of an accurate
    /// result for unusual formulas, but if a value in an area changes, the dependent
    /// formula should be marked as dirty.
    /// </summary>
    public IReadOnlyCollection<SheetArea> Areas => _areas;

    /// <summary>
    /// A collection of names in the formula. If a name changes (added, deleted),
    /// the formula dependencies should be refreshed, because new name might refer to
    /// different references (e.g., a name previously referred to <c>A5</c> and is redefined
    /// to <c>B7</c> or just value <c>7</c> =&gt; formula no longer depends on <c>A5</c>).
    /// </summary>
    /// <remarks>
    /// The dependency tree does not keep the names, because nothing reads them after the visit. A
    /// tree that reacts to a changed name must keep them again.
    /// </remarks>
    public IReadOnlyCollection<XLName> Names => _names;

    /// <summary>
    /// Whether some of the formula's precedents cannot be known, because the parser refuses its text
    /// or the text of a defined name it uses. <see cref="Areas"/> then holds only the precedents that
    /// are known, and the formula is taken to depend on every cell.
    /// </summary>
    public bool HasUnknownPrecedents { get; private set; }

    internal void MarkPrecedentsUnknown() => HasUnknownPrecedents = true;

    internal void AddAreas(List<SheetArea> sheetAreas)
    {
        // A loop over the list, not UnionWith: UnionWith takes an IEnumerable, so it boxes the
        // enumerator of the list, once for each reference in each formula of a build.
        foreach (var sheetArea in sheetAreas)
            _areas.Add(sheetArea);
    }

    internal void AddName(XLName name)
    {
        _names.Add(name);
    }

    /// <summary>
    /// A copy of <see cref="Areas"/>, in an array of the exact size.
    /// </summary>
    internal SheetArea[] ToAreaArray()
    {
        if (_areas.Count == 0)
            return [];

        var areas = new SheetArea[_areas.Count];
        _areas.CopyTo(areas);
        return areas;
    }

    /// <summary>
    /// Empty the collection, so that it can collect the next formula.
    /// </summary>
    internal void Clear()
    {
        _areas.Clear();
        _names.Clear();
        HasUnknownPrecedents = false;
    }
}
