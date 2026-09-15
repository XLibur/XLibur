using System.Collections.Generic;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Context for <see cref="DependenciesVisitor"/>, it is used
/// to collect all objects a formula depends on during calculation.
/// </summary>
internal sealed class DependenciesContext : IStructuredReferenceScope
{
    /// <param name="formulaArea">The area of the formula whose precedents are collected.</param>
    /// <param name="workbook">The workbook, to find names, sheets and tables.</param>
    /// <param name="dependencies">
    /// The collection that receives the precedents. The dependency tree passes the same empty
    /// collection for each formula, so a build does not allocate one per formula.
    /// </param>
    internal DependenciesContext(SheetArea formulaArea, XLWorkbook workbook, FormulaDependencies dependencies)
    {
        FormulaArea = formulaArea;
        Workbook = workbook;
        Dependencies = dependencies;
    }

    /// <summary>
    /// Make the context ready for the next formula. The dependency tree keeps one context for every
    /// formula it adds, so a build does not allocate one per formula (#513).
    /// </summary>
    internal void Reset(SheetArea formulaArea, XLWorkbook workbook, FormulaDependencies dependencies)
    {
        FormulaArea = formulaArea;
        Workbook = workbook;
        Dependencies = dependencies;
    }

    /// <inheritdoc />
    /// <remarks>
    /// Resolved from <see cref="FormulaArea"/> on demand, and only read for a table-less
    /// reference. The sheet always exists here: a formula is registered in the tree under the
    /// name of the sheet it was found on, so the lookup can only fail if the two have gone out
    /// of step. Mirrors <see cref="CalcContext"/>, which throws the same exception for the same
    /// member when it has no sheet to offer.
    /// </remarks>
    XLWorksheet IStructuredReferenceScope.Worksheet =>
        Workbook.TryGetWorksheet(FormulaArea.Name, out XLWorksheet? sheet)
            ? sheet
            : throw new MissingContextException();

    /// <inheritdoc />
    /// <remarks>
    /// For an array formula this is the top-left cell of the formula's area, which is the cell
    /// Excel treats as owning the formula.
    /// </remarks>
    Point IStructuredReferenceScope.FormulaPoint => FormulaArea.Area.FirstPoint;

    /// <summary>
    /// An area of a formula, in most cases just one cell, for array formulas area of cells.
    /// </summary>
    internal SheetArea FormulaArea { get; private set; }

    public XLWorkbook Workbook { get; private set; }

    /// <summary>
    /// The result. Visitor adds all areas/names formula depends on to this.
    /// </summary>
    internal FormulaDependencies Dependencies { get; private set; }

    /// <summary>
    /// The text that <see cref="PrecedentsFactory"/> reads: the cell formula, or the formula of a
    /// defined name inside it. A function from another cell reads its name from this text.
    /// </summary>
    internal string Text { get; set; } = string.Empty;

    /// <summary>
    /// Set by <see cref="PrecedentsFactory"/> when it cannot give the precedents that
    /// <see cref="DependenciesVisitor"/> gives, so the tree reads the formula through its AST.
    /// </summary>
    internal bool NeedsAst { get; set; }

    /// <summary>
    /// The areas of the reference nodes of the walk, which <see cref="PrecedentsFactory"/> refers to by
    /// their index. The list is kept for the next formula, so a walk allocates no node.
    /// </summary>
    private readonly List<ReferenceAreas> _nodes = [];

    /// <summary>
    /// Start the walk of a cell formula.
    /// </summary>
    internal void BeginWalk(string text)
    {
        Text = text;
        NeedsAst = false;
        _nodes.Clear();
    }

    /// <summary>
    /// Keep the areas of a reference node, and give its handle.
    /// </summary>
    internal int PushNode(ReferenceAreas areas)
    {
        _nodes.Add(areas);
        return _nodes.Count - 1;
    }

    /// <summary>
    /// The areas of the node with the handle, or <see cref="ReferenceAreas.None"/> for
    /// <see cref="PrecedentsFactory.None"/>.
    /// </summary>
    internal ReferenceAreas GetNode(int node) => node < 0 ? ReferenceAreas.None : _nodes[node];

    /// <summary>
    /// Add areas to a list of areas the formula depends on. Disregards duplicate entries.
    /// </summary>
    internal void AddAreas(in ReferenceAreas sheetAreas) => Dependencies.AddAreas(sheetAreas);

    /// <summary>
    /// Add name to a list of names the formula depends on. Disregards duplicate entries.
    /// </summary>
    internal void AddName(XLName name) => Dependencies.AddName(name);
}
