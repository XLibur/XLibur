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
    internal DependenciesContext(SheetArea formulaArea, XLWorkbook workbook)
    {
        FormulaArea = formulaArea;
        Workbook = workbook;
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
    internal SheetArea FormulaArea { get; }

    public XLWorkbook Workbook { get; }

    /// <summary>
    /// The result. Visitor adds all areas/names formula depends on to this.
    /// </summary>
    internal FormulaDependencies Dependencies { get; } = new();

    /// <summary>
    /// Add areas to a list of areas the formula depends on. Disregards duplicate entries.
    /// </summary>
    internal void AddAreas(List<SheetArea> sheetAreas) => Dependencies.AddAreas(sheetAreas);

    /// <summary>
    /// Add name to a list of names the formula depends on. Disregards duplicate entries.
    /// </summary>
    internal void AddName(XLName name) => Dependencies.AddName(name);
}
