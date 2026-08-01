using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Parser;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine.Visitors;

/// <summary>
/// A collection of all references in the book (not others) found in a formula.
/// Created by <see cref="CollectRefsFactory"/>.
/// </summary>
internal sealed class FormulaReferences
{
    /// <summary>
    /// Is there a <c>#REF!</c> anywhere in the formula?
    /// </summary>
    internal bool ContainsRefError { get; set; }

    /// <summary>
    /// Areas without a sheet found in the formula.
    /// </summary>
    internal HashSet<XLReference> References { get; } = [];

    /// <summary>
    /// Areas with a sheet found in the formula.
    /// </summary>
    internal HashSet<XLSheetReference> SheetReferences { get; } = [];

    /// <summary>
    /// Structured references (<c>Sales[Amount]</c>) found in the formula, kept as written rather
    /// than as areas: a table can be resized, renamed or deleted after the formula was parsed, so
    /// the area a reference covers is only true at the moment it is asked for.
    /// </summary>
    private HashSet<StructuredReference> StructuredReferences { get; } = new();

    internal static FormulaReferences ForFormula(string formula)
    {
        var references = new FormulaReferences();
        FormulaParser<object?, object?, FormulaReferences>.CellFormulaA1(formula, references,
            CollectRefsFactory.Instance);
        return references;
    }

    internal bool ContainsSheet(string worksheetName)
    {
        return SheetReferences.Any(x => XLHelper.SheetComparer.Equals(x.Sheet, worksheetName));
    }

    internal XLRanges GetExternalRanges(XLWorkbook workbook, Point anchor)
    {
        var list = new XLRanges();
        foreach (var reference in SheetReferences)
        {
            if (workbook.TryGetWorksheet(reference.Sheet, out XLWorksheet? sheet))
            {
                var rangeAddress = reference.Reference.ToRangeAddress(sheet, anchor);
                list.Add(sheet.Range(rangeAddress));
            }
        }

        if (StructuredReferences.Count > 0)
        {
            var scope = new DefinedNameScope(workbook, anchor);
            foreach (var reference in StructuredReferences)
            {
                // A reference that does not resolve today — an unknown table, a column since
                // renamed — contributes no range, the same answer an unknown table has always
                // given. Excel shows such a name as #REF! rather than refusing the workbook, and
                // whatever later restores the table makes the name resolve again.
                if (!StructuredReferenceResolver.TryResolve(scope, reference.ToNode(), out var sheet, out var area, out _))
                    continue;

                list.Add(sheet.Range(area.TopRow, area.LeftColumn, area.BottomRow, area.RightColumn));
            }
        }

        return list;
    }

    /// <summary>
    /// A structured reference as the formula wrote it. A value type so that the same reference
    /// written twice in one formula counts once.
    /// </summary>
    /// <remarks>
    /// Only references naming a table are collected — <see cref="CollectRefsFactory"/> does not
    /// override the table-less or external-workbook overloads — so <see cref="Table"/> is never
    /// null and <see cref="IStructuredReferenceScope.Worksheet"/> is never reached.
    /// </remarks>
    private readonly record struct StructuredReference(
        string Table,
        StructuredReferenceArea Area,
        string? FirstColumn,
        string? LastColumn)
    {
        internal StructuredReferenceNode ToNode() => new(null, Table, Area, FirstColumn, LastColumn);
    }

    /// <summary>
    /// The formula location a defined name can offer the resolver.
    /// </summary>
    /// <remarks>
    /// A defined name is not written in a cell, so it has no formula location of its own and the
    /// caller's anchor stands in for one. That only matters to <c>[@Column]</c>, which needs a row
    /// to mean anything: against the default anchor it lands outside the table's data and resolves
    /// to no range, which is the right answer for a name Excel would show as <c>#VALUE!</c>.
    /// </remarks>
    private readonly struct DefinedNameScope : IStructuredReferenceScope
    {
        public DefinedNameScope(XLWorkbook workbook, Point anchor)
        {
            Workbook = workbook;
            FormulaPoint = anchor;
        }

        public XLWorkbook Workbook { get; }

        public Point FormulaPoint { get; }

        /// <inheritdoc />
        /// <remarks>Read only for a table-less reference, which is never collected.</remarks>
        public XLWorksheet Worksheet =>
            throw new InvalidOperationException("A defined name has no sheet to resolve a table-less reference against.");
    }

    /// <summary>
    /// Factory to get all references (cells, tables, names) in local workbook.
    /// </summary>
    private sealed class CollectRefsFactory : CollectVisitor<FormulaReferences>
    {
        public static readonly CollectRefsFactory Instance = new();

        public override object? ErrorNode(FormulaReferences context, SymbolRange range, ReadOnlySpan<char> error)
        {
            context.ContainsRefError = true;
            return base.ErrorNode(context, range, error);
        }

        public override object? Reference(FormulaReferences context, SymbolRange range, ReferenceArea reference)
        {
            context.References.Add(new XLReference(reference));
            return base.Reference(context, range, reference);
        }

        public override object? SheetReference(FormulaReferences context, SymbolRange range, string sheet,
            ReferenceArea reference)
        {
            context.SheetReferences.Add(new XLSheetReference(sheet, new XLReference(reference)));
            return base.SheetReference(context, range, sheet, reference);
        }

        public override object? StructureReference(FormulaReferences context, SymbolRange range, string table,
            StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
        {
            context.StructuredReferences.Add(new StructuredReference(table, area, firstColumn, lastColumn));

            return base.StructureReference(context, range, table, area, firstColumn, lastColumn);
        }
    }
}
