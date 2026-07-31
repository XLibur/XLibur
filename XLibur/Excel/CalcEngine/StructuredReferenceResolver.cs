using System;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Parser;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// What <see cref="StructuredReferenceResolver"/> needs to know about the formula it is
/// resolving for. Each member is reached only on the paths that require it — a named table
/// reference never touches <see cref="Worksheet"/> or <see cref="FormulaPoint"/> — so an
/// implementation is free to throw for one it cannot supply, as
/// <see cref="CalcContext"/> does when evaluating an expression with no anchoring cell.
/// </summary>
internal interface IStructuredReferenceScope
{
    /// <summary>Workbook to look a named table up in.</summary>
    XLWorkbook Workbook { get; }

    /// <summary>Sheet holding the formula. Only needed for a table-less reference.</summary>
    XLWorksheet Worksheet { get; }

    /// <summary>Cell holding the formula. Only needed for <c>[@col]</c> and table-less.</summary>
    Point FormulaPoint { get; }
}

/// <summary>
/// Resolves a <see cref="StructuredReferenceNode"/> (<c>Table1[Amount]</c>,
/// <c>Table1[[#Headers],[#Data]]</c>, <c>[@Amount]</c>) to the sheet area it covers.
/// </summary>
/// <remarks>
/// Shared by <see cref="CalculationVisitor"/>, which evaluates the reference, and
/// <see cref="DependenciesVisitor"/>, which needs the same area to register the formula's
/// precedents. It works through <see cref="IStructuredReferenceScope"/> rather than taking the
/// values directly so that each is read only where it is actually used: an expression evaluated
/// without an anchoring cell can still resolve <c>Table1[Amount]</c>, and must still fail with
/// <c>#REF!</c> rather than throwing for an unknown table name.
/// </remarks>
internal static class StructuredReferenceResolver
{
    /// <summary>
    /// Resolve <paramref name="node"/> against <paramref name="scope"/>.
    /// </summary>
    /// <param name="scope">The formula being resolved for.</param>
    /// <param name="node">The reference to resolve.</param>
    /// <param name="worksheet">
    /// The sheet <paramref name="range"/> lies on, when this returns <c>true</c>. Table names are
    /// workbook scoped, so that is the sheet owning the table and not necessarily the scope's own
    /// sheet. For a table-less reference the two are the same by construction.
    /// </param>
    /// <param name="range">The area covered on <paramref name="worksheet"/>, when this returns <c>true</c>.</param>
    /// <param name="error">Why resolution failed, when this returns <c>false</c>.</param>
    internal static bool TryResolve<TScope>(
        TScope scope,
        StructuredReferenceNode node,
        [NotNullWhen(true)] out XLWorksheet? worksheet,
        out Area range,
        out XLError error)
        where TScope : IStructuredReferenceScope
    {
        // We don't support external links.
        if (node.Prefix is not null)
        {
            worksheet = null;
            range = default;
            error = XLError.CellReference;
            return false;
        }

        if (!TryGetTable(scope, node.Table, out var table) ||
            !TryGetColumns(table, node.FirstColumn, node.LastColumn, out var colStart, out var colEnd))
        {
            worksheet = null;
            range = default;
            error = XLError.CellReference;
            return false;
        }

        // Row range is always continuous, so the result is an area. [[#Header],[#Totals]] is
        // not allowed by grammar.
        if (!TryGetRows(scope, table, node.Area, out var rowStart, out var rowEnd, out error))
        {
            worksheet = null;
            range = default;
            return false;
        }

        worksheet = table.Worksheet;
        range = new Area(rowStart, colStart, rowEnd, colEnd);
        error = default;
        return true;
    }

    private static bool TryGetTable<TScope>(
        TScope scope,
        string? tableName,
        [NotNullWhen(true)] out XLTable? table)
        where TScope : IStructuredReferenceScope
    {
        // table-less references are allowed only in a table area. Excel doesn't allow
        // setting it in the GUI, but interprets such situations as #REF!.
        if (tableName is not null)
        {
            return scope.Workbook.TryGetTable(tableName, out table);
        }

        // Avoid LINQ allocation.
        var formulaPoint = scope.FormulaPoint;
        foreach (var sheetTable in scope.Worksheet.Tables)
        {
            if (sheetTable.Area.Contains(formulaPoint))
            {
                table = sheetTable;
                return true;
            }
        }

        table = null;
        return false;
    }

    private static bool TryGetColumns(
        XLTable table,
        string? firstColumn,
        string? lastColumn,
        out int columnStartNo,
        out int columnEndNo)
    {
        var area = table.Area;
        if (!TryGetColumn(table, firstColumn, area.LeftColumn, out columnStartNo))
        {
            columnEndNo = 0;
            return false;
        }

        if (!TryGetColumn(table, lastColumn, area.RightColumn, out columnEndNo))
        {
            return false;
        }

        if (columnStartNo > columnEndNo)
            (columnEndNo, columnStartNo) = (columnStartNo, columnEndNo);

        return true;
    }

    private static bool TryGetColumn(XLTable table, string? column, int defaultColumn, out int columnNo)
    {
        if (column is null)
        {
            columnNo = defaultColumn;
            return true;
        }

        if (!table.FieldNames.TryGetValue(column, out var field))
        {
            columnNo = 0;
            return false;
        }

        columnNo = field.Index + table.Area.LeftColumn;
        return true;
    }

    private static bool TryGetRows<TScope>(
        TScope scope,
        XLTable table,
        StructuredReferenceArea tableArea,
        out int rowStartNo,
        out int rowEndNo,
        out XLError error)
        where TScope : IStructuredReferenceScope
    {
        var area = table.Area;
        var dataEndRowNo = table.ShowTotalsRow ? area.BottomRow - 1 : area.BottomRow;

        if (tableArea == StructuredReferenceArea.Totals && !table.ShowTotalsRow)
        {
            rowStartNo = rowEndNo = 0;
            error = XLError.CellReference;
            return false;
        }

        if (tableArea == StructuredReferenceArea.ThisRow)
        {
            var thisRow = scope.FormulaPoint.Row;
            if (area.TopRow >= thisRow || dataEndRowNo < thisRow)
            {
                rowStartNo = rowEndNo = 0;
                error = XLError.IncompatibleValue;
                return false;
            }

            rowStartNo = thisRow;
            rowEndNo = thisRow;
            error = default;
            return true;
        }

        (rowStartNo, rowEndNo) = tableArea switch
        {
            StructuredReferenceArea.None or
            StructuredReferenceArea.Data => (area.TopRow + 1, dataEndRowNo),
            StructuredReferenceArea.Headers => (area.TopRow, area.TopRow),
            StructuredReferenceArea.Headers | StructuredReferenceArea.Data => (area.TopRow, dataEndRowNo),
            StructuredReferenceArea.Totals => (area.BottomRow, area.BottomRow),
            StructuredReferenceArea.Totals | StructuredReferenceArea.Data => (area.TopRow + 1, area.BottomRow),
            StructuredReferenceArea.All => (area.TopRow, area.BottomRow),
            _ => throw new NotSupportedException($"Unexpected value {tableArea}.")
        };

        error = default;
        return true;
    }
}
