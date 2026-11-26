using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Tables;
using XLibur.Extensions;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

/// <summary>
/// A writer for table definition part.
/// </summary>
internal static class TablePartWriter
{
    internal static void SynchronizeTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
    {
        // Remove table definition parts that are not a part of workbook
        foreach (var tableDefinitionPart in worksheetPart.GetPartsOfType<TableDefinitionPart>().ToList())
        {
            var partId = worksheetPart.GetIdOfPart(tableDefinitionPart);
            var xlWorkbookContainsTable = tables.Cast<XLTable>().Any(t => t.RelId == partId);
            if (!xlWorkbookContainsTable)
            {
                worksheetPart.DeletePart(tableDefinitionPart);
            }
        }

        foreach (var xlTable in tables.Cast<XLTable>())
        {
            if (string.IsNullOrEmpty(xlTable.RelId))
            {
                xlTable.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                worksheetPart.AddNewPart<TableDefinitionPart>(xlTable.RelId);
            }
        }
    }

    /// <summary>
    /// Gives each table the <c>table/@id</c> it is written under, unique in the workbook. A loaded
    /// table keeps its id (#903), and a new table, or one whose id an earlier table took, gets one
    /// past every id kept. An id a table slicer cache names without its table being loaded counts
    /// as kept, so a new table is never handed that slicer (#971).
    /// </summary>
    internal static void AssignTableIds(IEnumerable<XLWorksheet> worksheets, SaveContext context)
    {
        var kept = new HashSet<uint>();
        var unassigned = new List<XLTable>();
        uint lastId = worksheets
            .SelectMany(ws => ws.SlicersInternal.Items)
            .Select(slicer => slicer.Cache)
            .Where(cache => cache.Table is null)
            .Select(cache => cache.TableId ?? 0)
            .DefaultIfEmpty()
            .Max();
        foreach (var xlTable in worksheets.SelectMany(ws => ws.Tables.Cast<XLTable>()))
        {
            if (xlTable.Id is { } id && id > 0 && kept.Add(id))
            {
                context.TableIds[xlTable] = id;
                lastId = Math.Max(lastId, id);
            }
            else
            {
                unassigned.Add(xlTable);
            }
        }

        foreach (var xlTable in unassigned)
        {
            xlTable.Id = ++lastId;
            context.TableIds[xlTable] = lastId;
        }
    }

    internal static void GenerateTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
    {
        foreach (var xlTable in tables.Cast<XLTable>())
        {
            var relId = xlTable.RelId;
            var tableDefinitionPart = (TableDefinitionPart)worksheetPart.GetPartById(relId!);
            GenerateTableDefinitionPartContent(tableDefinitionPart, xlTable, context);
        }
    }

    private static void GenerateTableDefinitionPartContent(TableDefinitionPart tableDefinitionPart, XLTable xlTable, SaveContext context)
    {
        var loaded = tableDefinitionPart.Table;

        // Normalise per axis so a table created over a reversed selection is written as "B2:E5"
        // rather than "B5:E2" verbatim. Normalize() is the address-level twin of
        // Area.FromRangeAddress and is what this needs rather than the Area itself: Area.ToString
        // collapses a one-cell rectangle to "A1" where a table's ST_Ref wants both corners, and
        // XLRangeBase.SheetRange throws for a table whose address a delete has invalidated, where
        // the corner-by-corner form still renders the "#REF!:#REF!" the writer emitted before.
        var tableAddress = xlTable.RangeAddress.Normalize();
        var reference = tableAddress.FirstAddress + ":" + tableAddress.LastAddress;
        var tableName = GetTableName(xlTable.Name, context);
        var table = new Table
        {
            Id = context.TableIds[xlTable],
            Name = tableName,
            DisplayName = tableName,
            Reference = reference
        };

        if (!xlTable.ShowHeaderRow)
            table.HeaderRowCount = 0;

        if (xlTable.ShowTotalsRow)
            table.TotalsRowCount = 1;
        else
            table.TotalsRowShown = false;

        var tableColumns = BuildTableColumns(xlTable, loaded, context);

        var tableStyleInfo1 = BuildTableStyleInfo(xlTable);

        if (xlTable.ShowAutoFilter)
        {
            var autoFilter1 = new AutoFilter();
            SetAutoFilterRange(xlTable);
            AutoFilterWriter.PopulateAutoFilter(xlTable.AutoFilter, autoFilter1, context);
            table.AppendChild(autoFilter1);
        }

        table.AppendChild(tableColumns);
        table.AppendChild(tableStyleInfo1);

        if (loaded is not null)
            TablePartCarryOver.Apply(loaded, table, xlTable, context.LoadedDxfs);

        tableDefinitionPart.Table = table;
    }

    private static TableColumns BuildTableColumns(XLTable xlTable, Table? loaded, SaveContext context)
    {
        var fields = xlTable.Fields.Cast<XLTableField>().ToList();
        AssignColumnIds(fields, loaded);

        var tableColumns = new TableColumns { Count = (uint)fields.Count };
        foreach (var xlField in fields)
            tableColumns.AppendChild(BuildTableColumn(xlField, xlField.Id!.Value, xlTable, context));

        return tableColumns;
    }

    /// <summary>
    /// Gives each column the <c>tableColumn/@id</c> it is written under, unique in its table. A
    /// loaded column keeps its id (#903), and a new one gets one past every id the table has or its
    /// loaded part had, so a column removed since the load does not lend its id to a new one.
    /// </summary>
    private static void AssignColumnIds(List<XLTableField> fields, Table? loaded)
    {
        var lastId = loaded?.TableColumns?.Elements<TableColumn>().Select(c => c.Id?.Value ?? 0).DefaultIfEmpty().Max() ?? 0;
        var kept = new HashSet<uint>();
        var unassigned = new List<XLTableField>();
        foreach (var field in fields)
        {
            if (field.Id is { } id && id > 0 && kept.Add(id))
                lastId = Math.Max(lastId, id);
            else
                unassigned.Add(field);
        }

        foreach (var field in unassigned)
            field.Id = ++lastId;
    }

    private static TableColumn BuildTableColumn(XLTableField xlField, uint columnId, XLTable xlTable, SaveContext context)
    {
        var fieldName = xlField.Name;
        var tableColumn = new TableColumn
        {
            Id = columnId,
            Name = fieldName.Replace("_x000a_", "_x005f_x000a_").Replace(Environment.NewLine, "_x000a_")
        };

        ApplyColumnDataFormat(tableColumn, xlField, xlTable, context);
        ApplyColumnFormula(tableColumn, xlField, xlTable);

        if (xlTable.ShowTotalsRow)
            ApplyColumnTotalsRow(tableColumn, xlField);

        return tableColumn;
    }

    private static void ApplyColumnDataFormat(TableColumn tableColumn, IXLTableField xlField, XLTable xlTable, SaveContext context)
    {
        // https://github.com/XLibur/XLibur/issues/513
        if (xlField.IsConsistentStyle())
        {
            var style = ((XLStyle)xlField.Column.Cells()
                .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                .First()
                .Style).Value;

            if (!DefaultStyleValue.Equals(style) && context.DifferentialFormats.TryGetValue(style, out int id))
                tableColumn.DataFormatId = UInt32Value.FromUInt32(Convert.ToUInt32(id));
        }
        else
            tableColumn.DataFormatId = null;
    }

    private static void ApplyColumnFormula(TableColumn tableColumn, IXLTableField xlField, XLTable xlTable)
    {
        if (xlField.IsConsistentFormula())
        {
            string formula = xlField.Column.Cells()
                .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                .First()
                .FormulaA1;

            while (formula.StartsWith('=') && formula.Length > 1)
                formula = formula.Substring(1);

            if (!string.IsNullOrWhiteSpace(formula))
                tableColumn.CalculatedColumnFormula = new CalculatedColumnFormula { Text = formula };
        }
        else
            tableColumn.CalculatedColumnFormula = null;
    }

    private static void ApplyColumnTotalsRow(TableColumn tableColumn, IXLTableField xlField)
    {
        if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
        {
            tableColumn.TotalsRowFunction = xlField.TotalsRowFunction.ToOpenXml();

            if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                tableColumn.TotalsRowFormula = new TotalsRowFormula(xlField.TotalsRowFormulaA1);
        }

        if (!string.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
            tableColumn.TotalsRowLabel = xlField.TotalsRowLabel;
    }

    private static TableStyleInfo BuildTableStyleInfo(XLTable xlTable)
    {
        var tableStyleInfo = new TableStyleInfo
        {
            ShowFirstColumn = xlTable.EmphasizeFirstColumn,
            ShowLastColumn = xlTable.EmphasizeLastColumn,
            ShowRowStripes = xlTable.ShowRowStripes,
            ShowColumnStripes = xlTable.ShowColumnStripes
        };

        if (xlTable.Theme != XLTableTheme.None)
            tableStyleInfo.Name = xlTable.Theme.Name;

        return tableStyleInfo;
    }

    private static void SetAutoFilterRange(XLTable xlTable)
    {
        if (xlTable.ShowTotalsRow)
        {
            xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
        }
        else
            xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);
    }

    private static string GetTableName(string originalTableName, SaveContext context)
    {
        var tableName = originalTableName.RemoveSpecialCharacters();
        var name = tableName;
        if (context.TableNames.Contains(name))
        {
            var i = 1;
            name = tableName + i.ToInvariantString();
            while (context.TableNames.Contains(name))
            {
                i++;
                name = tableName + i.ToInvariantString();
            }
        }

        context.TableNames.Add(name);
        return name;
    }
}
