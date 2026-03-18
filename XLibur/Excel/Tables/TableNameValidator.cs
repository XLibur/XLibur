using System.Linq;

namespace XLibur.Excel.Tables;

/// <summary>
/// Consolidates table name validation logic per the OOXML spec:
/// table names cannot contain spaces, cannot be cell addresses,
/// and must be unique across all defined names in the workbook.
/// </summary>
internal static class TableNameValidator
{
    /// <summary>
    /// Validates if a proposed table name is valid in the context of a specific worksheet's workbook.
    /// </summary>
    /// <param name="tableName">Proposed table name.</param>
    /// <param name="worksheet">The worksheet the table will live on.</param>
    /// <param name="message">Validation failure message, empty if valid.</param>
    /// <returns>True if the proposed table name is valid.</returns>
    public static bool IsValidTableName(string tableName, IXLWorksheet worksheet, out string message)
    {
        message = "";

        var existingTableNames = worksheet.Tables.Select(t => t.Name);

        if (!XLHelper.ValidateName("table", tableName, string.Empty, existingTableNames, out message))
            return false;

        if (tableName.Contains(' '))
        {
            message = "Table names cannot contain spaces.";
            return false;
        }

        if (XLHelper.IsValidA1Address(tableName) || XLHelper.IsValidRCAddress(tableName))
        {
            message = $"Table name cannot be a valid cell address '{tableName}'.";
            return false;
        }

        if (IsDefinedNameConflict(tableName, worksheet.Workbook))
        {
            message = $"Table name must be unique across all defined names '{tableName}'.";
            return false;
        }

        return true;
    }

    private static bool IsDefinedNameConflict(string tableName, XLWorkbook workbook)
    {
        return workbook.DefinedNames.Contains(tableName) ||
               workbook.Worksheets.Any(ws => ws.DefinedNames.Contains(tableName));
    }
}
