using ClosedXML.Parser;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel.CalcEngine.Visitors;

/// <summary>
/// A factory to rename named reference object (sheets, tables ect.).
/// </summary>
internal sealed class RenameRefModVisitor : RefModVisitor
{
    private readonly Dictionary<string, string?>? _sheets;
    private readonly Dictionary<string, string>? _tables;

    /// <summary>
    /// A mapping of sheets, from old name (key) to a new name (value).
    /// The <c>null</c> value indicates sheet has been deleted.
    /// </summary>
    // Write-only (init) properties: intentional design for immutable configuration
#pragma warning disable S2376
    internal IReadOnlyDictionary<string, string?> Sheets
    {
        init => _sheets = value.ToDictionary(x => x.Key, x => x.Value, XLHelper.SheetComparer);
    }

    internal IReadOnlyDictionary<string, string> Tables
    {
        init => _tables = value.ToDictionary(x => x.Key, x => x.Value, XLHelper.NameComparer);
    }
#pragma warning restore S2376

    protected override string? ModifySheet(ModContext ctx, string sheetName)
    {
        if (_sheets is not null && _sheets.TryGetValue(sheetName, out var newName))
            return newName;

        return sheetName;
    }

    protected override string? ModifyTable(ModContext ctx, string table)
    {
        if (_tables is not null && _tables.TryGetValue(table, out var newName))
            return newName;

        return table;
    }
}
