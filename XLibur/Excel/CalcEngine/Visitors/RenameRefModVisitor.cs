using System.Collections.Generic;
using System.Linq;
using ClosedXML.Parser;

namespace XLibur.Excel.CalcEngine.Visitors;

/// <summary>
/// A factory to rename named reference object (sheets, tables ect.).
/// </summary>
internal sealed class RenameRefModVisitor : FormulaModifier
{
    private readonly Dictionary<string, string?>? _sheets;
    private readonly Dictionary<string, string>? _tables;
    private readonly IReadOnlyList<string>? _tabOrder;

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

    /// <summary>
    /// The workbook's sheets in tab order, a deleted sheet still among them. A 3D reference with a
    /// deleted sheet at one end narrows by it (see <see cref="ModifySheetRange"/>). Without it, such
    /// a reference becomes <c>#REF!</c>, as the parser's default has it.
    /// </summary>
    internal IReadOnlyList<string> TabOrder
    {
        init => _tabOrder = value;
    }
#pragma warning restore S2376

    /// <summary>
    /// Did a 3D reference narrow to a single sheet? The parser still writes it as a range of one
    /// sheet, <c>Last:Last!A1</c>, where Excel writes <c>Last!A1</c>, so the caller writes it again
    /// (see <see cref="SheetRewrite"/>).
    /// </summary>
    internal bool NarrowedToOneSheet { get; private set; }

    protected override string? ModifySheet(ModContext ctx, string sheetName)
    {
        if (_sheets is not null && _sheets.TryGetValue(sheetName, out var newName))
            return newName;

        return sheetName;
    }

    /// <summary>
    /// A 3D reference with a deleted sheet at one end narrows to the sheets left between its ends, as
    /// Excel does (spec 55, Q34): deleting <c>Sheet1</c> turns <c>SUM(Sheet1:Sheet3!A1)</c> into
    /// <c>SUM(Sheet2:Sheet3!A1)</c>. The sheet next to the deleted end in tab order, on the side of the
    /// other end, takes its place. A rename renames each end, as the default does.
    /// </summary>
    /// <remarks>
    /// The <c>delete-*</c> fixture shows Excel narrowing through two deletes down to one sheet:
    /// <c>SUM(First:Last!$A$1)</c> became <c>SUM(Last!$A$1)</c> once <c>First</c> and then <c>Data</c>
    /// were gone. A reference with the deleted sheet at both ends has nothing left, and one whose
    /// other end names no sheet of the workbook cannot be narrowed. Both become <c>#REF!</c>, as the
    /// default has it.
    /// </remarks>
    protected override SheetRange? ModifySheetRange(ModContext ctx, string firstSheet, string lastSheet)
    {
        var firstDeleted = IsDeleted(firstSheet);
        var lastDeleted = IsDeleted(lastSheet);
        if (firstDeleted == lastDeleted || _tabOrder is null)
            return base.ModifySheetRange(ctx, firstSheet, lastSheet);

        var firstIndex = IndexInTabOrder(firstSheet);
        var lastIndex = IndexInTabOrder(lastSheet);
        if (firstIndex < 0 || lastIndex < 0)
            return base.ModifySheetRange(ctx, firstSheet, lastSheet);

        // A reference can name its ends in either order, so "inwards" is whichever way the other
        // end lies. The ends differ, because only one of them is the deleted sheet.
        var inwards = firstIndex < lastIndex ? 1 : -1;
        var narrowed = firstDeleted
            ? new SheetRange(_tabOrder[firstIndex + inwards], lastSheet)
            : new SheetRange(firstSheet, _tabOrder[lastIndex - inwards]);

        if (XLHelper.SheetComparer.Equals(narrowed.FirstSheet, narrowed.LastSheet))
            NarrowedToOneSheet = true;

        return narrowed;
    }

    private bool IsDeleted(string sheetName)
        => _sheets is not null && _sheets.TryGetValue(sheetName, out var newName) && newName is null;

    private int IndexInTabOrder(string sheetName)
    {
        for (var i = 0; i < _tabOrder!.Count; i++)
        {
            if (XLHelper.SheetComparer.Equals(_tabOrder[i], sheetName))
                return i;
        }

        return -1;
    }

    protected override string? ModifyTable(ModContext ctx, string table)
    {
        if (_tables is not null && _tables.TryGetValue(table, out var newName))
            return newName;

        return table;
    }
}
