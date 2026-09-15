using System;
using ClosedXML.Parser;

namespace XLibur.Excel.CalcEngine.Visitors;

/// <summary>
/// Finds whether formula text refers to a sheet other than the one the formula is on, through a
/// reference that names the sheet: <c>Data!$A$1</c>, <c>Data:Other!$A$1</c> or <c>Data!#REF!</c>.
/// </summary>
/// <remarks>
/// These do not refer to another sheet: a reference with no sheet name, which is to the formula's own
/// sheet; a defined name, also one scoped to a sheet (<c>Data!Items</c>); a reference to another
/// workbook; text in a string, as in <c>INDIRECT("Data!A1")</c>; and a bare <c>#REF!</c>, which names no
/// sheet.
/// </remarks>
internal sealed class OtherSheetReferenceFinder : CollectVisitor<OtherSheetReferenceFinder.Search>
{
    private static readonly OtherSheetReferenceFinder Instance = new();

    private OtherSheetReferenceFinder()
    {
    }

    /// <summary>
    /// Whether <paramref name="formula"/> refers to a sheet other than <paramref name="sheetName"/>.
    /// Text the parser refuses refers to nothing anyone can tell, so the answer for it is <c>false</c>.
    /// </summary>
    /// <param name="formula">The formula text, in A1 notation, without a leading <c>=</c>.</param>
    /// <param name="sheetName">The sheet the formula is on.</param>
    internal static bool RefersToAnotherSheet(string formula, string sheetName)
    {
        // A reference names a sheet only before a '!', so text without one costs no parse.
        if (!formula.Contains('!'))
            return false;

        var search = new Search(sheetName);
        return FormulaText.TryWalk(formula, search, Instance, FormulaNotation.A1, out _, out _) && search.Found;
    }

    public override object? SheetReference(Search context, SymbolRange range, string sheet, ReferenceArea reference)
    {
        context.Check(sheet);
        return default;
    }

    public override object? Reference3D(Search context, SymbolRange range, string firstSheet, string lastSheet,
        ReferenceArea reference)
    {
        context.Check(firstSheet);
        context.Check(lastSheet);
        return default;
    }

    public override object? SheetErrorNode(Search context, SymbolRange range, int? workbookIndex, string sheet,
        ReadOnlySpan<char> error)
    {
        // With a workbook index, the sheet is in another workbook.
        if (workbookIndex is null)
            context.Check(sheet);

        return default;
    }

    /// <summary>The sheet the formula is on, and whether a reference names another.</summary>
    internal sealed class Search(string sheetName)
    {
        internal bool Found { get; private set; }

        internal void Check(string sheet)
        {
            if (!XLHelper.SheetComparer.Equals(sheet, sheetName))
                Found = true;
        }
    }
}
