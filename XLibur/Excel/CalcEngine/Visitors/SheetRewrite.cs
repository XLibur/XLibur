using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLibur.Parser;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine.Visitors;

/// <summary>
/// What a sheet rename or a sheet delete does to formula text that names the sheet. Every holder of
/// such text rewrites it through one of these, so a cell, a defined name, a conditional format, a
/// print area and a chart series all change the same way, and the way Excel changes them (spec 55).
/// </summary>
/// <remarks>
/// <para>
/// A rename renames the sheet wherever a formula names it. A delete does what the Excel-authored
/// fixtures show Excel doing:
/// </para>
/// <list type="bullet">
/// <item>A reference to the sheet becomes <c>#REF!</c>, and so does <c>Sheet!#REF!</c>.</item>
/// <item>A 3D reference with the sheet at one end narrows to the sheets that are left, by tab order
/// (see <see cref="RenameRefModVisitor"/>). Narrowed to one sheet, it is written as a plain sheet
/// reference, <c>SUM(Last!$A$1)</c> rather than <c>SUM(Last:Last!$A$1)</c>.</item>
/// <item>A reference to a name scoped to the sheet, where the name outlives the sheet (see
/// <see cref="XLDefinedNames.NamesOutlivingSheet"/>), becomes a reference to that name in this
/// workbook: <c>Data!Used</c> becomes <c>[0]!Used</c>.</item>
/// </list>
/// <para>
/// The parser's modifier can rename a sheet or give a part up as <c>#REF!</c>, but it cannot turn
/// one kind of reference into another. So the last two are written here, at the positions the
/// parser reports for text it has read. Text the parser refuses keeps its text (ADR 0002).
/// </para>
/// </remarks>
internal sealed class SheetRewrite
{
    private static readonly IReadOnlySet<string> NoNames = new HashSet<string>();

    private readonly XLWorkbook? _workbook;
    private readonly string _sheetName;
    private readonly string? _newSheetName;
    private readonly IReadOnlySet<string> _outlivingNames;
    private IReadOnlyList<string>? _tabOrder;

    private SheetRewrite(XLWorkbook? workbook, string sheetName, string? newSheetName,
        IReadOnlySet<string> outlivingNames)
    {
        _workbook = workbook;
        _sheetName = sheetName;
        _newSheetName = newSheetName;
        _outlivingNames = outlivingNames;
    }

    /// <summary>
    /// The rewrite for a rename of <paramref name="oldSheetName"/> to <paramref name="newSheetName"/>.
    /// </summary>
    internal static SheetRewrite Rename(string oldSheetName, string newSheetName)
        => new(null, oldSheetName, newSheetName, NoNames);

    /// <summary>
    /// The rewrite for the delete of <paramref name="sheetName"/> from <paramref name="workbook"/>.
    /// Made while the sheet is still in the workbook, as it is while
    /// <see cref="IWorkbookListener.OnSheetDeleting"/> is raised, so the tab order still has it.
    /// </summary>
    internal static SheetRewrite Delete(XLWorkbook workbook, string sheetName)
    {
        var outlivingNames = workbook.WorksheetsInternal.TryGetWorksheetByRawName(sheetName, out var sheet)
            ? sheet.DefinedNames.NamesOutlivingSheet
            : NoNames;
        return new SheetRewrite(workbook, sheetName, null, outlivingNames);
    }

    /// <summary>
    /// The workbook's sheets in tab order, read once and only when a formula names the deleted sheet.
    /// </summary>
    private IReadOnlyList<string> TabOrder
        => _tabOrder ??= _workbook is null
            ? []
            : _workbook.WorksheetsInternal.OrderBy<XLWorksheet, int>(s => s.Position).Select(s => s.Name).ToList();

    /// <summary>
    /// Rewrites <paramref name="text"/>.
    /// </summary>
    /// <param name="text">The formula text, in A1 notation, without a leading <c>=</c>.</param>
    /// <param name="formulaSheetName">The sheet the parser reads the formula as being on.</param>
    /// <param name="origin">The cell the formula is in.</param>
    /// <param name="rewritten">The rewritten text, or <paramref name="text"/> when it is refused.</param>
    /// <returns><c>false</c> when the parser refused the text, which is then kept as it is.</returns>
    internal bool TryRewrite(string text, string formulaSheetName, Point origin, out string rewritten)
    {
        rewritten = text;

        // A formula names a sheet only by writing its name, so text without the name has nothing
        // to rewrite and costs no parse. On most holders that is nearly every formula.
        if (!MentionsSheet(text))
            return true;

        var source = text;
        if (_outlivingNames.Count > 0 && !TryPointAtOutlivingNames(text, out source))
            return false;

        var modifier = new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { _sheetName, _newSheetName } },
            TabOrder = TabOrder,
        };
        if (!FormulaText.TryRewrite(source, formulaSheetName, origin, modifier, out var result, out _))
            return false;

        rewritten = modifier.NarrowedToOneSheet ? WriteOneSheetRangesAsSheets(result) : result;
        return true;
    }

    /// <summary>
    /// Adds to <paramref name="found"/> each of <paramref name="names"/> that <paramref name="text"/>
    /// refers to as a name scoped to <paramref name="sheetName"/>, which a formula writes
    /// <c>sheetName!Name</c>. Text the parser refuses refers to nothing anyone can tell.
    /// </summary>
    internal static void CollectNamesReferredTo(string text, string sheetName, IReadOnlySet<string> names,
        ISet<string> found)
    {
        if (!MentionsAny(text, names))
            return;

        var context = new ScopedNames(sheetName, names);
        if (!FormulaText.TryWalk(text, context, ScopedNameFinder.Instance, FormulaNotation.A1, out _, out _))
            return;

        foreach (var (_, name) in context.Found)
            found.Add(name);
    }

    private bool MentionsSheet(string text)
    {
        if (text.Contains(_sheetName, StringComparison.OrdinalIgnoreCase))
            return true;

        // Quoted, a name's apostrophes are doubled.
        return _sheetName.Contains('\'')
               && text.Contains(_sheetName.Replace("'", "''"), StringComparison.OrdinalIgnoreCase);
    }

    private static bool MentionsAny(string text, IReadOnlySet<string> names)
    {
        foreach (var name in names)
        {
            if (text.Contains(name, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Writes each reference to a name that outlives the deleted sheet as a reference to the name in
    /// this workbook, <c>[0]!Name</c>, as Excel does. Done before the parser's rewrite, which would
    /// otherwise give <c>Data!Used</c> up as <c>#REF!</c> like any other reference to the sheet.
    /// </summary>
    private bool TryPointAtOutlivingNames(string text, out string pointed)
    {
        pointed = text;
        if (!MentionsAny(text, _outlivingNames))
            return true;

        var context = new ScopedNames(_sheetName, _outlivingNames);
        if (!FormulaText.TryWalk(text, context, ScopedNameFinder.Instance, FormulaNotation.A1, out _, out _))
            return false;

        pointed = Splice(text, context.Found.Select(f => (f.Range, "[0]!" + f.Name)));
        return true;
    }

    /// <summary>
    /// Writes each 3D reference whose ends are one sheet as a reference to that sheet. The text is
    /// the parser's own output, so it reads back; were it not to, the range of one sheet is kept,
    /// which means the same.
    /// </summary>
    private static string WriteOneSheetRangesAsSheets(string text)
    {
        var found = new List<(SymbolRange Range, string Sheet, ReferenceArea Reference)>();
        if (!FormulaText.TryWalk(text, found, OneSheetRangeFinder.Instance, FormulaNotation.A1, out _, out _))
            return text;

        return Splice(text, found.Select(f => (f.Range, SheetPrefix(f.Sheet) + f.Reference.GetDisplayStringA1())));
    }

    /// <summary>
    /// A sheet prefix quoted the way the parser quotes one: only when the name needs it, with its
    /// apostrophes doubled.
    /// </summary>
    private static string SheetPrefix(string sheet)
        => NameUtils.ShouldQuote(sheet.AsSpan())
            ? "'" + sheet.Replace("'", "''") + "'!"
            : sheet + "!";

    /// <summary>
    /// Replaces each range of <paramref name="text"/> with its text, from the last to the first, so
    /// that every range still indexes the text it came from.
    /// </summary>
    private static string Splice(string text, IEnumerable<(SymbolRange Range, string Text)> replacements)
    {
        var ordered = replacements.OrderByDescending(r => r.Range.Start).ToList();
        if (ordered.Count == 0)
            return text;

        var sb = new StringBuilder(text);
        foreach (var (range, replacement) in ordered)
        {
            var length = Math.Min(range.End, text.Length) - range.Start;
            sb.Remove(range.Start, length).Insert(range.Start, replacement);
        }

        return sb.ToString();
    }

    /// <summary>The sheet and the names to find, and where each was found.</summary>
    private sealed class ScopedNames(string sheet, IReadOnlySet<string> names)
    {
        internal string Sheet { get; } = sheet;

        internal IReadOnlySet<string> Names { get; } = names;

        internal List<(SymbolRange Range, string Name)> Found { get; } = [];
    }

    /// <summary>Finds each <c>Sheet!Name</c> that names one of the names looked for.</summary>
    private sealed class ScopedNameFinder : CollectVisitor<ScopedNames>
    {
        internal static readonly ScopedNameFinder Instance = new();

        public override object? SheetName(ScopedNames context, SymbolRange range, string sheet, string name)
        {
            if (XLHelper.SheetComparer.Equals(sheet, context.Sheet) && context.Names.Contains(name))
                context.Found.Add((range, name));

            return default;
        }
    }

    /// <summary>Finds each 3D reference whose first and last sheets are one sheet.</summary>
    private sealed class OneSheetRangeFinder
        : CollectVisitor<List<(SymbolRange Range, string Sheet, ReferenceArea Reference)>>
    {
        internal static readonly OneSheetRangeFinder Instance = new();

        public override object? Reference3D(List<(SymbolRange Range, string Sheet, ReferenceArea Reference)> context,
            SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
        {
            if (XLHelper.SheetComparer.Equals(firstSheet, lastSheet))
                context.Add((range, firstSheet, reference));

            return default;
        }
    }
}
