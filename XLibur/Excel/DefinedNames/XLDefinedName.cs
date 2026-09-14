using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;
using XLibur.Extensions;

namespace XLibur.Excel;

[DebuggerDisplay("{_name}:{_formula}")]
internal sealed class XLDefinedName : IXLDefinedName, IWorkbookListener
{
    private readonly XLDefinedNames _container;
    private string _name;
    private string _formula = null!;
    private FormulaReferences _references = null!;
    private bool _isFormulaUnderstood;

    internal XLDefinedName(XLDefinedNames container, string name, bool validateName, string formula, string? comment,
        bool acceptUnusableFormula = false)
    {
        // Excel accepts invalid names per grammar (e.g. `[Foo]Bar`) as a valid name, and they can be
        // encountered in existing workbooks. We shouldn't throw exception on a load.
        if (validateName && !XLHelper.ValidateName("named range", name, out var error))
            throw new ArgumentException(error, nameof(name));

        _container = container;
        _name = name;
        SetFormula(formula, nameof(formula), acceptUnusableFormula);
        Visible = true;
        Comment = comment;
    }

    public bool IsValid => _isFormulaUnderstood && !_references.ContainsRefError;

    /// <summary>
    /// Was the formula parsed and accepted? A name loaded from a workbook may carry text this library
    /// cannot read, or a local reference it will not resolve, and everything that rewrites a formula
    /// has to leave such text alone rather than guess at what the addresses inside it mean.
    /// </summary>
    internal bool IsFormulaUnderstood => _isFormulaUnderstood;

    /// <summary>
    /// Is this one of the hidden names Excel keeps a ChartEx chart's references in? Each <c>cx:f</c>
    /// of such a chart names one, <c>_xlchart.v1.1</c>, instead of holding a reference, and the name
    /// holds the reference (the <c>chartex-pivotcf-*</c> fixture).
    /// </summary>
    internal bool IsChartData => !Visible && _name.StartsWith("_xlchart.", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The area this name refers to on <paramref name="sheetName"/>, when the formula is that one
    /// reference and nothing else, so that <paramref name="rewrite"/>, the delete of the sheet, leaves
    /// nothing of it but <c>#REF!</c>.
    /// </summary>
    internal bool TryGetOnlyAreaOn(string sheetName, SheetRewrite rewrite, out Area area)
    {
        area = default;
        if (!_isFormulaUnderstood || _references.SheetReferences.Count != 1)
            return false;

        var reference = _references.SheetReferences.First();
        if (!XLHelper.SheetComparer.Equals(reference.Sheet, sheetName))
            return false;

        if (!rewrite.TryRewrite(_formula, string.Empty, new Point(1, 1), out var rewritten) || rewritten != "#REF!")
            return false;

        area = reference.Reference.ToSheetRange(new Point(1, 1));
        return true;
    }

    /// <summary>
    /// Does the formula reach a cell or area on some sheet? A name that does not — a constant, a
    /// structured reference, a bare <c>#REF!</c> — has nothing a row or column shift can move. The
    /// answer is cached from the parse that stored the formula, so asking costs nothing.
    /// </summary>
    internal bool HasSheetReferences => _references.SheetReferences.Count > 0;

    public string Name
    {
        get => _name;
        set
        {
            if (XLHelper.NameComparer.Equals(_name, value))
                return;

            if (!XLHelper.ValidateName("named range", value, out var error))
                throw new ArgumentException(error, nameof(value));

            if (_container.Contains(value))
                throw new InvalidOperationException($"There is already a name '{value}'.");

            _container.Delete(_name);
            _name = value;
            _container.Add(_name, this);
        }
    }

    public IXLRanges Ranges => _references.GetExternalRanges(_container.Workbook, new Point(1, 1));

    public string? Comment { get; set; }

    public bool Visible { get; set; }

    public XLNamedRangeScope Scope => _container.Scope;

    public string RefersTo
    {
        get => _formula;
        set => SetFormula(value, nameof(value), acceptUnusable: false);
    }

    /// <summary>
    /// Replaces the formula, keeping one this library will not work with instead of rejecting it.
    /// </summary>
    /// <remarks>
    /// Every internal rewrite goes through here: a rename, a sheet deletion and a row or column shift
    /// all edit a formula that is already on the name, and none of them may fail because the text they
    /// were handed was unusable when the workbook was opened.
    /// </remarks>
    internal void SetRefersToUnchecked(string formula)
        => SetFormula(formula, nameof(formula), acceptUnusable: true);

    /// <summary>
    /// Parses <paramref name="value"/> and stores it. When the formula turns out to be one this
    /// library will not work with, <paramref name="acceptUnusable"/> decides between keeping the text
    /// verbatim — the reader's choice, so that one bad name cannot stop a workbook from opening — and
    /// telling the caller its formula is bad.
    /// </summary>
    /// <param name="value">The formula to store.</param>
    /// <param name="paramName">
    /// The name <paramref name="value"/> has in the API the caller reached this through, so a rejection
    /// names the parameter that was actually passed rather than whatever it was assigned to on the way.
    /// </param>
    /// <param name="acceptUnusable">Whether to keep an unusable formula rather than reject it.</param>
    private void SetFormula(string value, string paramName, bool acceptUnusable)
    {
        // paramName, not the implicit "value": this rejection names the caller's parameter like the
        // ones below it, rather than the local it happens to have been assigned to.
        ArgumentNullException.ThrowIfNull(value, paramName);

        var formula = value.TrimFormulaEqual();
        var rejection = RejectionOf(formula, paramName, out var references);
        if (rejection is not null)
        {
            if (!acceptUnusable)
                throw rejection;

            // Leniency is for text that arrived unusable from a file. A formula this library *did*
            // understand and has now rewritten into one it does not means the rewrite is wrong, not
            // the input — and the name would go quiet rather than fail: no exception, empty Ranges,
            // and silent exclusion from every later shift and rename. Assert rather than swallow, for
            // the same reason XLCellFormulaShifter catches only ParsingException.
            Debug.Assert(!_isFormulaUnderstood,
                $"A defined name's formula was understood and has been rewritten into one that is not: '{formula}'.");

            // The text is kept so the name is written back as it was found, but nothing resolves or
            // rewrites it: whatever references it holds, this library did not accept the formula.
            _isFormulaUnderstood = false;
            _references = new FormulaReferences();
            _formula = formula;
            return;
        }

        _isFormulaUnderstood = true;
        _references = references;
        _formula = formula;
    }

    /// <summary>
    /// The exception refusing <paramref name="formula"/>, or <c>null</c> if this library will work
    /// with it.
    /// </summary>
    /// <remarks>
    /// The two refusals are different failures and answer with different types. Text the parser cannot
    /// read is an <see cref="ExpressionParseException"/>, which is what <see cref="IXLCell.FormulaA1"/>
    /// already raises for the same input, and it carries the parser's own exception so the position it
    /// reports survives. Text that parses but names a cell without a sheet was understood; the argument
    /// is what is at fault, so that stays an <see cref="ArgumentException"/>.
    /// </remarks>
    private static Exception? RejectionOf(string formula, string paramName, out FormulaReferences references)
    {
        if (!FormulaReferences.TryForFormula(formula, out references, out var failure))
            return failure;

        if (references.References.Count > 0)
        {
            // `[MS-XLSX] 2.2.2.5: The formula MUST NOT use the local-cell-reference production
            // rule.` Excel will refuse to load a workbook with such a defined name (e.g. `A1`).
            // A bang reference is the replacement Excel expects, and it gets past this guard: the
            // parser hands `!A1` to `IAstFactory.BangReference`, not `Reference`, so it is not
            // counted here. It evaluates against the sheet of the cell using the name, and so
            // does a bang name (`!Total`). See https://github.com/XLibur/XLibur/issues/446.
            return new ArgumentException($"Formula '{formula}' contains references without a sheet.", paramName);
        }

        return null;
    }

    IXLDefinedName IXLDefinedName.CopyTo(IXLWorksheet targetSheet) => CopyTo((XLWorksheet)targetSheet);

    void IXLDefinedName.Delete() => _container.Delete(Name);

    /// <summary>
    /// Try to resolve the first sheet reference in the formula to a worksheet and area.
    /// Avoids materializing <see cref="XLRange"/> or <see cref="XLRanges"/> objects.
    /// </summary>
    internal bool TryGetFirstSheetArea(XLWorkbook workbook, out XLWorksheet? sheet, out Area sheetArea)
    {
        var anchor = new Point(1, 1);
        foreach (var reference in _references.SheetReferences)
        {
            if (workbook.TryGetWorksheet(reference.Sheet, out sheet))
            {
                sheetArea = reference.Reference.ToSheetRange(anchor);
                return true;
            }
        }

        sheet = null;
        sheetArea = default;
        return false;
    }

    internal XLDefinedName CopyTo(XLWorksheet targetSheet)
    {
        var sheet = _container.Worksheet;
        if (targetSheet == sheet)
            throw new InvalidOperationException("Cannot copy named range to the worksheet it already belongs to.");

        if (sheet is null)
            throw new InvalidOperationException("Cannot copy workbook scoped defined name.");

        var targetTables = targetSheet.Tables.ToDictionary<XLTable, Area>(x => x.SheetRange);
        var tableRenames = new Dictionary<string, string>();
        foreach (var table in sheet.Tables)
        {
            if (targetTables.TryGetValue(table.SheetRange, out var targetTable))
            {
                tableRenames.Add(table.Name, targetTable.Name);
            }
        }

        // Only a formula this library accepted is re-pointed at the target sheet, just as rename, delete
        // and shift re-point only such a formula. Any other text is copied verbatim: the copy is as
        // broken as the original, which is the only honest answer for text whose meaning was never
        // established. That covers a formula the parser reads but this library rejects, such as one
        // with a reference that has no sheet, and a formula the parser refuses (ADR 0002).
        var modifier = new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { sheet.Name, targetSheet.Name } },
            Tables = tableRenames,
        };
        var copiedFormula = _isFormulaUnderstood
                            && FormulaText.TryRewrite(_formula, sheet.Name, new Point(1, 1), modifier,
                                out var rewritten, out _)
            ? rewritten
            : _formula;

        var copiedName = new XLDefinedName(targetSheet.DefinedNames, Name, false, copiedFormula, Comment,
            acceptUnusableFormula: !_isFormulaUnderstood);
        return targetSheet.DefinedNames.Add(Name, copiedName);
    }

    public IXLDefinedName SetRefersTo(IXLRangeBase range)
    {
        return SetRefersTo(RangeToFixed(range));
    }

    public IXLDefinedName SetRefersTo(IXLRanges ranges)
    {
        var unionFormula = string.Join(",", ranges.Select(RangeToFixed));
        return SetRefersTo(unionFormula);
    }

    public IXLDefinedName SetRefersTo(string formula)
    {
        SetFormula(formula, nameof(formula), acceptUnusable: false);
        return this;
    }

    public override string ToString()
    {
        return _formula;
    }

    internal void Add(string rangeAddress)
    {
        var byExclamation = rangeAddress.Split('!');
        var wsName = byExclamation[0].Replace("'", "");
        var rng = byExclamation[1];
        var rangeToAdd = _container.Workbook.WorksheetsInternal.Worksheet(wsName).Range(rng);

        var ranges = new XLRanges { rangeToAdd };
        RefersTo = _formula + "," + string.Join(",", ranges.Select(RangeToFixed));
    }

    /// <summary>
    /// The sheet is renamed wherever the formula names it: in a reference, a sheet-qualified name
    /// (<c>Data!Local</c>) and either end of a 3D reference. The <c>rename-*</c> fixture shows Excel
    /// renaming all of them, at every scope (D55).
    /// </summary>
    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        RewriteFormula(SheetRewrite.Rename(oldSheetName, newSheetName), newSheetName);
    }

    /// <summary>
    /// A reference to the deleted sheet becomes <c>#REF!</c>, at every scope, as Excel does. The
    /// <c>delete-*</c> fixture turns a name scoped to another sheet into <c>#REF!</c> (D54), and the
    /// <c>refdelete-*</c> fixture turns a reference a row delete had already broken, <c>Data!#REF!</c>,
    /// into a plain <c>#REF!</c> (D56). A 3D reference narrows instead, and a reference to a name that
    /// outlives the sheet points at that name in this workbook (see <see cref="SheetRewrite"/>).
    /// </summary>
    /// <remarks>
    /// A name scoped to the deleted sheet is rewritten too. It goes with its sheet, unless something
    /// outside the sheet refers to it; then it moves to workbook scope with the text this leaves it
    /// (see <see cref="XLDefinedNames.FindNamesOutlivingSheet"/>).
    /// </remarks>
    void IWorkbookListener.OnSheetDeleting(string sheetName)
    {
        RewriteFormula(SheetRewrite.Delete(_container.Workbook, sheetName), string.Empty);
    }

    /// <summary>
    /// Adds to <paramref name="found"/> each of <paramref name="names"/> that the formula refers to as a
    /// name scoped to <paramref name="sheetName"/>. A formula this library did not accept refers to
    /// nothing it can tell.
    /// </summary>
    internal void CollectReferencesToNames(string sheetName, IReadOnlySet<string> names, ISet<string> found)
    {
        if (_isFormulaUnderstood)
            SheetRewrite.CollectNamesReferredTo(_formula, sheetName, names, found);
    }

    /// <summary>
    /// Rewrites the formula the way <paramref name="rewrite"/> says. Only a formula this library
    /// accepted is rewritten. One the parser refuses keeps its text (ADR 0002), and so does one the
    /// parser reads but this library rejects: its references were never established, so they are not
    /// this library's to re-point.
    /// </summary>
    /// <param name="rewrite">What the rename or the delete does to formula text.</param>
    /// <param name="formulaSheetName">The sheet the parser reads the formula as being on.</param>
    private void RewriteFormula(SheetRewrite rewrite, string formulaSheetName)
    {
        if (!_isFormulaUnderstood)
            return;

        if (!rewrite.TryRewrite(_formula, formulaSheetName, new Point(1, 1), out var modified)
            || modified == _formula)
            return;

        SetRefersToUnchecked(modified);
    }

    private static string RangeToFixed(IXLRangeBase range)
    {
        return range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true);
    }
}
