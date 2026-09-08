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
    private const string RefError = "#REF!";

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
        ArgumentNullException.ThrowIfNull(value);

        var formula = value.TrimFormulaEqual();
        var rejection = RejectionOf(formula, paramName, out var references);
        if (rejection is not null)
        {
            if (!acceptUnusable)
                throw rejection;

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
            return new ExpressionParseException(failure.Message, failure);

        if (references.References.Count > 0)
        {
            // `[MS-XLSX] 2.2.2.5: The formula MUST NOT use the local-cell-reference production
            // rule.` Excel will refuse to load a workbook with such a defined name (e.g. `A1`).
            // In theory, defined name should support bang references as a replacement for local
            // references, but ClosedParser doesn't support it yet.
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

        var copiedFormula = FormulaTransformation.SafeModifyA1(_formula, sheet.Name, 1, 1, new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { sheet.Name, targetSheet.Name } },
            Tables = tableRenames,
        });
        var copiedName = new XLDefinedName(targetSheet.DefinedNames, Name, false, copiedFormula, Comment);
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

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        RenameFormulaSheet(oldSheetName, newSheetName);
    }

    internal void OnWorksheetDeleted(string worksheetName)
    {
        RenameFormulaSheet(worksheetName, null);
        DropSheetPrefixOfRefError(worksheetName);
    }

    /// <summary>
    /// A reference that a row or column deletion has already reduced to <c>#REF!</c> keeps its sheet
    /// prefix (<c>'Sheet 1'!#REF!</c>), which is what Excel does while the sheet still exists. The
    /// parser reports that prefix as part of an error node rather than as a sheet reference, so
    /// <see cref="RenameFormulaSheet"/> never sees it and the prefix would outlive the sheet it names.
    /// Excel treats a defined name pointing at an absent sheet as a broken file, so drop the prefix and
    /// leave the bare <c>#REF!</c> that the rest of the deleted-sheet handling produces.
    /// </summary>
    private void DropSheetPrefixOfRefError(string worksheetName)
    {
        var prefixedRefError = worksheetName.EscapeSheetName() + "!" + RefError;
        if (!_formula.Contains(prefixedRefError, StringComparison.OrdinalIgnoreCase))
            return;

        SetRefersToUnchecked(_formula.Replace(prefixedRefError, RefError, StringComparison.OrdinalIgnoreCase));
    }

    private void RenameFormulaSheet(string oldSheetName, string? newSheetName)
    {
        if (!_references.ContainsSheet(oldSheetName))
            return;

        var modified = FormulaTransformation.SafeModifyA1(_formula, newSheetName ?? string.Empty, 1, 1, new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { oldSheetName, newSheetName } }
        });

        SetRefersToUnchecked(modified);
    }

    private static string RangeToFixed(IXLRangeBase range)
    {
        return range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true);
    }
}
