using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XLibur.Excel.Tables;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLWorksheets : IXLWorksheets, IEnumerable<XLWorksheet>
{
    private readonly XLWorkbook _workbook;
    private readonly Dictionary<string, XLWorksheet> _worksheets = new(StringComparer.OrdinalIgnoreCase);
    internal ICollection<string> Deleted { get; private set; }

    /// <summary>
    /// SheetId that will be assigned to the next created sheet.
    /// </summary>
    private uint _nextSheetId = 1;

    #region Constructor

    public XLWorksheets(XLWorkbook workbook)
    {
        _workbook = workbook;
        Deleted = new HashSet<string>();
    }

    #endregion Constructor

    #region IEnumerable<XLWorksheet> Members

    public IEnumerator<XLWorksheet> GetEnumerator()
    {
        return ((IEnumerable<XLWorksheet>)_worksheets.Values).GetEnumerator();
    }

    #endregion IEnumerable<XLWorksheet> Members

    #region IXLWorksheets Members

    public int Count
    {
        [DebuggerStepThrough]
        get => _worksheets.Count;
    }

    public bool Contains(string sheetName)
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        return _worksheets.ContainsKey(sheetName);
    }

    bool IXLWorksheets.TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXLWorksheet? worksheet)
    {
        if (TryGetWorksheet(sheetName, out var foundSheet))
        {
            worksheet = foundSheet;
            return true;
        }

        worksheet = null;
        return false;
    }

    internal bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out XLWorksheet? worksheet)
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        if (_worksheets.TryGetValue(sheetName.UnescapeSheetName(), out worksheet))
        {
            return true;
        }

        worksheet = null;
        return false;
    }

    /// <summary>
    /// Find a sheet by the name it actually has, without treating the name as formula syntax.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="TryGetWorksheet(string,out XLWorksheet)"/> and
    /// <see cref="Worksheet(string)"/> both call <c>UnescapeSheetName</c>, which strips surrounding
    /// apostrophes and collapses <c>''</c> to <c>'</c>. That is right for a name lifted out of a
    /// formula, where <c>'It''s'!A1</c> refers to the sheet <c>It's</c>. It is wrong for a name read
    /// from the <c>name</c> attribute of <c>&lt;sheet&gt;</c>, which is the name itself: Excel
    /// permits an apostrophe anywhere but the first and last character, so <c>Ann''s</c> is a legal
    /// sheet name that unescaping turns into <c>Ann's</c> — a sheet that does not exist.
    /// </para>
    /// <para>
    /// The loader used the unescaping lookups on raw names and had two symptoms for it (D40): the
    /// pass that reads sheet contents took its "this shouldn't be possible" branch and skipped the
    /// sheet, loading it empty and silently; the pass that reads pivot tables threw
    /// <see cref="ArgumentException"/> out of <c>new XLWorkbook(stream)</c>. Use this from anything
    /// holding a name that came from the file rather than from a formula.
    /// </para>
    /// </remarks>
    internal bool TryGetWorksheetByRawName(string sheetName, [NotNullWhen(true)] out XLWorksheet? worksheet)
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        return _worksheets.TryGetValue(sheetName, out worksheet);
    }

    public IXLWorksheet Worksheet(string sheetName)
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        sheetName = sheetName.UnescapeSheetName();

        if (_worksheets.TryGetValue(sheetName, out var w))
            return w;

        throw new ArgumentException("There isn't a worksheet named '" + sheetName + "'.");
    }

    public IXLWorksheet Worksheet(int position)
    {
        var wsCount = _worksheets.Values.Count(w => w.Position == position);
        return wsCount switch
        {
            0 => throw new ArgumentException("There isn't a worksheet associated with that position."),
            > 1 => throw new ArgumentException(
                "Can't retrieve a worksheet because there are multiple worksheets associated with that position."),
            _ => _worksheets.Values.Single(w => w.Position == position)
        };
    }

    public IXLWorksheet Add()
    {
        return Add(GetNextWorksheetName());
    }

    public IXLWorksheet Add(int position)
    {
        return Add(GetNextWorksheetName(), position);
    }

    public IXLWorksheet Add(string sheetName)
    {
        EnsureNameIsFree(sheetName, nameof(sheetName));
        var sheet = new XLWorksheet(sheetName, _workbook, GetNextSheetId());
        Add(sheetName, sheet);
        sheet._position = _worksheets.Count + _workbook.UnsupportedSheets.Count;
        return sheet;
    }

    public IXLWorksheet Add(string sheetName, int position)
    {
        return Add(sheetName, position, GetNextSheetId());
    }

    internal XLWorksheet Add(string sheetName, int position, uint sheetId)
    {
        // Before anything moves: a refused name must leave the tab order as it was.
        EnsureNameIsFree(sheetName, nameof(sheetName));

        _worksheets.Values.Where(w => w._position >= position).ForEach(w => w._position += 1);
        _workbook.UnsupportedSheets.Where(w => w.Position >= position).ForEach(w => w.Position += 1);

        // If the loaded sheetId is greater than current, just make sure our next sheetId is even bigger.
        _nextSheetId = Math.Max(_nextSheetId, sheetId + 1);
        var sheet = new XLWorksheet(sheetName, _workbook, sheetId);
        Add(sheetName, sheet);
        sheet._position = position;
        return sheet;
    }

    private void Add(string sheetName, XLWorksheet sheet)
    {
        if (!_worksheets.TryAdd(sheetName, sheet))
            throw new ArgumentException($"A worksheet with the same name ({sheetName}) has already been added.", nameof(sheetName));

        _workbook.NotifyWorksheetAdded(sheet);
    }

    public IXLWorksheet Add(DataTable dataTable)
    {
        return Add(dataTable, dataTable.TableName);
    }

    public IXLWorksheet Add(DataTable dataTable, string sheetName)
    {
        return Add(dataTable, sheetName, TableNameGenerator.GetNewTableName(_workbook));
    }

    public IXLWorksheet Add(DataTable dataTable, string sheetName, string tableName)
    {
        var ws = Add(sheetName);
        ws.Cell(1, 1).InsertTable(dataTable, tableName);
        return ws;
    }

    public void Add(DataSet dataSet)
    {
        foreach (DataTable t in dataSet.Tables)
            Add(t);
    }

    public void Delete(string sheetName)
    {
        ArgumentException.ThrowIfNullOrEmpty(sheetName);
        Delete(_worksheets[sheetName].Position);
    }

    /// <summary>
    /// Deletes <paramref name="sheet"/> itself, the sheet <see cref="IXLWorksheet.Delete"/> was called
    /// on. A sheet that is already deleted is left as it is, as a rename of one changes nothing in the
    /// workbook. Looking the sheet up by its name instead would find a sheet added since under that
    /// name, and delete that one.
    /// </summary>
    internal void Delete(XLWorksheet sheet)
    {
        if (IsRegistered(sheet))
            Delete(sheet.Position);
    }

    /// <summary>
    /// The one implementation of deleting a sheet. <see cref="IXLWorksheet.Delete"/> and
    /// <see cref="Delete(string)"/> both come here, so every way of deleting a sheet does the same.
    /// </summary>
    public void Delete(int position)
    {
        var wsCount = _worksheets.Values.Count(w => w.Position == position);
        switch (wsCount)
        {
            case 0:
                throw new ArgumentException("There isn't a worksheet associated with that index.");
            case > 1:
                throw new ArgumentException(
                    "Can't delete the worksheet because there are multiple worksheets associated with that index.");
        }

        var ws = _worksheets.Values.Single(w => w.Position == position);

        // 1. Find the names scoped to the sheet that outlive it, before any holder has rewritten a
        //    reference to them, since what counts as a reference to such a name is what the holders
        //    rewrite. Find, too, the hidden names of ChartEx charts' references that go with the sheet,
        //    which the charts read.
        ws.DefinedNames.FindNamesOutlivingSheet();
        _workbook.DefinedNamesInternal.FindChartDataNamesGoingWithSheet(ws.Name);

        // 2. Every holder hears of it, while the sheet can still be resolved. A listener does not
        //    throw (see IWorkbookListener), so nothing here catches.
        foreach (var listener in GetWorkbookListeners())
            listener.OnSheetDeleting(ws.Name);

        // 3. A range or an address that outlives the sheet reads #REF! from now on.
        ws.IsDeleted = true;

        // 4. Remove the sheet, and close the gap it leaves in the tab order.
        if (!string.IsNullOrWhiteSpace(ws.RelId) && !Deleted.Contains(ws.RelId))
            Deleted.Add(ws.RelId);

        _worksheets.RemoveAll(w => w.Position == position);
        _worksheets.Values.Where(w => w.Position > position).ForEach(w => w._position -= 1);
        _workbook.UnsupportedSheets.Where(w => w.Position > position).ForEach(w => w.Position -= 1);

        // 5. The names that outlive the sheet move to workbook scope, as Excel moves them. The rest of
        //    the sheet's names go with it, and so do the hidden names of ChartEx references to it.
        foreach (var name in ws.DefinedNames.NamesOutlivingSheet)
            _workbook.DefinedNamesInternal.AdoptFromDeletedSheet(ws.DefinedNames.DefinedName(name));

        _workbook.DefinedNamesInternal.RemoveChartDataNamesGoingWithSheet();

        // 6. Dispose what the sheet held.
        ws.Cleanup();
    }

    IEnumerator<IXLWorksheet> IEnumerable<IXLWorksheet>.GetEnumerator()
    {
        return _worksheets.Values.Cast<IXLWorksheet>().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion IXLWorksheets Members

    /// <summary>
    /// The one implementation of renaming a sheet. The setter of <see cref="XLWorksheet.Name"/>
    /// delegates here entirely, and passes the field that holds the sheet's name as
    /// <c>sheetName</c>, which only this method writes. The collection's key and the sheet's name
    /// change together, and only then does any listener hear of it.
    /// </summary>
    internal void Rename(XLWorksheet sheet, string newSheetName, ref string sheetName)
    {
        var oldSheetName = sheetName;
        if (oldSheetName == newSheetName)
            return;

        XLHelper.ValidateSheetName(newSheetName);

        // A deleted sheet is not in the collection, and nothing holds its name, so a rename only
        // changes what it is called. Looking the sheet up by its old name instead would find a sheet
        // added since under that name, and change that sheet's key behind its back.
        if (!IsRegistered(sheet))
        {
            sheetName = newSheetName;
            return;
        }

        if (!XLHelper.SheetComparer.Equals(oldSheetName, newSheetName))
            EnsureNameIsFree(newSheetName, nameof(newSheetName));

        _worksheets.Remove(oldSheetName);
        sheetName = newSheetName;
        Add(newSheetName, sheet);

        foreach (var listener in GetWorkbookListeners())
            listener.OnSheetRenamed(oldSheetName, newSheetName);
    }

    /// <summary>
    /// Every component that holds text naming a sheet, in the order it hears of a rename or a
    /// delete. <c>SheetLifecycleTests</c> pins the order.
    /// </summary>
    /// <remarks>
    /// <para>
    /// One order serves both events. The calc engine comes first: on a rename it renames its
    /// dependency tree, and on a delete it drops the tree and marks every formula dirty. Neither
    /// reads the formula text that the holders after it change, so the engine and the holders
    /// commute on both events.
    /// </para>
    /// <para>
    /// Each holder changes only its own text, and none of them reads another's, so the holders
    /// commute with each other too. The one thing a delete decides across holders, which of the
    /// sheet's names outlive it, is decided before any of them hears of it (see
    /// <see cref="Delete(int)"/>).
    /// </para>
    /// <para>
    /// Hyperlinks are not here. Excel leaves an internal link's location as it is on a rename and on
    /// a delete (the <c>rename-*</c> and <c>delete-*</c> fixtures), and so does XLibur.
    /// </para>
    /// </remarks>
    internal IEnumerable<IWorkbookListener> GetWorkbookListeners()
    {
        yield return _workbook.CalcEngine;

        foreach (var sheet in _worksheets.Values)
        {
            yield return sheet.Internals.CellsCollection;
        }

        foreach (var definedName in _workbook.DefinedNamesInternal)
            yield return definedName;

        foreach (var sheet in _worksheets.Values)
        {
            foreach (var definedName in sheet.DefinedNames)
            {
                yield return definedName;
            }
        }

        foreach (var sheet in _worksheets.Values)
        {
            yield return sheet.ConditionalFormats;
            yield return (XLPrintAreas)sheet.PageSetup.PrintAreas;
            yield return (XLCharts)sheet.Charts;
        }

        yield return _workbook.PivotCachesInternal;
    }

    #region Private members

    private string GetNextWorksheetName()
    {
        var worksheetNumber = Count + 1;
        var sheetName = $"Sheet{worksheetNumber}";
        while (_worksheets.ContainsKey(sheetName) || IsHeldByUnsupportedSheet(sheetName))
        {
            worksheetNumber++;
            sheetName = $"Sheet{worksheetNumber}";
        }
        return sheetName;
    }

    /// <summary>
    /// A <c>sheetId</c> that no sheet in the workbook has had in this session, counting the sheets
    /// XLibur keeps but does not model.
    /// </summary>
    /// <remarks>
    /// An unsupported sheet, such as a chartsheet, keeps the id it was loaded with, and a save writes
    /// its <c>&lt;sheet&gt;</c> element back with that id. The loader moves <see cref="_nextSheetId"/>
    /// past each worksheet it adds, but not past those sheets. So a new worksheet took the id of a
    /// chartsheet that had the highest one. The writer matches <c>&lt;sheet&gt;</c> elements by
    /// <c>sheetId</c>, so it gave the new worksheet the chartsheet's <c>r:id</c>, and the save threw
    /// (D76). The ids are read from <c>UnsupportedSheets</c> itself, so no second record of those
    /// sheets has to be kept in step with it. Ids only go up, so a deleted sheet's id is not used
    /// again in the same session.
    /// </remarks>
    private uint GetNextSheetId()
    {
        foreach (var unsupportedSheet in _workbook.UnsupportedSheets)
            _nextSheetId = Math.Max(_nextSheetId, unsupportedSheet.SheetId + 1);

        return _nextSheetId++;
    }

    /// <summary>
    /// Is <paramref name="sheet"/> the sheet this collection holds under its name? A deleted sheet is
    /// not, and neither is it when a sheet has been added since under the deleted one's name.
    /// </summary>
    private bool IsRegistered(XLWorksheet sheet)
        => _worksheets.TryGetValue(sheet.Name, out var current) && ReferenceEquals(current, sheet);

    /// <summary>
    /// Refuses a name another sheet already has: a modelled sheet, or an unsupported sheet such as a
    /// chartsheet, which XLibur keeps and writes back although it does not model it. A file holding
    /// both would declare the name twice, and XLibur refuses to load such a file (D34). Names are
    /// compared as sheet names are everywhere else, ignoring case.
    /// </summary>
    private void EnsureNameIsFree(string sheetName, string paramName)
    {
        if (_worksheets.ContainsKey(sheetName))
            throw new ArgumentException($"A worksheet with the same name ({sheetName}) has already been added.", paramName);

        if (IsHeldByUnsupportedSheet(sheetName))
        {
            throw new ArgumentException(
                $"The workbook already has a sheet named '{sheetName}' that XLibur keeps but does not model, such as a chartsheet.",
                paramName);
        }
    }

    private bool IsHeldByUnsupportedSheet(string sheetName)
        => _workbook.UnsupportedSheets.Exists(s => XLHelper.SheetComparer.Equals(s.Name, sheetName));

    #endregion Private members
}
