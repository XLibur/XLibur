using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

internal sealed class XLPrintAreas : IXLPrintAreas, IWorkbookListener
{
    readonly List<IXLRange> ranges = new List<IXLRange>();
    private readonly XLWorksheet worksheet;

    /// <summary>
    /// Stores the raw defined name text when the print area contains a formula
    /// (e.g. OFFSET) that cannot be resolved to a simple range reference.
    /// When set, this value is written as-is to the workbook on save.
    /// </summary>
    internal string? FormulaReference { get; set; }

    public XLPrintAreas(XLWorksheet worksheet)
    {
        this.worksheet = worksheet;
    }

    public XLPrintAreas(XLPrintAreas defaultPrintAreas, XLWorksheet worksheet)
    {
        ranges = defaultPrintAreas.ranges.ToList();
        FormulaReference = defaultPrintAreas.FormulaReference;
        this.worksheet = worksheet;
    }

    public void Clear()
    {
        ranges.Clear();
    }

    public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        ranges.Add(worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
    }

    public void Add(string rangeAddress)
    {
        ranges.Add(worksheet.Range(rangeAddress)!);
    }

    public void Add(string firstCellAddress, string lastCellAddress)
    {
        ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
    {
        ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public IEnumerator<IXLRange> GetEnumerator()
    {
        return ranges.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// A print area kept as formula text names the renamed sheet by its new name, as Excel writes it:
    /// the <c>rename-*</c> fixture turns <c>OFFSET(Data!$A$1,0,0,4,2)</c> into
    /// <c>OFFSET(Renamed!$A$1,0,0,4,2)</c> (D68). A print area held as ranges is live and needs nothing.
    /// </summary>
    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        => RewriteFormulaReference(SheetRewrite.Rename(oldSheetName, newSheetName));

    /// <summary>
    /// A print area goes with its own sheet, as the <c>delete-*</c> fixture shows. On any other sheet a
    /// print area is a defined name scoped to that sheet, and changes as one does: a reference to the
    /// deleted sheet becomes <c>#REF!</c> (D54). No fixture holds such a print area; this follows the
    /// name it is in the file.
    /// </summary>
    void IWorkbookListener.OnSheetDeleting(string sheetName)
    {
        if (XLHelper.SheetComparer.Equals(sheetName, worksheet.Name))
            return;

        RewriteFormulaReference(SheetRewrite.Delete(worksheet.Workbook, sheetName));
    }

    /// <remarks>A formula the parser refuses keeps its text (ADR 0002).</remarks>
    private void RewriteFormulaReference(SheetRewrite rewrite)
    {
        if (FormulaReference is { } formula
            && rewrite.TryRewrite(formula, worksheet.Name, new Point(1, 1), out var rewritten))
            FormulaReference = rewritten;
    }
}
