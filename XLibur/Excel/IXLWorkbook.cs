#nullable disable


using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Excel;

public interface IXLWorkbook : IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>, IDisposable
{
    string Author { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's calculation mode.
    /// </summary>
    XLCalculateMode CalculateMode { get; set; }

    bool CalculationOnSave { get; set; }

    /// <summary>
    ///   Gets or sets the default column width for the workbook.
    ///   <para>All new worksheets will use this column width.</para>
    /// </summary>
    double ColumnWidth { get; set; }

    IXLCustomProperties CustomProperties { get; }

    bool DefaultRightToLeft { get; }

    bool DefaultShowFormulas { get; }

    bool DefaultShowGridLines { get; }

    bool DefaultShowOutlineSymbols { get; }

    bool DefaultShowRowColHeaders { get; }

    bool DefaultShowRuler { get; }

    bool DefaultShowWhiteSpace { get; }

    bool DefaultShowZeros { get; }

    IXLFileSharing FileSharing { get; }

    bool ForceFullCalculation { get; set; }

    bool FullCalculationOnLoad { get; set; }

    bool FullPrecision { get; set; }

    bool LockStructure { get; set; }

    bool LockWindows { get; set; }

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    IXLDefinedNames NamedRanges { get; }

    /// <summary>
    ///   Gets an object to manipulate this workbook's defined names.
    /// </summary>
    IXLDefinedNames DefinedNames { get; }

    /// <summary>
    ///   Gets or sets the default outline options for the workbook.
    ///   <para>All new worksheets will use these outline options.</para>
    /// </summary>
    IXLOutline Outline { get; set; }

    /// <summary>
    ///   Gets or sets the default page options for the workbook.
    ///   <para>All new worksheets will use these page options.</para>
    /// </summary>
    IXLPageSetup PageOptions { get; set; }

    /// <summary>
    ///   Gets all pivot caches in a workbook. A one cache can be
    ///   used by multiple tables. Unused caches are not saved.
    /// </summary>
    IXLPivotCaches PivotCaches { get; }

    /// <summary>
    ///   Gets or sets the workbook's properties.
    /// </summary>
    XLWorkbookProperties Properties { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's reference style.
    /// </summary>
    XLReferenceStyle ReferenceStyle { get; set; }

    bool RightToLeft { get; set; }

    /// <summary>
    ///   Gets or sets the default row height for the workbook.
    ///   <para>All new worksheets will use this row height.</para>
    /// </summary>
    double RowHeight { get; set; }

    bool ShowFormulas { get; set; }

    bool ShowGridLines { get; set; }

    bool ShowOutlineSymbols { get; set; }

    bool ShowRowColHeaders { get; set; }

    bool ShowRuler { get; set; }

    bool ShowWhiteSpace { get; set; }

    bool ShowZeros { get; set; }

    /// <summary>
    ///   Gets or sets the default style for the workbook.
    ///   <para>All new worksheets will use this style.</para>
    /// </summary>
    IXLStyle Style { get; set; }

    /// <summary>
    ///   Gets an object to manipulate this workbook's theme.
    /// </summary>
    IXLTheme Theme { get; }

    bool Use1904DateSystem { get; set; }

    /// <summary>
    ///   Gets an object to manipulate the worksheets.
    /// </summary>
    IXLWorksheets Worksheets { get; }

    IXLWorksheet AddWorksheet();

    IXLWorksheet AddWorksheet(int position);

    IXLWorksheet AddWorksheet(string sheetName);

    IXLWorksheet AddWorksheet(string sheetName, int position);

    void AddWorksheet(DataSet dataSet);

    void AddWorksheet(IXLWorksheet worksheet);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1). The dataTable's name is used for the
    /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
    /// </summary>
    /// <param name="dataTable">Datatable to insert</param>
    /// <returns>Inserted Worksheet</returns>
    IXLWorksheet AddWorksheet(DataTable dataTable);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1). The sheetName provided is used for the
    /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
    /// </summary>
    /// <param name="dataTable">dataTable to insert as Excel Table</param>
    /// <param name="sheetName">Worksheet and Excel Table name</param>
    /// <returns>Inserted Worksheet</returns>
    IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1).
    /// </summary>
    /// <param name="dataTable">dataTable to insert as Excel Table</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="tableName">Excel Table name</param>
    /// <returns>Inserted Worksheet</returns>
    IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName, string tableName);

    IXLCell Cell(string namedCell);

    IXLCells Cells(string namedCells);

    IXLCustomProperty CustomProperty(string name);

    /// <summary>
    /// Evaluate a formula expression.
    /// </summary>
    /// <param name="expression">Formula expression to evaluate.</param>
    /// <exception cref="MissingContextException">
    /// If the expression contains a function that requires a context (e.g. current cell or worksheet).
    /// </exception>
    XLCellValue Evaluate(string expression);

    IXLCells FindCells(Func<IXLCell, bool> predicate);

    IXLColumns FindColumns(Func<IXLColumn, bool> predicate);

    IXLRows FindRows(Func<IXLRow, bool> predicate);

#nullable enable
    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    IXLDefinedName? NamedRange(string name);

    /// <summary>
    /// Try to find a defined name. If <paramref name="name"/> specifies a sheet, try to find
    /// name in the sheet first and fall back to the workbook if not found in the sheet.
    /// <para>
    /// <example>
    /// Requested name <c>Sheet1!Name</c> will first try to find <c>Name</c> in a sheet
    /// <c>Sheet1</c> (if such sheet exists) and if not found there, tries to find <c>Name</c>
    /// in workbook.
    /// </example>
    /// </para>
    /// <para>
    /// <example>
    /// Requested name <c>Name</c> will be searched only in a workbooks <see cref="DefinedNames"/>.
    /// </example>
    /// </para>
    /// </summary>
    /// <param name="name">Name of requested name, either plain name (e.g. <c>Name</c>) or with
    /// sheet specified (e.g. <c>Sheet!Name</c>).</param>
    /// <returns>Found name or null.</returns>
    IXLDefinedName? DefinedName(string name);
#nullable disable

    IXLRange Range(string range);

    IXLRange RangeFromFullAddress(string rangeAddress, out IXLWorksheet ws);

    IXLRanges Ranges(string ranges);

    /// <summary>
    /// Force recalculation of all cell formulas.
    /// </summary>
    void RecalculateAllFormulas();

    /// <summary>
    ///   Saves the current workbook.
    /// </summary>
    void Save();

    /// <summary>
    ///   Saves the current workbook and optionally performs validation
    /// </summary>
    void Save(bool validate, bool evaluateFormulae = false);

    void Save(SaveOptions options);

    /// <summary>
    ///   Saves the current workbook to a file.
    /// </summary>
    void SaveAs(string file);

    /// <summary>
    ///   Saves the current workbook to a file and optionally validates it.
    /// </summary>
    void SaveAs(string file, bool validate, bool evaluateFormulae = false);

    void SaveAs(string file, SaveOptions options);

    /// <summary>
    ///   Saves the current workbook to a stream.
    /// </summary>
    void SaveAs(Stream stream);

    /// <summary>
    ///   Saves the current workbook to a stream and optionally validates it.
    /// </summary>
    void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false);

    void SaveAs(Stream stream, SaveOptions options);

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    IEnumerable<IXLCell> Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal, bool searchFormulae = false);

    XLWorkbook SetLockStructure(bool value);

    XLWorkbook SetLockWindows(bool value);

    XLWorkbook SetUse1904DateSystem();

    XLWorkbook SetUse1904DateSystem(bool value);

    /// <summary>
    /// Gets the Excel table of the given name
    /// </summary>
    /// <param name="tableName">Name of the table to return.</param>
    /// <param name="comparisonType">One of the enumeration values that specifies how the strings will be compared.</param>
    /// <returns>The table with given name</returns>
    /// <exception cref="ArgumentOutOfRangeException">If no tables with this name could be found in the workbook.</exception>
    IXLTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase);

    bool TryGetWorksheet(string name, out IXLWorksheet worksheet);

    IXLWorksheet Worksheet(string name);

    IXLWorksheet Worksheet(int position);
}
