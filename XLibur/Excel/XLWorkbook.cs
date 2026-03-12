using ClosedXML.Excel.CalcEngine;
using ClosedXML.Graphics;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using static XLibur.Excel.XLProtectionAlgorithm;

namespace XLibur.Excel;

// ReSharper disable once InconsistentNaming
public enum XLCalculateMode
{
    Auto,
    AutoNoTable,
    Manual,
    Default
}

// ReSharper disable once InconsistentNaming
public enum XLReferenceStyle
{
    R1C1,
    A1,
    Default
}

// ReSharper disable once InconsistentNaming
public enum XLCellSetValueBehavior
{
    /// <summary>
    /// Analyze input string and convert value. For avoid analyzing use escape symbol '
    /// </summary>
    Smart = 0,

    /// <summary>
    /// Direct set value. If a value has an unsupported type-value will be stored as string returned by <see cref = "object.ToString()" />
    /// </summary>
    Simple = 1,
}

// ReSharper disable once InconsistentNaming
public partial class XLWorkbook : IXLWorkbook
{
    #region Static

    public static IXLStyle DefaultStyle => XLStyle.Default;

    internal static XLStyleValue DefaultStyleValue => XLStyleValue.Default;

    public static double DefaultRowHeight { get; private set; }

    public static double DefaultColumnWidth { get; private set; }

    public static IXLPageSetup DefaultPageOptions
    {
        get
        {
            var defaultPageOptions = new XLPageSetup(null!, null!)
            {
                PageOrientation = XLPageOrientation.Default,
                Scale = 100,
                PaperSize = XLPaperSize.LetterPaper,
                Margins = new XLMargins
                {
                    Top = 0.75,
                    Bottom = 0.5,
                    Left = 0.75,
                    Right = 0.75,
                    Header = 0.5,
                    Footer = 0.75
                },
                ScaleHFWithDocument = true,
                AlignHFWithMargins = true,
                PrintErrorValue = XLPrintErrorValues.Displayed,
                ShowComments = XLShowCommentsValues.None
            };
            return defaultPageOptions;
        }
    }

    public static IXLOutline DefaultOutline => new XLOutline
    {
        SummaryHLocation = XLOutlineSummaryHLocation.Right,
        SummaryVLocation = XLOutlineSummaryVLocation.Bottom
    };

    /// <summary>
    ///   Behavior for <see cref = "IXLCell.set_Value" />
    /// </summary>
    public static XLCellSetValueBehavior CellSetValueBehavior { get; set; }

    public static XLWorkbook OpenFromTemplate(string path)
    {
        return new XLWorkbook(path, asTemplate: true);
    }

    #endregion Static

    internal readonly List<UnsupportedSheet> UnsupportedSheets = [];

    internal IXLGraphicEngine GraphicEngine { get; }

    internal double DpiX { get; }

    internal double DpiY { get; }

    internal XLPivotCaches PivotCachesInternal { get; }

    internal SharedStringTable SharedStringTable { get; } = new();

    #region Nested Type : XLLoadSource

    // ReSharper disable once InconsistentNaming
    private enum XLLoadSource
    {
        New,
        File,
        Stream
    };

    #endregion Nested Type : XLLoadSource

    internal XLWorksheets WorksheetsInternal { get; private set; }

    /// <summary>
    ///   Gets an object to manipulate the worksheets.
    /// </summary>
    public IXLWorksheets Worksheets => WorksheetsInternal;

    internal XLDefinedNames DefinedNamesInternal { get; }

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    public IXLDefinedNames NamedRanges => DefinedNamesInternal;

    /// <summary>
    ///   Gets an object to manipulate this workbook's named ranges.
    /// </summary>
    public IXLDefinedNames DefinedNames => DefinedNamesInternal;

    /// <summary>
    ///   Gets an object to manipulate this workbook's theme.
    /// </summary>
    public IXLTheme Theme { get; private set; } = null!;

    /// <summary>
    /// All pivot caches in the workbook, whether they have a pivot table or not.
    /// </summary>
    public IXLPivotCaches PivotCaches => PivotCachesInternal;

    /// <summary>
    ///   Gets or sets the default style for the workbook.
    ///   <para>All new worksheets will use this style.</para>
    /// </summary>
    public IXLStyle Style { get; set; }

    /// <summary>
    ///   Gets or sets the default row height for the workbook.
    ///   <para>All new worksheets will use this row height.</para>
    /// </summary>
    public double RowHeight { get; set; }

    /// <summary>
    ///   Gets or sets the default column width for the workbook.
    ///   <para>All new worksheets will use this column width.</para>
    /// </summary>
    public double ColumnWidth { get; set; }

    /// <summary>
    ///   Gets or sets the default page options for the workbook.
    ///   <para>All new worksheets will use these page options.</para>
    /// </summary>
    public IXLPageSetup PageOptions { get; set; }

    /// <summary>
    ///   Gets or sets the default outline options for the workbook.
    ///   <para>All new worksheets will use these outline options.</para>
    /// </summary>
    public IXLOutline Outline { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's properties.
    /// </summary>
    public XLWorkbookProperties Properties { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's calculation mode.
    /// </summary>
    public XLCalculateMode CalculateMode { get; set; }

    public bool CalculationOnSave { get; set; }

    public bool ForceFullCalculation { get; set; }

    public bool FullCalculationOnLoad { get; set; }

    public bool FullPrecision { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's reference style.
    /// </summary>
    public XLReferenceStyle ReferenceStyle { get; set; }

    public IXLCustomProperties CustomProperties { get; private set; }

    public bool ShowFormulas { get; set; }

    public bool ShowGridLines { get; set; }

    public bool ShowOutlineSymbols { get; set; }

    public bool ShowRowColHeaders { get; set; }

    public bool ShowRuler { get; set; }

    public bool ShowWhiteSpace { get; set; }

    public bool ShowZeros { get; set; }

    public bool RightToLeft { get; set; }

    public bool DefaultShowFormulas => false;

    public bool DefaultShowGridLines => true;

    public bool DefaultShowOutlineSymbols => true;

    public bool DefaultShowRowColHeaders => true;

    public bool DefaultShowRuler => true;

    public bool DefaultShowWhiteSpace => true;

    public bool DefaultShowZeros => true;

    public IXLFileSharing FileSharing { get; } = new XLFileSharing();

    public bool DefaultRightToLeft => false;

    private void InitializeTheme()
    {
        Theme = new XLTheme
        {
            Text1 = XLColor.FromHtml("#FF000000"),
            Background1 = XLColor.FromHtml("#FFFFFFFF"),
            Text2 = XLColor.FromHtml("#FF1F497D"),
            Background2 = XLColor.FromHtml("#FFEEECE1"),
            Accent1 = XLColor.FromHtml("#FF4F81BD"),
            Accent2 = XLColor.FromHtml("#FFC0504D"),
            Accent3 = XLColor.FromHtml("#FF9BBB59"),
            Accent4 = XLColor.FromHtml("#FF8064A2"),
            Accent5 = XLColor.FromHtml("#FF4BACC6"),
            Accent6 = XLColor.FromHtml("#FFF79646"),
            Hyperlink = XLColor.FromHtml("#FF0000FF"),
            FollowedHyperlink = XLColor.FromHtml("#FF800080")
        };
    }

    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    public IXLDefinedName? NamedRange(string name) => DefinedName(name);

    /// <inheritdoc/>
    public IXLDefinedName? DefinedName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (name.Contains('!'))
        {
            var split = name.Split('!');
            var first = split[0];
            var wsName = first.StartsWith("'") ? first.Substring(1, first.Length - 2) : first;
            var sheetlessName = split[1];
            if (TryGetWorksheet(wsName, out XLWorksheet? ws))
            {
                if (ws.DefinedNames.TryGetScopedValue(sheetlessName, out var sheetDefinedName))
                    return sheetDefinedName;
            }

            name = sheetlessName;
        }

        return DefinedNamesInternal.TryGetScopedValue(name, out var definedName) ? definedName : null;
    }

    public bool TryGetWorksheet(string name, [NotNullWhen(true)] out IXLWorksheet? worksheet)
    {
        if (TryGetWorksheet(name, out XLWorksheet? foundSheet))
        {
            worksheet = foundSheet;
            return true;
        }

        worksheet = null;
        return false;
    }

    internal bool TryGetWorksheet(string name, [NotNullWhen(true)] out XLWorksheet? worksheet)
    {
        return WorksheetsInternal.TryGetWorksheet(name, out worksheet);
    }

    public IXLRange? RangeFromFullAddress(string rangeAddress, out IXLWorksheet? ws)
    {
        ArgumentNullException.ThrowIfNull(rangeAddress);
        if (!rangeAddress.Contains('!'))
        {
            ws = null;
            return null;
        }

        var split = rangeAddress.Split('!');
        var wsName = split[0].UnescapeSheetName();
        if (TryGetWorksheet(wsName, out XLWorksheet? sheet))
        {
            ws = sheet;
            return sheet.Range(split[1]);
        }

        ws = null;
        return null;
    }

    public IXLCell? CellFromFullAddress(string cellAddress, out IXLWorksheet? ws)
    {
        ArgumentNullException.ThrowIfNull(cellAddress);
        if (!cellAddress.Contains('!'))
        {
            ws = null;
            return null;
        }

        var split = cellAddress.Split('!');
        var wsName = split[0].UnescapeSheetName();
        if (TryGetWorksheet(wsName, out XLWorksheet? sheet))
        {
            ws = sheet;
            return sheet.Cell(split[1]);
        }

        ws = null;
        return null;
    }

    /// <summary>
    ///   Saves the current workbook.
    /// </summary>
    public void Save()
    {
        Save(false);
    }

    /// <summary>
    ///   Saves the current workbook and optionally performs validation
    /// </summary>
    public void Save(bool validate, bool evaluateFormulae = false)
    {
        Save(new SaveOptions
        {
            ValidatePackage = validate,
            EvaluateFormulasBeforeSaving = evaluateFormulae,
            GenerateCalculationChain = true
        });
    }

    public void Save(SaveOptions options)
    {
        CheckForWorksheetsPresent();
        if (_loadSource == XLLoadSource.New)
            throw new InvalidOperationException("This is a new file. Please use one of the 'SaveAs' methods.");

        if (_loadSource == XLLoadSource.Stream)
        {
            CreatePackage(_originalStream!, false, _spreadsheetDocumentType, options);
        }
        else
            CreatePackage(_originalFile!, _spreadsheetDocumentType, options);
    }

    /// <summary>
    ///   Saves the current workbook to a file.
    /// </summary>
    public void SaveAs(string file)
    {
        SaveAs(file, false);
    }

    /// <summary>
    ///   Saves the current workbook to a file and optionally validates it.
    /// </summary>
    public void SaveAs(string file, bool validate, bool evaluateFormulae = false)
    {
        SaveAs(file, new SaveOptions
        {
            ValidatePackage = validate,
            EvaluateFormulasBeforeSaving = evaluateFormulae,
            GenerateCalculationChain = true
        });
    }

    public void SaveAs(string file, SaveOptions options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(file);
        CheckForWorksheetsPresent();

        var directoryName = Path.GetDirectoryName(file);
        if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

        if (_loadSource == XLLoadSource.New)
        {
            if (File.Exists(file))
                File.Delete(file);

            CreatePackage(file, GetSpreadsheetDocumentType(file), options);
        }
        else if (_loadSource == XLLoadSource.File)
        {
            if (string.Compare(_originalFile!.Trim(), file.Trim(), StringComparison.OrdinalIgnoreCase) != 0)
            {
                File.Copy(_originalFile!, file, true);
                File.SetAttributes(file, FileAttributes.Normal);
            }

            CreatePackage(file, GetSpreadsheetDocumentType(file), options);
        }
        else if (_loadSource == XLLoadSource.Stream)
        {
            _originalStream!.Position = 0;

            using var fileStream = File.Create(file);
            CopyStream(_originalStream!, fileStream);
            CreatePackage(fileStream, false, _spreadsheetDocumentType, options);
        }

        _loadSource = XLLoadSource.File;
        _originalFile = file;
        _originalStream = null;
    }

    private static SpreadsheetDocumentType GetSpreadsheetDocumentType(string filePath)
    {
        var extension = Path.GetExtension(filePath);

        if (string.IsNullOrEmpty(extension)) throw new ArgumentException("Empty extension is not supported.");
        extension = extension[1..].ToLowerInvariant();

        return extension switch
        {
            "xlsm" => SpreadsheetDocumentType.MacroEnabledWorkbook,
            "xltm" => SpreadsheetDocumentType.MacroEnabledTemplate,
            "xlsx" => SpreadsheetDocumentType.Workbook,
            "xltx" => SpreadsheetDocumentType.Template,
            _ => throw new ArgumentException(
                $"Extension '{extension}' is not supported. Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.")
        };
    }

    private void CheckForWorksheetsPresent()
    {
        if (Worksheets.Count == 0)
            throw new InvalidOperationException("Workbooks need at least one worksheet.");
    }

    /// <summary>
    ///   Saves the current workbook to a stream.
    /// </summary>
    public void SaveAs(Stream stream)
    {
        SaveAs(stream, false);
    }

    /// <summary>
    ///   Saves the current workbook to a stream and optionally validates it.
    /// </summary>
    public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false)
    {
        SaveAs(stream, new SaveOptions
        {
            ValidatePackage = validate,
            EvaluateFormulasBeforeSaving = evaluateFormulae,
            GenerateCalculationChain = true
        });
    }

    public void SaveAs(Stream stream, SaveOptions options)
    {
        CheckForWorksheetsPresent();
        if (_loadSource == XLLoadSource.New)
        {
            // This method or better the method SpreadsheetDocument.Create which is called
            // inside 'CreatePackage' need a stream which CanSeek & CanRead
            // and an ordinary Response stream of a webserver can't do this,
            // so we have to ask and provide a way around this
            if (stream is { CanRead: true, CanSeek: true, CanWrite: true })
            {
                // all is fine the package can be created directly
                CreatePackage(stream, true, _spreadsheetDocumentType, options);
            }
            else
            {
                // the harder way
                using var ms = new MemoryStream();
                CreatePackage(ms, true, _spreadsheetDocumentType, options);
                // not really necessary, because I changed CopyStream too.
                // For better understanding and if somebody in the future provides a changed version of CopyStream
                ms.Position = 0;
                CopyStream(ms, stream);
            }
        }
        else if (_loadSource == XLLoadSource.File)
        {
            using (var fileStream = new FileStream(_originalFile!, FileMode.Open, FileAccess.Read))
            {
                CopyStream(fileStream, stream);
            }

            CreatePackage(stream, false, _spreadsheetDocumentType, options);
        }
        else if (_loadSource == XLLoadSource.Stream)
        {
            _originalStream!.Position = 0;
            if (_originalStream != stream)
                CopyStream(_originalStream!, stream);

            CreatePackage(stream, false, _spreadsheetDocumentType, options);
        }

        _loadSource = XLLoadSource.Stream;
        _originalStream = stream;
        _originalFile = null;
    }

    internal static void CopyStream(Stream input, Stream output)
    {
        var buffer = new byte[8 * 1024];
        int len;
        // It is always a good idea to rewind the input stream, or not?
        if (input.CanSeek)
            input.Seek(0, SeekOrigin.Begin);
        while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            output.Write(buffer, 0, len);
        // And flushing the output after writing
        output.Flush();
    }

    public IXLTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
    {
        return !TryGetTable(tableName, out var table, comparisonType)
            ? throw new ArgumentOutOfRangeException($"Table {tableName} was not found.")
            : table;
    }

    /// <summary>
    /// Try to find a table with <paramref name="tableName"/> in a workbook.
    /// </summary>
    internal bool TryGetTable(string tableName, [NotNullWhen(true)] out XLTable? table,
        StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
    {
        table = WorksheetsInternal
            .SelectMany<XLWorksheet, XLTable>(ws => ws.Tables)
            .FirstOrDefault(t => t.Name.Equals(tableName, comparisonType));

        return table is not null;
    }

    /// <summary>
    /// Try to find a table that covers same area as the <paramref name="area"/> in a workbook.
    /// </summary>
    internal bool TryGetTable(XLBookArea area, [NotNullWhen(true)] out XLTable? foundTable)
    {
        foreach (var sheet in WorksheetsInternal)
        {
            if (XLHelper.SheetComparer.Equals(sheet.Name, area.Name))
            {
                foreach (var table in sheet.Tables)
                {
                    if (table.Area != area.Area)
                        continue;

                    foundTable = table;
                    return true;
                }

                // No other sheet has correct name.
                break;
            }
        }

        foundTable = null;
        return false;
    }

    public IXLWorksheet Worksheet(string name)
    {
        return WorksheetsInternal.Worksheet(name);
    }

    public IXLWorksheet Worksheet(int position)
    {
        return WorksheetsInternal.Worksheet(position);
    }

    public IXLCustomProperty CustomProperty(string name)
    {
        return CustomProperties.CustomProperty(name);
    }

    public IXLCells FindCells(Func<IXLCell, bool> predicate)
    {
        var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
        foreach (var ws in WorksheetsInternal)
        {
            foreach (var xlCell in ws.CellsUsed(XLCellsUsedOptions.All))
            {
                var cell = (XLCell)xlCell;
                if (predicate(cell))
                    cells.Add(cell);
            }
        }

        return cells;
    }

    public IXLRows FindRows(Func<IXLRow, bool> predicate)
    {
        var rows = new XLRows(worksheet: null);
        foreach (var ws in WorksheetsInternal)
        {
            foreach (var row in ws.Rows().Where(predicate))
                rows.Add((XLRow)row);
        }

        return rows;
    }

    public IXLColumns FindColumns(Func<IXLColumn, bool> predicate)
    {
        var columns = new XLColumns(worksheet: null);
        foreach (var ws in WorksheetsInternal)
        {
            foreach (var column in ws.Columns().Where(predicate))
                columns.Add((XLColumn)column);
        }

        return columns;
    }

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    public IEnumerable<IXLCell> Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal,
        bool searchFormulae = false)
    {
        foreach (var ws in WorksheetsInternal)
        {
            foreach (var cell in ws.Search(searchText, compareOptions, searchFormulae))
                yield return cell;
        }
    }

    #region Fields

    private XLLoadSource _loadSource = XLLoadSource.New;
    private string? _originalFile;
    private Stream? _originalStream;

    #endregion Fields

    #region Constructor

    /// <summary>
    ///   Creates a new Excel workbook.
    /// </summary>
    public XLWorkbook()
        : this(new LoadOptions())
    {
    }

    internal XLWorkbook(string file, bool asTemplate)
        : this(new LoadOptions())
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(file);
        LoadSheetsFromTemplate(file);
    }

    /// <summary>
    ///   Opens an existing workbook from a file.
    /// </summary>
    /// <param name = "file">The file to open.</param>
    public XLWorkbook(string file)
        : this(file, new LoadOptions())
    {
    }

    public XLWorkbook(string file, LoadOptions loadOptions)
        : this(loadOptions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(file);
        _loadSource = XLLoadSource.File;
        _originalFile = file;
        _spreadsheetDocumentType = GetSpreadsheetDocumentType(_originalFile);
        Load(file);

        if (loadOptions.RecalculateAllFormulas)
            RecalculateAllFormulas();
    }

    /// <summary>
    ///   Opens an existing workbook from a stream.
    /// </summary>
    /// <param name = "stream">The stream to open.</param>
    public XLWorkbook(Stream stream)
        : this(stream, new LoadOptions())
    {
    }

    public XLWorkbook(Stream stream, LoadOptions loadOptions)
        : this(loadOptions)
    {
        ArgumentNullException.ThrowIfNull(stream);
        _loadSource = XLLoadSource.Stream;
        _originalStream = stream;
        Load(stream);

        if (loadOptions.RecalculateAllFormulas)
            RecalculateAllFormulas();
    }

    public XLWorkbook(LoadOptions loadOptions)
    {
        if (loadOptions is null)
            throw new ArgumentNullException(nameof(loadOptions));

        DpiX = loadOptions.Dpi.X;
        DpiY = loadOptions.Dpi.Y;
        GraphicEngine = loadOptions.GraphicEngine ??
                        LoadOptions.DefaultGraphicEngine ?? DefaultGraphicEngine.Instance.Value;
        Protection = new XLWorkbookProtection(DefaultProtectionAlgorithm);
        DefaultRowHeight = 15;
        DefaultColumnWidth = 8.43;
        Style = new XLStyle(null!, DefaultStyle);
        RowHeight = DefaultRowHeight;
        ColumnWidth = DefaultColumnWidth;
        PageOptions = DefaultPageOptions;
        Outline = DefaultOutline;
        Properties = new XLWorkbookProperties();
        CalculateMode = XLCalculateMode.Default;
        ReferenceStyle = XLReferenceStyle.Default;
        InitializeTheme();
        ShowFormulas = DefaultShowFormulas;
        ShowGridLines = DefaultShowGridLines;
        ShowOutlineSymbols = DefaultShowOutlineSymbols;
        ShowRowColHeaders = DefaultShowRowColHeaders;
        ShowRuler = DefaultShowRuler;
        ShowWhiteSpace = DefaultShowWhiteSpace;
        ShowZeros = DefaultShowZeros;
        RightToLeft = DefaultRightToLeft;
        WorksheetsInternal = new XLWorksheets(this);
        DefinedNamesInternal = new XLDefinedNames(this);
        PivotCachesInternal = new XLPivotCaches(this);
        CustomProperties = new XLCustomProperties(this);
        ShapeIdManager = new XLIdManager();
        Author = Environment.UserName;
    }

    #endregion Constructor

    #region Nested type: UnsupportedSheet

    internal sealed class UnsupportedSheet
    {
        public bool IsActive;
        public uint SheetId;
        public int Position;
    }

    #endregion Nested type: UnsupportedSheet

    public IXLCell? Cell(string namedCell)
    {
        var namedRange = DefinedName(namedCell);
        return namedRange != null
            ? namedRange.Ranges.FirstOrDefault()?.FirstCell()
            : CellFromFullAddress(namedCell, out _);
    }

    public IXLCells Cells(string namedCells)
    {
        return Ranges(namedCells).Cells();
    }

    public IXLRange? Range(string range)
    {
        var namedRange = DefinedName(range);
        return namedRange != null ? namedRange.Ranges.FirstOrDefault() : RangeFromFullAddress(range, out _);
    }

    public IXLRanges Ranges(string ranges)
    {
        var retVal = new XLRanges();
        var rangePairs = ranges.Split(',');
        foreach (var range in rangePairs.Select(r => Range(r.Trim())).Where(range => range != null))
        {
            retVal.Add(range!);
        }

        return retVal;
    }

    internal XLIdManager ShapeIdManager { get; private set; }

    public void Dispose()
    {
        Worksheets.ForEach(w => ((XLWorksheet)w).Cleanup());
    }


    public bool Use1904DateSystem { get; set; }

    public XLWorkbook SetUse1904DateSystem()
    {
        return SetUse1904DateSystem(true);
    }

    public XLWorkbook SetUse1904DateSystem(bool value)
    {
        Use1904DateSystem = value;
        return this;
    }

    public IXLWorksheet AddWorksheet()
    {
        return Worksheets.Add();
    }

    public IXLWorksheet AddWorksheet(int position)
    {
        return Worksheets.Add(position);
    }

    public IXLWorksheet AddWorksheet(string sheetName)
    {
        return Worksheets.Add(sheetName);
    }

    public IXLWorksheet AddWorksheet(string sheetName, int position)
    {
        return Worksheets.Add(sheetName, position);
    }

    public void AddWorksheet(DataSet dataSet)
    {
        Worksheets.Add(dataSet);
    }

    public void AddWorksheet(IXLWorksheet worksheet)
    {
        worksheet.CopyTo(this, worksheet.Name);
    }

    public IXLWorksheet AddWorksheet(DataTable dataTable)
    {
        return Worksheets.Add(dataTable);
    }

    public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName)
    {
        return Worksheets.Add(dataTable, sheetName);
    }

    public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName, string tableName)
    {
        return Worksheets.Add(dataTable, sheetName, tableName);
    }

    private XLCalcEngine? _calcEngine;

    internal XLCalcEngine CalcEngine
    {
        get { return _calcEngine ??= new XLCalcEngine(CultureInfo.CurrentCulture); }
    }

    public XLCellValue Evaluate(string expression)
    {
        return CalcEngine.EvaluateFormula(expression, this).ToCellValue();
    }

    /// <summary>
    /// Force recalculation of all cell formulas.
    /// </summary>
    public void RecalculateAllFormulas()
    {
        foreach (var sheet in WorksheetsInternal)
            sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(XLSheetRange.Full);

        CalcEngine.Recalculate(this, null);
    }

    private static XLCalcEngine? _calcEngineExpr;
    private readonly SpreadsheetDocumentType _spreadsheetDocumentType;

    private static XLCalcEngine CalcEngineExpr
    {
        get { return _calcEngineExpr ??= new XLCalcEngine(CultureInfo.InvariantCulture); }
    }

    /// <summary>
    /// Evaluate a formula and return a value. Formulas with References don't work,
    /// and culture used for conversion is invariant.
    /// </summary>
    public static XLCellValue EvaluateExpr(string expression)
    {
        return CalcEngineExpr.EvaluateFormula(expression).ToCellValue();
    }

    /// <summary>
    /// Evaluate a formula and return a value. Use current culture.
    /// </summary>
    internal static XLCellValue EvaluateExprCurrent(string expression)
    {
        return new XLCalcEngine(CultureInfo.CurrentCulture).EvaluateFormula(expression).ToCellValue();
    }

    public string Author { get; set; }

    public bool LockStructure
    {
        get => Protection.IsProtected && !Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure);
        set
        {
            if (!Protection.IsProtected)
                throw new InvalidOperationException(
                    $"Enable workbook protection before setting the {nameof(LockStructure)} property");

            Protection.AllowElement(XLWorkbookProtectionElements.Structure, value);
        }
    }

    public XLWorkbook SetLockStructure(bool value)
    {
        LockStructure = value;
        return this;
    }

    public bool LockWindows
    {
        get => Protection.IsProtected && !Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows);
        set
        {
            if (!Protection.IsProtected)
                throw new InvalidOperationException(
                    $"Enable workbook protection before setting the {nameof(LockWindows)} property");

            Protection.AllowElement(XLWorkbookProtectionElements.Windows, value);
        }
    }

    public XLWorkbook SetLockWindows(bool value)
    {
        LockWindows = value;
        return this;
    }

    public bool IsPasswordProtected => Protection.IsPasswordProtected;

    public bool IsProtected => Protection.IsProtected;

    IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protection
    {
        get => Protection;
        set => Protection = (XLWorkbookProtection)value;
    }

    internal XLWorkbookProtection Protection
    {
        get;
        set => field = value.Clone().CastTo<XLWorkbookProtection>();
    }

    public IXLWorkbookProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protection.Protect(algorithm);
    }

    public IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements)
        => Protection.Protect(allowedElements);

    public IXLWorkbookProtection Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
        => Protection.Protect(algorithm, allowedElements);

    public IXLWorkbookProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm)

    {
        return Protect(password, algorithm, XLWorkbookProtectionElements.Windows);
    }

    public IXLWorkbookProtection Protect(string password, Algorithm algorithm,
        XLWorkbookProtectionElements allowedElements)
    {
        return Protection.Protect(password, algorithm, allowedElements);
    }

    IXLElementProtection IXLProtectable.Protect(Algorithm algorithm)
    {
        return Protect(algorithm);
    }

    IXLElementProtection IXLProtectable.Protect(string password, Algorithm algorithm)
    {
        return Protect(password, algorithm);
    }

    IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protect(
        XLWorkbookProtectionElements allowedElements)
        => Protect(allowedElements);

    IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protect(
        Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
        => Protect(algorithm, allowedElements);

    IXLWorkbookProtection IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>.Protect(string password,
        Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
        => Protect(password, algorithm, allowedElements);

    public IXLWorkbookProtection Unprotect()
    {
        return Protection.Unprotect();
    }

    public IXLWorkbookProtection Unprotect(string password)
    {
        return Protection.Unprotect(password);
    }

    IXLElementProtection IXLProtectable.Unprotect()
    {
        return Unprotect();
    }

    IXLElementProtection IXLProtectable.Unprotect(string password)
    {
        return Unprotect(password);
    }

    /// <summary>
    /// Notify various components of a workbook that a sheet has been added.
    /// </summary>
    internal void NotifyWorksheetAdded(XLWorksheet newSheet)
    {
        _calcEngine?.OnAddedSheet(newSheet);
    }

    /// <summary>
    /// Notify various components of a workbook that the sheet is about to be removed.
    /// </summary>
    internal void NotifyWorksheetDeleting(XLWorksheet sheet)
    {
        _calcEngine?.OnDeletingSheet(sheet);
    }

    public override string ToString()
    {
        return _loadSource switch
        {
            XLLoadSource.New => "XLWorkbook(new)",
            XLLoadSource.File => $"XLWorkbook({_originalFile})",
            XLLoadSource.Stream => $"XLWorkbook({_originalStream})",
            _ => throw new NotImplementedException()
        };
    }
}
