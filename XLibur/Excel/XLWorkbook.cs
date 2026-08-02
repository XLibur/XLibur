using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;
using XLibur.Excel.Tables;
using XLibur.Extensions;
using XLibur.Graphics;
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
// S4136 wants every overload group adjacent. Members here are ordered by the lifecycle stage they belong to,
// which is the order a reader follows; regrouping by name would break it.
#pragma warning disable S4136
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

    internal IXLFontEngine FontEngine { get; }

    internal double DpiX { get; }

    internal double DpiY { get; }

    internal XLPivotCaches PivotCachesInternal { get; }

    internal SharedStringTable SharedStringTable { get; } = new();

    internal XLInCellImageStore InCellImages { get; } = new();

    #region Nested Type : XLLoadSource

    // ReSharper disable once InconsistentNaming
    private enum XLLoadSource
    {
        New,
        File,
        Stream
    }

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

    public IXLPersons Persons => PersonsInternal;

    internal XLPersons PersonsInternal { get; } = new();

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
            var wsName = first.UnescapeSheetName();
            var sheetlessName = split[1];
            if (TryGetWorksheet(wsName, out XLWorksheet? ws) && ws.DefinedNames.TryGetScopedValue(sheetlessName, out var sheetDefinedName))
                return sheetDefinedName;

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
    ///   Saves the current workbook back to the file or stream it was loaded from.
    /// </summary>
    /// <remarks>
    /// Preserves the encryption of the origin: a workbook opened with
    /// <see cref="LoadOptions.Password"/> is written back encrypted with that same password. Use
    /// <see cref="SaveAs(string)"/> to write a workbook without the encryption it was loaded with.
    /// </remarks>
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

        // A null password here means "leave the encryption as it was", not "write plaintext": Save
        // puts a workbook back the way it came, so one opened with a password goes back encrypted
        // under that password. Naming a password rotates it, or encrypts an origin that was plain.
        // Dropping encryption is deliberately not expressible here — that is what SaveAs is for.
        var password = string.IsNullOrEmpty(options.Password) ? _encryptionPassword : options.Password;
        if (!string.IsNullOrEmpty(password))
        {
            SaveEncrypted(password, options);
            return;
        }

        if (_loadSource == XLLoadSource.Stream)
        {
            CreatePackage(_originalStream!, false, _spreadsheetDocumentType, options);
        }
        else
            CreatePackage(_originalFile!, _spreadsheetDocumentType, options);
    }

    /// <summary>
    /// Writes the workbook back to its origin as an encrypted compound file.
    /// </summary>
    private void SaveEncrypted(string password, SaveOptions options)
    {
        // Build the whole package before touching the destination. When the origin is not yet
        // encrypted the destination is also the file or stream this package is being copied from,
        // so writing to it first would pull the ground out from under the copy.
        var package = BuildPackageInMemory(options);

        var file = _encryptedFile ?? (_loadSource == XLLoadSource.File ? _originalFile : null);
        if (file is not null)
        {
            using (var container = EncryptToBuffer(package, password))
            using (var destination = File.Create(file))
                container.WriteTo(destination);

            AdoptEncryptedOrigin(package, password, file, stream: null);
            return;
        }

        // Checked before the encryption rather than after it, so an unusable stream costs the
        // caller an exception rather than a key derivation first.
        var originStream = _encryptedStream ?? _originalStream!;
        if (!originStream.CanWrite || !originStream.CanSeek)
        {
            throw new InvalidOperationException(
                "The stream this workbook was loaded from is not writable and seekable, so Save cannot write the encrypted workbook back to it. Use one of the 'SaveAs' methods.");
        }

        using (var container = EncryptToBuffer(package, password))
        {
            originStream.Position = 0;
            container.WriteTo(originStream);
        }

        // Only now, once the new content is in: anything the old container had beyond the end of
        // the new one is what is left to drop.
        originStream.SetLength(originStream.Position);

        AdoptEncryptedOrigin(package, password, file: null, originStream);
    }

    /// <summary>
    /// Encrypts a package into a buffer, so that a destination is overwritten only once the bytes
    /// that replace it exist.
    /// </summary>
    /// <remarks>
    /// Encrypting straight into the destination would empty it first and then spend the key
    /// derivation, the encryption and the integrity hash with nothing in it, because the container
    /// is only written once all of that is done. A failure anywhere in there — and these are the
    /// steps that allocate the workbook twice over — would leave an empty file where the workbook
    /// was. This trades one more copy of the encrypted container for that not happening.
    /// </remarks>
    private MemoryStream EncryptToBuffer(MemoryStream package, string password)
    {
        var container = new MemoryStream();
        Encryptor.Encrypt(container, package.ToArray(), password);

        // No rewind: WriteTo copies the whole buffer regardless of position.
        return container;
    }

    /// <summary>
    /// Records the encrypted compound file just written as the workbook's origin, and keeps the
    /// plaintext package that produced it as the base the next save copies and patches.
    /// </summary>
    /// <remarks>
    /// The container that was written is a compound file rather than a package, so it cannot serve
    /// as that base itself. Neither can whatever the base was a moment ago, because this save may
    /// have written over it. The package built here is the one copy of the plaintext that is
    /// certainly still good.
    /// </remarks>
    private void AdoptEncryptedOrigin(MemoryStream package, string password, string? file, Stream? stream)
    {
        _loadSource = XLLoadSource.Stream;
        _originalStream = package;
        _originalFile = null;

        _encryptionPassword = password;
        _encryptedFile = file;
        _encryptedStream = stream;
    }

    /// <summary>
    /// Forgets that the workbook came from an encrypted container, after a save that wrote it out
    /// as an ordinary package.
    /// </summary>
    private void ClearEncryptedOrigin()
    {
        _encryptionPassword = null;
        _encryptedFile = null;
        _encryptedStream = null;
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

        // SaveAs states the encryption of the file it writes: a password means encrypt with it,
        // no password means write plaintext, whichever the workbook was loaded as. That asymmetry
        // with Save is the point — Save puts a file back as it was, SaveAs describes a new one.
        if (!string.IsNullOrEmpty(options.Password))
        {
            var package = BuildPackageInMemory(options);
            using (var container = EncryptToBuffer(package, options.Password))
            using (var destination = File.Create(file))
                container.WriteTo(destination);

            AdoptEncryptedOrigin(package, options.Password, file, stream: null);
            return;
        }

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
                File.Copy(_originalFile, file, true);
                File.SetAttributes(file, FileAttributes.Normal);
            }

            CreatePackage(file, GetSpreadsheetDocumentType(file), options);
        }
        else if (_loadSource == XLLoadSource.Stream)
        {
            _originalStream!.Position = 0;

            using var fileStream = File.Create(file);
            CopyStream(_originalStream, fileStream);
            CreatePackage(fileStream, false, _spreadsheetDocumentType, options);
        }

        _loadSource = XLLoadSource.File;
        _originalFile = file;
        _originalStream = null;
        ClearEncryptedOrigin();
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

        if (!string.IsNullOrEmpty(options.Password))
        {
            var package = BuildPackageInMemory(options);
            Encryptor.Encrypt(stream, package.ToArray(), options.Password);

            AdoptEncryptedOrigin(package, options.Password, file: null, stream);
            return;
        }

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
                CopyStream(_originalStream, stream);

            CreatePackage(stream, false, _spreadsheetDocumentType, options);
        }

        _loadSource = XLLoadSource.Stream;
        _originalStream = stream;
        _originalFile = null;
        ClearEncryptedOrigin();
    }

    /// <summary>
    /// Builds the complete unencrypted package in memory so it can be encrypted as a whole. Mirrors
    /// what the stream overload of <see cref="SaveAs(Stream, SaveOptions)"/> does per load source,
    /// but into a buffer, because a compound file has to be written in one go rather than patched
    /// in place. The returned stream is positioned at the start and owned by the caller, which
    /// keeps it as the workbook's new package base — see <see cref="AdoptEncryptedOrigin"/>.
    /// </summary>
    private MemoryStream BuildPackageInMemory(SaveOptions options)
    {
        var package = new MemoryStream();

        if (_loadSource == XLLoadSource.New)
        {
            CreatePackage(package, true, _spreadsheetDocumentType, options);
        }
        else if (_loadSource == XLLoadSource.File)
        {
            using (var fileStream = new FileStream(_originalFile!, FileMode.Open, FileAccess.Read))
                CopyStream(fileStream, package);

            CreatePackage(package, false, _spreadsheetDocumentType, options);
        }
        else if (_loadSource == XLLoadSource.Stream)
        {
            _originalStream!.Position = 0;
            CopyStream(_originalStream, package);
            CreatePackage(package, false, _spreadsheetDocumentType, options);
        }

        package.Position = 0;
        return package;
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

    internal static void CopyStream(Stream input, Stream output)
    {
        if (input.CanSeek)
            input.Seek(0, SeekOrigin.Begin);

        input.CopyTo(output);
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
    internal bool TryGetTable(SheetArea area, [NotNullWhen(true)] out XLTable? foundTable)
    {
        var sheet = WorksheetsInternal.FirstOrDefault<XLWorksheet>(s => XLHelper.SheetComparer.Equals(s.Name, area.Name));
        if (sheet is not null)
        {
            foreach (var table in sheet.Tables)
            {
                if (table.Area != area.Area) continue;
                foundTable = table;
                return true;
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

    /// <summary>
    /// Where the plaintext package a save copies and patches comes from. For an encrypted workbook
    /// this is always <see cref="XLLoadSource.Stream"/> over a package held in memory, because the
    /// container the workbook actually came from is a compound file rather than a package.
    /// </summary>
    private XLLoadSource _loadSource = XLLoadSource.New;

    private string? _originalFile;
    private Stream? _originalStream;

    /// <summary>
    /// The password this workbook was opened with, or last saved under. Non-null exactly when the
    /// workbook's origin is an encrypted container, which is what lets <see cref="Save()"/> put it
    /// back the way it came. Held for the lifetime of the workbook rather than for the load, so a
    /// caller who wants the password out of memory sooner should keep the workbook short-lived.
    /// </summary>
    private string? _encryptionPassword;

    /// <summary>
    /// The encrypted container <see cref="Save()"/> writes back to. Exactly one of these is set
    /// while <see cref="_encryptionPassword"/> is non-null, and neither is otherwise.
    /// </summary>
    private string? _encryptedFile;

    private Stream? _encryptedStream;

    /// <summary>
    /// The encryption a save runs the built package through before it touches the destination.
    /// </summary>
    /// <remarks>
    /// Settable so that a test can substitute an encryption that fails, which is the only way to
    /// exercise what a save leaves behind when the encryption does. Per workbook rather than static,
    /// so a substitution cannot outlive the workbook it was made for. Production code never assigns
    /// it.
    /// </remarks>
    internal IO.Encryption.IWorkbookEncryptor Encryptor { get; set; } =
        IO.Encryption.WorkbookEncryptor.Default;

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
        SharedStringTable.TrimExcess();
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
        _spreadsheetDocumentType = GetSpreadsheetDocumentType(file);

        var decrypted = OpenDecryptedIfEncrypted(file, loadOptions.Password);
        if (decrypted is not null)
        {
            // The workbook is backed by the decrypted package rather than by the file on disk. The
            // file is a compound file, not a package, so the save path could not copy and patch it.
            // It stays the origin all the same, so Save writes it back re-encrypted rather than
            // patching the copy in memory that nothing would ever read again.
            _loadSource = XLLoadSource.Stream;
            _originalStream = decrypted;
            _encryptionPassword = loadOptions.Password;
            _encryptedFile = file;
            Load(decrypted);
        }
        else
        {
            _loadSource = XLLoadSource.File;
            _originalFile = file;
            Load(file);
        }

        SharedStringTable.TrimExcess();

        if (loadOptions.RecalculateAllFormulas)
            RecalculateAllFormulas();
    }

    /// <summary>
    /// Returns the decrypted package when <paramref name="file"/> is an encrypted workbook, or
    /// <c>null</c> when it is an ordinary one and should be loaded straight from disk.
    /// </summary>
    private static MemoryStream? OpenDecryptedIfEncrypted(string file, string? password)
    {
        using var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.Read);
        return IO.Encryption.EncryptedPackageContainer.IsCompoundFile(stream)
            ? IO.Encryption.WorkbookEncryption.Decrypt(stream, password)
            : null;
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

        // A seekable stream can be sniffed for the compound file signature that marks an encrypted
        // workbook. One that isn't can only be an ordinary package, and is passed straight through
        // so that callers streaming a plain .xlsx keep working exactly as before.
        var isEncrypted = stream.CanSeek && IO.Encryption.EncryptedPackageContainer.IsCompoundFile(stream);
        if (isEncrypted)
        {
            // As with the file constructor: the decrypted package backs the workbook, while the
            // caller's stream stays the origin Save writes the re-encrypted container back to.
            _originalStream = IO.Encryption.WorkbookEncryption.Decrypt(stream, loadOptions.Password);
            _encryptionPassword = loadOptions.Password;
            _encryptedStream = stream;
        }
        else
            _originalStream = stream;

        Load(_originalStream);
        SharedStringTable.TrimExcess();

        if (loadOptions.RecalculateAllFormulas)
            RecalculateAllFormulas();
    }

    public XLWorkbook(LoadOptions loadOptions)
    {
        ArgumentNullException.ThrowIfNull(loadOptions);

        DpiX = loadOptions.Dpi.X;
        DpiY = loadOptions.Dpi.Y;
        var explicitGraphic = loadOptions.GraphicEngine;
        var fontEngine = loadOptions.FontEngine
                         ?? (explicitGraphic as IXLFontEngine)
                         ?? LoadOptions.DefaultFontEngine
                         ?? DefaultFontEngineProbe.TryResolveDefault();
        GraphicEngine = explicitGraphic
                        ?? LoadOptions.DefaultGraphicEngine
                        ?? (fontEngine is not null
                            ? new DefaultGraphicEngine(fontEngine)
                            : throw new InvalidOperationException(
                                "No font engine is available. Install the XLibur.Fonts.SkiaSharp package (or XLibur.Bundle) " +
                                "for the default engine, or register one explicitly by setting LoadOptions.FontEngine / " +
                                "LoadOptions.DefaultFontEngine (e.g. SkiaSharpFontBootstrap.Register() or SixLaborsV1FontBootstrap.Register())."));
        FontEngine = fontEngine
                     ?? (GraphicEngine as IXLFontEngine ?? new GraphicEngineFontAdapter(GraphicEngine));
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
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposing)
            return;

        Worksheets.ForEach(w => ((XLWorksheet)w).Cleanup());

        // Release calc engine and its heavy structures (DependencyTree,
        // CalculationChain, ExpressionCache, ArrayPool buffers).
        _calcEngine = null;

        // Release shared string table entries and reverse dictionary.
        SharedStringTable.Clear();

        // Dispose in-cell image MemoryStreams and release collections.
        InCellImages.Dispose();
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

    /// <summary>
    /// Monotonic counter incremented on every workbook edit (cell value change, formula
    /// change, structural change). Formulas record the epoch at which they were last
    /// evaluated; a formula is dirty when its recorded epoch differs from this one.
    /// </summary>
    /// <remarks>Starts at 1 so that the default <c>0</c> on <see cref="XLCellFormula"/>
    /// always reads as "never evaluated".</remarks>
    internal long EditEpoch { get; private set; } = 1;

    internal void BumpEditEpoch() => EditEpoch++;

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
            sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(Area.Full);

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
        // An encrypted workbook is backed by a package in memory, which says nothing useful about
        // where it came from. Name the container instead.
        if (_encryptedFile is not null)
            return $"XLWorkbook({_encryptedFile}, encrypted)";

        if (_encryptedStream is not null)
            return $"XLWorkbook({_encryptedStream}, encrypted)";

        return _loadSource switch
        {
            XLLoadSource.New => "XLWorkbook(new)",
            XLLoadSource.File => $"XLWorkbook({_originalFile})",
            XLLoadSource.Stream => $"XLWorkbook({_originalStream})",
            _ => throw new NotImplementedException()
        };
    }
}
