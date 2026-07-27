using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.IO;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.Streaming.StreamingPackageWriter;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.Streaming;

/// <summary>
/// A forward-only writer for .xlsx workbooks that are too large to hold in memory.
/// </summary>
/// <remarks>
/// <para>
/// Rows are serialised straight into the package as they are appended, so peak memory does not
/// grow with the number of rows. In exchange the writer is append-only: rows go in ascending
/// order, one worksheet is written at a time, nothing can be read back or revised, and formulas
/// are passed through verbatim rather than evaluated. Use <see cref="XLWorkbook"/> when any of
/// that matters.
/// </para>
/// <para>
/// Two things are still proportional to the data rather than constant. Distinct <em>strings</em>
/// are held until <see cref="Finish"/> under
/// <see cref="XLStreamingStringStorage.SharedStrings"/> - switch to
/// <see cref="XLStreamingStringStorage.Inline"/> if that set is unbounded. Distinct
/// <em>styles</em> are held too, but their number is bounded by how many distinct formats the
/// caller actually uses, not by row count.
/// </para>
/// <para>
/// <see cref="Finish"/> must be called to produce a readable file: it writes the shared strings,
/// the styles, the workbook part and the package plumbing, none of which are known until the
/// last row is in. Disposing without finishing abandons the write and leaves an incomplete
/// package.
/// </para>
/// </remarks>
/// <example>
/// <code>
/// using var wb = XLStreamingWorkbook.Create(stream);
/// var sheet = wb.AddWorksheet("Data");
/// sheet.FreezeRows(1);
/// sheet.AppendRow("Name", "Qty");
/// for (var i = 0; i &lt; 1_000_000; i++)
///     sheet.AppendRow($"Item {i}", i);
/// wb.Finish();
/// </code>
/// </example>
public sealed class XLStreamingWorkbook : IDisposable
{
    private const string OfficeDocumentRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    private const string WorksheetRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    private const string StylesRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    private const string SharedStringsRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

    private const string WorkbookContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    private const string WorksheetContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    private const string StylesContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

    private const string SharedStringsContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

    private const string RelationshipsContentType =
        "application/vnd.openxmlformats-package.relationships+xml";

    private const string StylesPartName = "xl/styles.xml";
    private const string SharedStringsPartName = "xl/sharedStrings.xml";
    private const string WorkbookPartName = "xl/workbook.xml";

    private readonly StreamingPackageWriter _package;
    private readonly Stream? _ownedStream;
    private readonly List<XLStreamingWorksheet> _worksheets = [];
    private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

    private XLStreamingWorksheet? _openWorksheet;
    private bool _finished;
    private bool _disposed;

    private XLStreamingWorkbook(Stream output, Stream? ownedStream, XLStreamingOptions options)
    {
        _ownedStream = ownedStream;
        _package = new StreamingPackageWriter(output, leaveOpen: ownedStream is null, options.CompressionLevel);
        Options = options;
        SharedStrings = new StreamingSharedStringTable();
        Styles = new StreamingStyleTable();
    }

    internal XLStreamingOptions Options { get; }

    internal StreamingSharedStringTable SharedStrings { get; }

    internal StreamingStyleTable Styles { get; }

    /// <summary>
    /// Start writing a workbook to a stream. The stream only has to be writable - it is never
    /// read back or seeked - so a workbook can be written straight to a network stream. It is
    /// left open when this workbook is disposed.
    /// </summary>
    public static XLStreamingWorkbook Create(Stream output) => Create(output, null);

    /// <inheritdoc cref="Create(Stream)"/>
    public static XLStreamingWorkbook Create(Stream output, XLStreamingOptions? options)
    {
        ArgumentNullException.ThrowIfNull(output);

        return new XLStreamingWorkbook(output, null, options ?? new XLStreamingOptions());
    }

    /// <summary>
    /// Start writing a workbook to a file, creating or overwriting it.
    /// </summary>
    public static XLStreamingWorkbook Create(string path) => Create(path, null);

    /// <inheritdoc cref="Create(string)"/>
    public static XLStreamingWorkbook Create(string path, XLStreamingOptions? options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var directoryName = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directoryName))
            Directory.CreateDirectory(directoryName);

        var fileStream = File.Create(path);
        try
        {
            return new XLStreamingWorkbook(fileStream, fileStream, options ?? new XLStreamingOptions());
        }
        catch
        {
            fileStream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// A fresh, unattached style initialised to the workbook default. Configure it and pass it
    /// to a row or a cell.
    /// </summary>
    /// <remarks>
    /// The writer interns the style's value at the point of use, so one instance can be
    /// reconfigured and handed to later rows without disturbing the rows already written.
    /// </remarks>
    public IXLStyle CreateStyle() => XLStyle.CreateEmptyStyle();

    /// <summary>
    /// Add a worksheet and make it the one being written. Any worksheet still open is completed
    /// first - only one can be open at a time, because both write to the same package.
    /// </summary>
    public XLStreamingWorksheet AddWorksheet(string name)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ThrowIfFinished();

        XLHelper.ValidateSheetName(name);
        if (!_sheetNames.Add(name))
            throw new ArgumentException($"A worksheet named '{name}' has already been added.", nameof(name));

        _openWorksheet?.Complete();

        var index = _worksheets.Count + 1;
        var worksheet = new XLStreamingWorksheet(this, name, index);
        _worksheets.Add(worksheet);
        _openWorksheet = worksheet;
        return worksheet;
    }

    /// <summary>
    /// Complete the workbook: finish any open worksheet, then write the shared strings, the
    /// styles, the workbook part and the package plumbing. Repeat calls do nothing. No rows can
    /// be appended afterwards.
    /// </summary>
    public void Finish()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_finished)
            return;

        _openWorksheet?.Complete();
        _openWorksheet = null;

        // Every worksheet has to be closed before the shared strings and styles are written:
        // both are still accumulating entries while rows are being appended.
        foreach (var worksheet in _worksheets)
            worksheet.Complete();

        var writeSharedStrings = Options.StringStorage == XLStreamingStringStorage.SharedStrings &&
                                 SharedStrings.Count > 0;

        if (writeSharedStrings)
        {
            using var xml = _package.CreatePart(SharedStringsPartName);
            SharedStrings.Write(xml);
        }

        WriteStylesPart();
        WriteWorkbookPart();
        WriteWorkbookRelationships(writeSharedStrings);
        WritePackageRelationships();
        WriteContentTypes(writeSharedStrings);

        _finished = true;
    }

    internal XmlWriter CreatePart(string entryName) => _package.CreatePart(entryName);

    private void WriteStylesPart()
    {
        var stylesheet = new Stylesheet();
        WorkbookStylesPartWriter.GenerateStreamingContent(stylesheet, Styles.OrderedStyles, new SaveContext());

        using var xml = _package.CreatePart(StylesPartName);
        xml.WriteStartDocument(true);
        stylesheet.WriteTo(xml);
        xml.WriteEndDocument();
    }

    private void WriteWorkbookPart()
    {
        using var xml = _package.CreatePart(WorkbookPartName);

        xml.WriteStartDocument(true);
        xml.WriteStartElement("workbook", Main2006SsNs);
        xml.WriteAttributeString("xmlns", "r", null, RelationshipsNs);

        if (Options.Use1904DateSystem)
        {
            xml.WriteStartElement("workbookPr", Main2006SsNs);
            xml.WriteAttributeString("date1904", TrueValue);
            xml.WriteEndElement();
        }

        xml.WriteStartElement("sheets", Main2006SsNs);
        foreach (var worksheet in _worksheets)
        {
            xml.WriteStartElement("sheet", Main2006SsNs);
            xml.WriteAttributeString("name", worksheet.Name);
            xml.WriteAttribute("sheetId", (uint)worksheet.Index);
            xml.WriteAttributeString("id", RelationshipsNs, WorksheetRelationshipId(worksheet.Index));
            xml.WriteEndElement(); // sheet
        }

        xml.WriteEndElement(); // sheets
        xml.WriteEndElement(); // workbook
        xml.WriteEndDocument();
    }

    private void WriteWorkbookRelationships(bool writeSharedStrings)
    {
        using var xml = _package.CreatePart("xl/_rels/workbook.xml.rels");

        xml.WriteStartDocument(true);
        xml.WriteStartElement("Relationships", PackageRelationshipsNs);

        foreach (var worksheet in _worksheets)
        {
            WriteRelationship(xml, WorksheetRelationshipId(worksheet.Index), WorksheetRelType,
                XLStreamingWorksheet.EntryName(worksheet.Index)["xl/".Length..]);
        }

        var nextId = _worksheets.Count + 1;
        WriteRelationship(xml, $"rId{nextId}", StylesRelType, "styles.xml");

        if (writeSharedStrings)
            WriteRelationship(xml, $"rId{nextId + 1}", SharedStringsRelType, "sharedStrings.xml");

        xml.WriteEndElement(); // Relationships
        xml.WriteEndDocument();
    }

    private void WritePackageRelationships()
    {
        using var xml = _package.CreatePart("_rels/.rels");

        xml.WriteStartDocument(true);
        xml.WriteStartElement("Relationships", PackageRelationshipsNs);
        WriteRelationship(xml, "rId1", OfficeDocumentRelType, WorkbookPartName);
        xml.WriteEndElement();
        xml.WriteEndDocument();
    }

    private void WriteContentTypes(bool writeSharedStrings)
    {
        using var xml = _package.CreatePart("[Content_Types].xml");

        xml.WriteStartDocument(true);
        xml.WriteStartElement("Types", ContentTypesNs);

        xml.WriteStartElement("Default", ContentTypesNs);
        xml.WriteAttributeString("Extension", "rels");
        xml.WriteAttributeString("ContentType", RelationshipsContentType);
        xml.WriteEndElement();

        WriteContentTypeOverride(xml, "/" + WorkbookPartName, WorkbookContentType);

        foreach (var worksheet in _worksheets)
        {
            WriteContentTypeOverride(xml, "/" + XLStreamingWorksheet.EntryName(worksheet.Index),
                WorksheetContentType);
        }

        WriteContentTypeOverride(xml, "/" + StylesPartName, StylesContentType);

        if (writeSharedStrings)
            WriteContentTypeOverride(xml, "/" + SharedStringsPartName, SharedStringsContentType);

        xml.WriteEndElement(); // Types
        xml.WriteEndDocument();
    }

    private static void WriteContentTypeOverride(XmlWriter xml, string partName, string contentType)
    {
        xml.WriteStartElement("Override", ContentTypesNs);
        xml.WriteAttributeString("PartName", partName);
        xml.WriteAttributeString("ContentType", contentType);
        xml.WriteEndElement();
    }

    private static void WriteRelationship(XmlWriter xml, string id, string type, string target)
    {
        xml.WriteStartElement("Relationship", PackageRelationshipsNs);
        xml.WriteAttributeString("Id", id);
        xml.WriteAttributeString("Type", type);
        xml.WriteAttributeString("Target", target);
        xml.WriteEndElement();
    }

    private static string WorksheetRelationshipId(int index) => $"rId{index}";

    private void ThrowIfFinished()
    {
        if (_finished)
            throw new InvalidOperationException(
                $"The workbook has been finished. Create a new {nameof(XLStreamingWorkbook)} to write more.");
    }

    /// <summary>
    /// Release the package. <see cref="Finish"/> is <em>not</em> called implicitly: disposing a
    /// workbook that was never finished abandons the write, because a package without its
    /// workbook part cannot be opened.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _package.Dispose();
        _ownedStream?.Dispose();
    }
}
