using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Coordinates;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

/// <summary>
/// The <c>&lt;sheetData&gt;</c> of a worksheet part as the file had it, which a save writes in place
/// of the cells streamed from the model when they have not changed since the load (#702, tier 2).
/// The rest of the part is still written from the model.
/// </summary>
/// <remarks>
/// The bytes are read from the part in the package being saved, before the save replaces it. A save
/// of a loaded workbook starts from a copy of the loaded package, so the part is the one the load
/// read; its markup around the cells is compared with the copy the load kept to make sure. Reading
/// the part again costs an inflate. Keeping the bytes from the load would save that, but would hold
/// the cells of every sheet in memory until a save that may never come.
/// <para>
/// A kept cell names a shared string and a cell format by their index in the file. Both still name
/// the same thing: the save writes every loaded text at its file index, or keeps no cells at all
/// (<see cref="SharedStringTable.WritesTextsAtFileIndices"/>), and the styles part is loaded and
/// only appended to.
/// </para>
/// </remarks>
internal sealed class LoadedSheetData : IDisposable
{
    private readonly WorksheetPartBuffer _part;

    /// <summary>
    /// The namespace declarations to add to the start tag, as UTF-8, each with a leading space.
    /// </summary>
    private readonly byte[] _declarations;

    private LoadedSheetData(WorksheetPartBuffer part, byte[] declarations)
    {
        _part = part;
        _declarations = declarations;
    }

    /// <summary>
    /// Can the save keep the cells of <paramref name="xlWorksheet"/> as the file had them? Both
    /// tiers of #702 ask this: keeping the whole part, and keeping only its cells.
    /// </summary>
    /// <param name="xlWorksheet">The sheet being saved.</param>
    /// <param name="options">The options of the save.</param>
    /// <param name="context">The save.</param>
    internal static bool CanKeepCells(XLWorksheet xlWorksheet, SaveOptions options, SaveContext context) =>
        !options.RewriteUnchangedSheets
        && context.StylesheetWasLoaded
        && CellsCanBeKept(xlWorksheet)
        && context.SharedStringsAtFileIndices(xlWorksheet.Workbook.SharedStringTable);

    /// <summary>
    /// Reads the cells of <paramref name="worksheetPart"/> as the file had them, or returns null when
    /// the save has to write them from the model. Call it only when <see cref="CanKeepCells"/> says
    /// the cells can be kept.
    /// </summary>
    /// <param name="worksheetPart">The part in the package being saved, not yet written.</param>
    /// <param name="loadedMarkup">
    /// The part without its cells, as the load kept it (<see cref="XLWorksheet.TakePartWithoutSheetData"/>).
    /// </param>
    /// <param name="worksheet">The root the save writes, which the kept cells have to fit under.</param>
    internal static LoadedSheetData? TryRead(WorksheetPart worksheetPart, byte[] loadedMarkup, Worksheet worksheet)
    {
        var part = WorksheetPartBuffer.TryRead(worksheetPart);
        if (part is null)
            return null;

        // An empty <sheetData/> is as quick to write as to copy.
        if (part.SheetData.IsEmptyElement
            || !part.EqualsWithoutSheetData(loadedMarkup)
            || !TryGetDeclarationsToAdd(part, worksheet, out var declarations))
        {
            part.Dispose();
            return null;
        }

        return new LoadedSheetData(part, declarations);
    }

    /// <summary>
    /// Writes the kept <c>&lt;sheetData&gt;</c> at the position of <paramref name="xml"/>.
    /// </summary>
    /// <param name="xml">The writer of the part, positioned where <c>&lt;sheetData&gt;</c> goes.</param>
    /// <param name="output">The stream <paramref name="xml"/> writes to.</param>
    /// <remarks>
    /// The bytes go to the stream past the writer, so they are written exactly as the file had
    /// them; <see cref="XmlWriter.WriteRaw(string)"/> would respell line breaks. The empty raw write
    /// first closes the root's start tag, which the writer holds open while attributes could still
    /// follow, and the flush then hands on everything the writer buffered.
    /// </remarks>
    internal void WriteTo(XmlWriter xml, Stream output)
    {
        xml.WriteRaw(string.Empty);
        xml.Flush();

        var part = _part.Xml;
        var sheetData = _part.SheetData;
        var startTagClose = sheetData.StartTagEnd - 1;

        output.Write(part[sheetData.Start..startTagClose]);
        output.Write(_declarations);
        output.Write(part[startTagClose..sheetData.End]);
    }

    public void Dispose() => _part.Dispose();

    /// <summary>
    /// Would the save write the sheet's cells as the file had them, and can it keep them?
    /// </summary>
    /// <remarks>
    /// Excluded, as #702 lists them: formulas, whose text and cached values the calc engine owns;
    /// cell images and dynamic arrays, whose <c>vm</c> and <c>cm</c> indices point into metadata the
    /// save writes again; tables with a totals row, whose labels the writer takes from the table;
    /// and pivot tables.
    /// </remarks>
    private static bool CellsCanBeKept(XLWorksheet sheet)
    {
        var cells = sheet.Internals.CellsCollection;
        if (!cells.FormulaSlice.IsEmpty
            || sheet.PivotTables.Any<IXLPivotTable>()
            || sheet.Tables.Any<IXLTable>(table => table.ShowTotalsRow)
            || HasCellMetadata(cells.MiscSlice))
        {
            return false;
        }

        return !sheet.CellsChangedSinceLoad();
    }

    private static bool HasCellMetadata(Slice<XLMiscSliceContent> misc)
    {
        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(misc, Area.Full);
        while (enumerator.MoveNext())
        {
            ref readonly var content = ref enumerator.Current;
            if (content.CellMetaIndex is not null || content.ValueMetaIndex is not null || content.CellImage is not null)
                return true;
        }

        return false;
    }

    /// <summary>
    /// Works out what to declare on the kept start tag, so every prefix inside it resolves as it did
    /// in the file. Returns false when a reader does not find <c>&lt;sheetData&gt;</c> where
    /// <see cref="SheetDataLocator"/> put it.
    /// </summary>
    /// <remarks>
    /// The rewritten root keeps the declarations of the loaded one, so this usually adds nothing. It
    /// adds one where the writer drops a declaration, such as a default namespace beside a prefixed
    /// root, or binds a prefix otherwise.
    /// </remarks>
    private static bool TryGetDeclarationsToAdd(WorksheetPartBuffer part, Worksheet worksheet, out byte[] declarations)
    {
        declarations = [];

        Dictionary<string, string> rootDeclarations;
        Dictionary<string, string> ownDeclarations;
        using (var stream = part.OpenRead())
        using (var reader = PartXmlReader.Create(stream))
        {
            if (reader.MoveToContent() != XmlNodeType.Element)
                return false;

            rootDeclarations = ReadDeclarations(reader);
            if (!MoveToSheetData(reader, part.SheetData.Prefix))
                return false;

            ownDeclarations = ReadDeclarations(reader);
        }

        var written = GetWrittenScope(worksheet);
        var toAdd = new StringBuilder();
        foreach (var (prefix, uri) in rootDeclarations)
        {
            if (!ownDeclarations.ContainsKey(prefix) && !(written.TryGetValue(prefix, out var writtenUri) && writtenUri == uri))
                AppendDeclaration(toAdd, prefix, uri);
        }

        // An unprefixed name that was in no namespace in the file must not fall into a default
        // namespace the rewritten root declares.
        if (written.ContainsKey(string.Empty) && !rootDeclarations.ContainsKey(string.Empty)
                                               && !ownDeclarations.ContainsKey(string.Empty))
        {
            AppendDeclaration(toAdd, string.Empty, string.Empty);
        }

        declarations = Encoding.UTF8.GetBytes(toAdd.ToString());
        return true;
    }

    /// <summary>
    /// The namespace declarations on the element the reader is on, by prefix; the default namespace
    /// under an empty prefix.
    /// </summary>
    private static Dictionary<string, string> ReadDeclarations(XmlReader reader)
    {
        var declarations = new Dictionary<string, string>();
        if (!reader.MoveToFirstAttribute())
            return declarations;

        do
        {
            if (reader.Prefix == "xmlns")
                declarations[reader.LocalName] = reader.Value;
            else if (reader.Prefix.Length == 0 && reader.LocalName == "xmlns")
                declarations[string.Empty] = reader.Value;
        } while (reader.MoveToNextAttribute());

        reader.MoveToElement();
        return declarations;
    }

    /// <summary>
    /// From the root, moves to the first child named <c>sheetData</c>, and confirms it is the one
    /// <see cref="SheetDataLocator"/> found: in the main namespace, and written with the same prefix.
    /// </summary>
    private static bool MoveToSheetData(XmlReader reader, string prefix)
    {
        if (reader.IsEmptyElement)
            return false;

        reader.Read();
        while (!reader.EOF && reader.Depth >= 1)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            if (reader.LocalName == "sheetData")
                return reader.NamespaceURI == Main2006SsNs && reader.Prefix == prefix;

            reader.Skip();
        }

        return false;
    }

    /// <summary>
    /// The prefixes in scope inside the root that <c>WorksheetPartWriter.StreamToPart</c> writes for
    /// <paramref name="worksheet"/>: the declarations it writes, the root's own prefix, and those the
    /// writer declares for the root's attributes. The default namespace is declared only when the
    /// root is in it.
    /// </summary>
    private static Dictionary<string, string> GetWrittenScope(Worksheet worksheet)
    {
        var scope = new Dictionary<string, string>();
        foreach (var attribute in worksheet.GetAttributes())
        {
            if (!string.IsNullOrEmpty(attribute.Prefix) && attribute.Prefix != "xml")
                scope.TryAdd(attribute.Prefix, attribute.NamespaceUri);
        }

        foreach (var declaration in worksheet.NamespaceDeclarations)
        {
            if (!string.IsNullOrEmpty(declaration.Key))
                scope[declaration.Key] = declaration.Value;
        }

        scope[worksheet.Prefix] = worksheet.NamespaceUri;
        return scope;
    }

    private static void AppendDeclaration(StringBuilder builder, string prefix, string uri)
    {
        builder.Append(prefix.Length == 0 ? " xmlns=\"" : $" xmlns:{prefix}=\"");
        foreach (var c in uri)
        {
            builder.Append(c switch
            {
                '&' => "&amp;",
                '<' => "&lt;",
                '"' => "&quot;",
                _ => c.ToString(),
            });
        }

        builder.Append('"');
    }
}
