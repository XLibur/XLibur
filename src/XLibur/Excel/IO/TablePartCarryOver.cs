using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Carries what a loaded table part holds and the model does not over to the part a save writes in
/// its place (#904): the sort state, the header, border and totals formats, the named cell styles,
/// the totals a hidden totals row remembers, the insert row, and attributes and extensions XLibur
/// does not know.
/// </summary>
/// <remarks>
/// <para>
/// A column is matched to its loaded column by <c>tableColumn/@id</c>, which a save keeps (#903), so
/// a renamed or moved column keeps what it had and a removed one takes it away.
/// </para>
/// <para>
/// What depends on where the columns are is dropped once the model changed them, rather than written
/// over a table it no longer describes. The sort state and the insert row need the same range and the
/// same columns in the same order, since a sort condition names its column by cell reference. The
/// totals of a hidden totals row need the same column names, since a totals formula names columns by
/// name.
/// </para>
/// </remarks>
internal static class TablePartCarryOver
{
    private static readonly HashSet<string> ModelTableAttributes =
        ["id", "name", "displayName", "ref", "headerRowCount", "totalsRowCount", "totalsRowShown"];

    private static readonly HashSet<string> ModelColumnAttributes = ["id", "name", "dataDxfId"];

    private static readonly HashSet<string> ShownTotalsRowAttributes = ["totalsRowFunction", "totalsRowLabel"];

    private static readonly HashSet<string> LoadedRangeAttributes = ["insertRow", "insertRowShift"];

    internal static void Apply(Table loaded, Table written, bool showTotalsRow, LoadedDifferentialFormats? dxfs)
    {
        var sameRange = loaded.Reference?.Value == written.Reference?.Value &&
                        Columns(loaded).Select(c => c.Id?.Value).SequenceEqual(Columns(written).Select(c => c.Id?.Value));
        var keepTotals = !showTotalsRow &&
                         Columns(loaded).Select(c => c.Name?.Value).SequenceEqual(Columns(written).Select(c => c.Name?.Value));

        CopyAttributes(loaded, written,
            name => !ModelTableAttributes.Contains(name) && (sameRange || !LoadedRangeAttributes.Contains(name)), dxfs);

        foreach (var declaration in loaded.NamespaceDeclarations)
        {
            if (written.LookupNamespace(declaration.Key) is null)
                written.AddNamespaceDeclaration(declaration.Key, declaration.Value);
        }

        if (loaded.MCAttributes is { } mcAttributes)
            written.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = mcAttributes.Ignorable };

        if (!showTotalsRow)
            written.TotalsRowShown = loaded.TotalsRowShown?.Value is { } shown ? shown : null;

        if (sameRange && loaded.SortState?.CloneNode(true) is SortState sortState && TryRemapDxfIds(sortState, dxfs))
            written.SortState = sortState;

        if (loaded.TableExtensionList is { } extensions)
            written.TableExtensionList = (TableExtensionList)extensions.CloneNode(true);

        ApplyColumns(loaded, written, keepTotals, dxfs);
    }

    private static void ApplyColumns(Table loaded, Table written, bool keepTotals, LoadedDifferentialFormats? dxfs)
    {
        var loadedColumns = new Dictionary<uint, TableColumn>();
        foreach (var column in Columns(loaded))
        {
            if (column.Id?.Value is { } id)
                loadedColumns.TryAdd(id, column);
        }

        foreach (var column in Columns(written))
        {
            if (column.Id?.Value is { } id && loadedColumns.TryGetValue(id, out var loadedColumn))
                ApplyColumn(loadedColumn, column, keepTotals, dxfs);
        }
    }

    private static void ApplyColumn(TableColumn loaded, TableColumn written, bool keepTotals,
        LoadedDifferentialFormats? dxfs)
    {
        CopyAttributes(loaded, written,
            name => !ModelColumnAttributes.Contains(name) && (keepTotals || !ShownTotalsRowAttributes.Contains(name)), dxfs);

        foreach (var child in loaded.ChildElements)
        {
            if (child is CalculatedColumnFormula || (child is TotalsRowFormula && !keepTotals))
                continue;

            written.AppendChild(child.CloneNode(true));
        }
    }

    private static IEnumerable<TableColumn> Columns(Table table) =>
        table.TableColumns?.Elements<TableColumn>() ?? [];

    private static void CopyAttributes(OpenXmlElement from, OpenXmlElement to, Func<string, bool> carry,
        LoadedDifferentialFormats? dxfs)
    {
        foreach (var attribute in from.GetAttributes())
        {
            var isMain = string.IsNullOrEmpty(attribute.NamespaceUri);
            if (isMain && !carry(attribute.LocalName))
                continue;

            if (to.GetAttributes().Any(a => a.LocalName == attribute.LocalName && a.NamespaceUri == attribute.NamespaceUri))
                continue;

            if (!isMain || !IsDxfId(attribute.LocalName))
            {
                to.SetAttribute(attribute);
                continue;
            }

            if (Remap(attribute.Value, dxfs) is { } dxfId)
                to.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, dxfId));
        }
    }

    private static bool TryRemapDxfIds(OpenXmlElement element, LoadedDifferentialFormats? dxfs)
    {
        foreach (var descendant in element.Descendants().Prepend(element))
        {
            foreach (var attribute in descendant.GetAttributes())
            {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri) || !IsDxfId(attribute.LocalName))
                    continue;

                if (Remap(attribute.Value, dxfs) is not { } dxfId)
                    return false;

                descendant.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, dxfId));
            }
        }

        return true;
    }

    private static bool IsDxfId(string localName) =>
        localName == "dxfId" || localName.EndsWith("DxfId", StringComparison.Ordinal);

    private static string? Remap(string? loadedId, LoadedDifferentialFormats? dxfs)
    {
        if (dxfs is null || !uint.TryParse(loadedId, NumberStyles.None, CultureInfo.InvariantCulture, out var id))
            return null;

        return dxfs.WrittenIndexOf(id)?.ToString(CultureInfo.InvariantCulture);
    }
}
