using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ConditionalFormats;

namespace XLibur.Excel.IO;

/// <summary>
/// Bundles the stylesheet and its sub-collections that are always passed together during loading.
/// </summary>
internal sealed record StylesheetData(
    Stylesheet? Stylesheet,
    NumberingFormats? NumberingFormats,
    Fills? Fills,
    Borders? Borders,
    Fonts? Fonts,
    Dictionary<int, DifferentialFormat> DifferentialFormats)
{
    /// <summary>
    /// <c>numFmtId</c> → format code for every custom format the workbook declares. Built once, so
    /// resolving a style's number format is a dictionary hit rather than a scan of
    /// <see cref="NumberingFormats"/> per style.
    /// </summary>
    /// <remarks>
    /// Derived from <see cref="NumberingFormats"/> in the initialiser rather than passed in, so it
    /// cannot fall out of step with the element it summarises. <c>StylesheetData</c> is
    /// constructed once per workbook load, so this is built once per load and not per worksheet.
    /// </remarks>
    internal Dictionary<int, string> CustomNumberFormats { get; } =
        BuildCustomNumberFormats(NumberingFormats);

    /// <summary>
    /// Admits an entry only when it has both an id and a non-empty format code, and keeps the
    /// first of any duplicated id.
    /// </summary>
    /// <remarks>
    /// Both rules preserve what the code this replaced did. Skipping a <c>&lt;numFmt&gt;</c> with
    /// no format code is equivalent to the scan it replaces, which accepted such an element and
    /// then fell through to the built-in-id branch because its format code was empty; a lookup
    /// that never admits the entry falls through by missing instead, and lands in the same place.
    /// Keeping the first of a duplicated id matches the scan's <c>FirstOrDefault</c> — and is why
    /// this uses <c>TryAdd</c>: the per-load dictionary this replaces used <c>Add</c>, so a file
    /// declaring the same <c>numFmtId</c> twice threw.
    /// </remarks>
    private static Dictionary<int, string> BuildCustomNumberFormats(NumberingFormats? numberingFormats)
    {
        var map = new Dictionary<int, string>();
        if (numberingFormats is null)
            return map;

        foreach (var nf in numberingFormats.Elements<NumberingFormat>())
        {
            var numberFormatId = checked((int?)nf.NumberFormatId?.Value);
            var formatCode = nf.FormatCode?.Value;
            if (numberFormatId is null || string.IsNullOrEmpty(formatCode))
                continue;

            map.TryAdd(numberFormatId.Value, formatCode);
        }

        return map;
    }
}

internal sealed class LoadContext
{
    /// <summary>
    /// Conditional formats for pivot tables, loaded from sheets. Key is sheet name, value is the
    /// conditional formats.
    /// </summary>
    private readonly Dictionary<string, List<XLConditionalFormat>> _pivotCfs = new(XLHelper.SheetComparer);

    internal void AddPivotTableCf(string sheetName, XLConditionalFormat conditionalFormat)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out var list))
        {
            list = new List<XLConditionalFormat>();
            _pivotCfs[sheetName] = list;
        }

        list.Add(conditionalFormat);
    }

    internal XLConditionalFormat GetPivotCf(string sheetName, int priority)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out var list))
            throw PivotCfNotFoundException(sheetName, priority);

        var pivotCf = list.SingleOrDefault(x => x.Priority == priority);
        if (pivotCf is null)
            throw PivotCfNotFoundException(sheetName, priority);

        return pivotCf;
    }

    internal XLNumberFormatValue? GetNumberFormat(int? numberFormatId)
    {
        if (numberFormatId is not { } id)
            return null;

        var key = StyleDecoder.NumberFormatKey(id, Styles, XLNumberFormatValue.Default.Key);
        return XLNumberFormatValue.FromKey(ref key);
    }

    private StyleValueCache? _styleCache;

    /// <summary>
    /// The stylesheet and its sub-collections, populated once from the workbook styles part.
    /// </summary>
    internal StylesheetData Styles { get; set; } = null!;

    /// <summary>
    /// Per-workbook cache of style values resolved from <c>cellXfs</c> indexes. Shared by every
    /// worksheet because <see cref="Styles"/> is workbook-global.
    /// </summary>
    internal StyleValueCache StyleCache => _styleCache ??= new StyleValueCache(Styles);

    /// <summary>
    /// Maps 1-based vm (value metadata) index to cell image info loaded from rich data parts.
    /// Populated by <see cref="RichDataReader"/>.
    /// </summary>
    internal Dictionary<uint, XLCellImage>? RichValueImages { get; set; }

    /// <summary>
    /// The set of 1-based cell-metadata (<c>cm</c>) indexes that reference the <c>XLDAPR</c>
    /// dynamic-array future-metadata type. A cell whose <c>cm</c> is in this set carries a
    /// dynamic-array formula (as opposed to a legacy CSE array). <c>null</c> when the workbook
    /// has no dynamic-array metadata. Populated once from the cell-metadata part.
    /// </summary>
    internal HashSet<uint>? DynamicArrayCmIndexes { get; set; }

    /// <summary>
    /// Populate <see cref="DynamicArrayCmIndexes"/> from the workbook's cell-metadata part by
    /// finding every cell-metadata record that references the <c>XLDAPR</c> type.
    /// </summary>
    internal void LoadDynamicArrayMetadata(Metadata? metadata)
    {
        if (metadata?.MetadataTypes is not { } metadataTypes)
            return;

        uint typeIndex = 0;
        uint xldaprTypeIndex = 0;
        foreach (var metadataType in metadataTypes.Elements<MetadataType>())
        {
            typeIndex++;
            if (metadataType.Name?.Value == "XLDAPR")
            {
                xldaprTypeIndex = typeIndex;
                break;
            }
        }

        if (xldaprTypeIndex == 0 || metadata.GetFirstChild<CellMetadata>() is not { } cellMetadata)
            return;

        uint cmIndex = 0;
        foreach (var block in cellMetadata.Elements<MetadataBlock>())
        {
            cmIndex++;

            // A block may hold several records; it is a dynamic-array cell if any of them
            // references the XLDAPR type (add the block index at most once).
            foreach (var record in block.Elements<MetadataRecord>())
            {
                if (record.TypeIndex?.Value == xldaprTypeIndex)
                {
                    (DynamicArrayCmIndexes ??= new HashSet<uint>()).Add(cmIndex);
                    break;
                }
            }
        }
    }

    private static Exception PivotCfNotFoundException(string sheetName, int priority)
    {
        return PartStructureException.ExpectedElementNotFound($"conditional formatting for pivot table in sheet {sheetName} with priority {priority}");
    }
}
