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
    Dictionary<int, DifferentialFormat> DifferentialFormats);

internal sealed class LoadContext
{
    /// <summary>
    /// Conditional formats for pivot tables, loaded from sheets. Key is sheet name, value is the
    /// conditional formats.
    /// </summary>
    private readonly Dictionary<string, List<XLConditionalFormat>> _pivotCfs = new(XLHelper.SheetComparer);

    /// <summary>
    /// A dictionary of styles from <c>styles.xml</c>. Used in other places that reference number style by id reference.
    /// </summary>
    private readonly Dictionary<int, string> _numberFormats = new();

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

    internal void LoadNumberFormats(NumberingFormats? numberingFormats)
    {
        if (numberingFormats is null)
            return;

        foreach (var nf in numberingFormats.ChildElements.Cast<NumberingFormat>())
        {
            var numberFormatId = checked((int?)nf.NumberFormatId?.Value);
            var formatCode = nf.FormatCode?.Value;
            if (numberFormatId is null || string.IsNullOrEmpty(formatCode))
                continue;

            _numberFormats.Add(numberFormatId.Value, formatCode);
        }
    }

    internal XLNumberFormatValue? GetNumberFormat(int? numberFormatId)
    {
        if (numberFormatId is null)
        {
            return null;
        }

        if (_numberFormats.TryGetValue(numberFormatId.Value, out var formatCode))
        {
            var customFormatKey = new XLNumberFormatKey
            {
                NumberFormatId = -1,
                Format = formatCode,
            };
            return XLNumberFormatValue.FromKey(ref customFormatKey);
        }

        var predefinedFormatKey = new XLNumberFormatKey
        {
            NumberFormatId = numberFormatId.Value,
            Format = string.Empty,
        };
        return XLNumberFormatValue.FromKey(ref predefinedFormatKey);
    }

    /// <summary>
    /// The stylesheet and its sub-collections, populated once from the workbook styles part.
    /// </summary>
    internal StylesheetData Styles { get; set; } = null!;

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
            if (block.GetFirstChild<MetadataRecord>()?.TypeIndex?.Value == xldaprTypeIndex)
                (DynamicArrayCmIndexes ??= new HashSet<uint>()).Add(cmIndex);
        }
    }

    private static Exception PivotCfNotFoundException(string sheetName, int priority)
    {
        return PartStructureException.ExpectedElementNotFound($"conditional formatting for pivot table in sheet {sheetName} with priority {priority}");
    }
}
