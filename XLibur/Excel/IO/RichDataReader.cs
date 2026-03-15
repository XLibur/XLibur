using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Drawings;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads rich data parts (rdrichvalue.xml, rdrichvaluestructure.xml, richValueRel.xml)
/// to populate in-cell images during workbook load.
/// </summary>
internal static class RichDataReader
{
    // Relationship types used to find the parts
    private const string RichValueRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
    private const string RichValueStructureRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
    private const string RichValueRelRelType = "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel";

    // XML namespaces
    private const string RvNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private const string RvsNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
    private const string RvrNs = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Load rich data parts from the workbook part. If successful, populates
    /// <see cref="LoadContext.RichValueImages"/> with vm-to-CellImage mapping.
    /// </summary>
    internal static void LoadRichData(WorkbookPart workbookPart, XLWorkbook workbook, LoadContext context)
    {
        // Find rich data parts by relationship type
        var richValuePart = FindPartByRelType(workbookPart, RichValueRelType);
        var richValueStructurePart = FindPartByRelType(workbookPart, RichValueStructureRelType);
        var richValueRelPart = FindPartByRelType(workbookPart, RichValueRelRelType);

        // If any of the required parts are missing, skip
        if (richValuePart is null || richValueStructurePart is null || richValueRelPart is null)
            return;

        // 1. Parse structure to find _localImage structure index
        var localImageStructureIndex = FindLocalImageStructureIndex(richValueStructurePart);
        if (localImageStructureIndex < 0)
            return;

        // 2. Parse richValueRel to get image relationship IDs in order
        var imageRelIds = ParseRichValueRel(richValueRelPart);

        // 3. Load image bytes from relationship targets into the workbook store
        var relIndexToImageStoreIndex = new Dictionary<int, int>();
        for (var i = 0; i < imageRelIds.Count; i++)
        {
            var relId = imageRelIds[i];
            if (!richValueRelPart.TryGetPartById(relId, out var imagePart))
                continue;

            var format = DetectFormat(imagePart);
            var ms = new MemoryStream();
            using (var imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            {
                imageStream.CopyTo(ms);
            }

            ms.Position = 0;
            var storeIndex = workbook.InCellImages.AddDirect(ms, format);
            relIndexToImageStoreIndex[i] = storeIndex;
        }

        // 4. Parse rdrichvalue.xml to build rv index -> (imageStoreIndex, altText)
        var rvEntries = ParseRichValues(richValuePart, localImageStructureIndex, relIndexToImageStoreIndex);
        if (rvEntries.Count == 0)
            return;

        // 5. Parse metadata.xml to map vm (1-based) -> rv index (0-based)
        var vmToRvIndex = ParseValueMetadata(workbookPart);
        if (vmToRvIndex.Count == 0)
            return;

        // 6. Build final vm -> CellImage map
        var richValueImages = new Dictionary<uint, XLCellImage>();
        foreach (var (vm, rvIndex) in vmToRvIndex)
        {
            if (rvIndex >= 0 && rvIndex < rvEntries.Count && rvEntries[rvIndex] is { } entry)
            {
                richValueImages[vm] = entry;
            }
        }

        if (richValueImages.Count > 0)
            context.RichValueImages = richValueImages;
    }

    private static OpenXmlPart? FindPartByRelType(WorkbookPart workbookPart, string relType)
    {
        foreach (var idPart in workbookPart.Parts)
        {
            if (idPart.OpenXmlPart.RelationshipType == relType)
                return idPart.OpenXmlPart;
        }

        return null;
    }

    /// <summary>
    /// Find the 0-based index of the _localImage structure in rdrichvaluestructure.xml.
    /// </summary>
    private static int FindLocalImageStructureIndex(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream);

        var structureIndex = 0;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "s")
            {
                var typeAttr = reader.GetAttribute("t");
                if (typeAttr == "_localImage")
                    return structureIndex;

                structureIndex++;
            }
        }

        return -1;
    }

    /// <summary>
    /// Parse richValueRel.xml to get ordered list of relationship IDs.
    /// </summary>
    private static List<string> ParseRichValueRel(OpenXmlPart part)
    {
        var relIds = new List<string>();

        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "rel")
            {
                // The r:id attribute is in the relationships namespace
                var id = reader.GetAttribute("id", RNs);
                if (id is not null)
                    relIds.Add(id);
            }
        }

        return relIds;
    }

    /// <summary>
    /// Parse rdrichvalue.xml. Returns a list indexed by rv index, with null for non-image entries.
    /// </summary>
    private static List<XLCellImage?> ParseRichValues(
        OpenXmlPart part,
        int localImageStructureIndex,
        Dictionary<int, int> relIndexToImageStoreIndex)
    {
        var entries = new List<XLCellImage?>();

        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream);

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "rv")
                continue;

            var sAttr = reader.GetAttribute("s");
            if (sAttr is null || !int.TryParse(sAttr, out var structureIndex) ||
                structureIndex != localImageStructureIndex)
            {
                entries.Add(null);
                continue;
            }

            var values = ReadRvChildValues(reader);
            entries.Add(BuildCellImageFromValues(values, relIndexToImageStoreIndex));
        }

        return entries;
    }

    private static List<string> ReadRvChildValues(XmlReader reader)
    {
        var values = new List<string>();
        using var rvReader = reader.ReadSubtree();
        while (rvReader.Read())
        {
            if (rvReader.NodeType == XmlNodeType.Element && rvReader.LocalName == "v")
            {
                values.Add(rvReader.ReadElementContentAsString());
            }
        }

        return values;
    }

    private static XLCellImage? BuildCellImageFromValues(List<string> values,
        Dictionary<int, int> relIndexToImageStoreIndex)
    {
        if (values.Count >= 1 && int.TryParse(values[0], out var relIndex) &&
            relIndexToImageStoreIndex.TryGetValue(relIndex, out var imageStoreIndex))
        {
            var altText = values.Count >= 3 ? values[2] : string.Empty;
            return new XLCellImage(imageStoreIndex, altText);
        }

        return null;
    }

    /// <summary>
    /// Parse metadata.xml to find XLRICHVALUE type and map valueMetadata vm (1-based) -> rv index.
    /// </summary>
    private static Dictionary<uint, int> ParseValueMetadata(WorkbookPart workbookPart)
    {
        var result = new Dictionary<uint, int>();

        var cellMetadataPart = workbookPart.CellMetadataPart;
        if (cellMetadataPart?.Metadata is null)
            return result;

        var metadata = cellMetadataPart.Metadata;
        var metadataTypes = metadata.MetadataTypes;
        if (metadataTypes is null)
            return result;

        // Find XLRICHVALUE type index (1-based)
        uint? richValueTypeIndex = null;
        uint typeIdx = 1;
        foreach (var mt in metadataTypes.Elements<MetadataType>())
        {
            if (mt.Name?.Value == "XLRICHVALUE")
            {
                richValueTypeIndex = typeIdx;
                break;
            }

            typeIdx++;
        }

        if (richValueTypeIndex is null)
            return result;

        // Find valueMetadata records referencing this type
        var valueMeta = metadata.GetFirstChild<ValueMetadata>();
        if (valueMeta is null)
            return result;

        uint vmIndex = 1; // 1-based
        foreach (var bk in valueMeta.Elements<MetadataBlock>())
        {
            var rc = bk.GetFirstChild<MetadataRecord>();
            if (rc?.TypeIndex?.Value == richValueTypeIndex)
            {
                var fmIndex = (int)(rc.Val?.Value ?? 0);
                var rvIndex = FindRvIndexFromFutureMetadata(metadata, richValueTypeIndex.Value, fmIndex);
                result[vmIndex] = rvIndex;
            }

            vmIndex++;
        }

        return result;
    }

    /// <summary>
    /// Find the rv index from futureMetadata block at the given index for XLRICHVALUE type.
    /// </summary>
    private static int FindRvIndexFromFutureMetadata(Metadata metadata, uint typeIndex, int blockIndex)
    {
        // Find the futureMetadata element with name "XLRICHVALUE"
        foreach (var fm in metadata.Elements<FutureMetadata>())
        {
            if (fm.Name?.Value != "XLRICHVALUE")
                continue;

            return FindRvIndexInFutureMetadataBlocks(fm, blockIndex);
        }

        // Fallback: use the block index directly
        return blockIndex;
    }

    private static int FindRvIndexInFutureMetadataBlocks(FutureMetadata fm, int blockIndex)
    {
        var idx = 0;
        foreach (var block in fm.Elements<FutureMetadataBlock>())
        {
            if (idx == blockIndex)
                return ExtractRvIndexFromBlock(block, blockIndex);

            idx++;
        }

        return blockIndex;
    }

    private static int ExtractRvIndexFromBlock(FutureMetadataBlock block, int blockIndex)
    {
        var extList = block.GetFirstChild<ExtensionList>();
        if (extList is not null)
        {
            foreach (var ext in extList.Elements<Extension>())
            {
                foreach (var child in ext.ChildElements)
                {
                    var iAttr = child.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == "i");
                    if (iAttr.Value is not null && int.TryParse(iAttr.Value, out var rvIndex))
                        return rvIndex;
                }
            }
        }

        // If no explicit rvb element found, the blockIndex itself is the rv index
        return blockIndex;
    }

    private static XLPictureFormat DetectFormat(OpenXmlPart imagePart)
    {
        var contentType = imagePart.ContentType;
        return contentType switch
        {
            "image/png" => XLPictureFormat.Png,
            "image/jpeg" => XLPictureFormat.Jpeg,
            "image/gif" => XLPictureFormat.Gif,
            "image/bmp" => XLPictureFormat.Bmp,
            "image/tiff" => XLPictureFormat.Tiff,
            "image/x-icon" => XLPictureFormat.Icon,
            "image/x-pcx" => XLPictureFormat.Pcx,
            "image/x-emf" => XLPictureFormat.Emf,
            "image/x-wmf" => XLPictureFormat.Wmf,
            "image/webp" => XLPictureFormat.Webp,
            "image/svg+xml" => XLPictureFormat.Svg,
            _ => XLPictureFormat.Unknown,
        };
    }
}
