using System.Collections.Generic;
using System.Globalization;
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
    /// <summary>
    /// Load rich data parts from the workbook part. If successful, populates
    /// <see cref="LoadContext.RichValueImages"/> with vm-to-CellImage mapping.
    /// </summary>
    internal static void LoadRichData(WorkbookPart workbookPart, XLWorkbook workbook, LoadContext context)
    {
        // Find rich data parts by relationship type. Excel relates the rich value relationship
        // part by its 2022 type; XLibur used the 2017 one before #739.
        var richValuePart = FindPartByRelType(workbookPart, RichDataWriter.RichValueRelType);
        var richValueStructurePart = FindPartByRelType(workbookPart, RichDataWriter.RichValueStructureRelType);
        var richValueRelPart = FindPartByRelType(workbookPart, RichDataWriter.RichValueRelRelType)
                               ?? FindPartByRelType(workbookPart, RichDataWriter.LegacyRichValueRelRelType);

        if (richValuePart is null)
            return;

        // 1. Parse structures to find each _localImage structure and the positions of its keys
        var localImageStructures = richValueStructurePart is null
            ? new Dictionary<int, LocalImageLayout>()
            : FindLocalImageStructures(richValueStructurePart);

        // 2. Parse rdrichvalue.xml to build rv index -> (rel index, alt text)
        var richValues = ParseRichValues(richValuePart, localImageStructures);
        if (richValues.Count == 0)
            return;

        // 3. Parse richValueRel to get image relationship IDs in order
        var imageRelIds = richValueRelPart is null ? [] : ParseRichValueRel(richValueRelPart);

        // The save rewrites the rich data parts from the images it knows and drops every other
        // rich value. Unless each rich value is a local image whose relationship resolves, the
        // workbook's rich data is left as the file has it, and so are its cells, and no image
        // can be added: the rewrite would drop the other rich values and leave their cells'
        // vm indexes naming the new images, or nothing.
        if (richValueRelPart is null ||
            !richValues.TrueForAll(rv => rv is { } v && v.RelIndex >= 0 && v.RelIndex < imageRelIds.Count &&
                                         richValueRelPart.TryGetPartById(imageRelIds[v.RelIndex], out _)))
        {
            workbook.KeepsLoadedRichData = true;
            return;
        }

        // 4. Load image bytes from relationship targets into the workbook store
        var relIndexToImageStoreIndex = LoadImages(richValueRelPart, imageRelIds, workbook);
        var rvEntries = richValues
            .Select(rv => (XLCellImage?)new XLCellImage(relIndexToImageStoreIndex[rv!.Value.RelIndex], rv.Value.AltText))
            .ToList();

        // 5. Parse metadata.xml to map vm (1-based) -> rv index (0-based)
        var vmToRvIndex = ParseValueMetadata(workbookPart);
        if (vmToRvIndex.Count == 0)
            return;

        // 6. Build final vm -> CellImage map
        var richValueImages = MapValueMetadataToImages(vmToRvIndex, rvEntries);
        if (richValueImages.Count > 0)
            context.RichValueImages = richValueImages;
    }

    /// <summary>
    /// Load the image of each relationship into the workbook store. Returns a map of rel index -> image store index.
    /// </summary>
    private static Dictionary<int, int> LoadImages(OpenXmlPart richValueRelPart, List<string> imageRelIds,
        XLWorkbook workbook)
    {
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

        return relIndexToImageStoreIndex;
    }

    private static Dictionary<uint, XLCellImage> MapValueMetadataToImages(Dictionary<uint, int> vmToRvIndex,
        List<XLCellImage?> rvEntries)
    {
        var richValueImages = new Dictionary<uint, XLCellImage>();
        foreach (var (vm, rvIndex) in vmToRvIndex)
        {
            if (rvIndex >= 0 && rvIndex < rvEntries.Count && rvEntries[rvIndex] is { } entry)
            {
                richValueImages[vm] = entry;
            }
        }

        return richValueImages;
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
    /// The positions, among the values of a rich value, of a local image's keys.
    /// </summary>
    /// <param name="RelKey">Position of the image's index into richValueRel.xml.</param>
    /// <param name="TextKey">Position of the alt text, or -1 if the structure has none.</param>
    private readonly record struct LocalImageLayout(int RelKey, int TextKey);

    /// <summary>
    /// Find every _localImage structure in rdrichvaluestructure.xml, by 0-based structure index.
    /// Excel writes one structure per set of keys, e.g. one with alt text and one without, so a
    /// workbook can have several.
    /// </summary>
    private static Dictionary<int, LocalImageLayout> FindLocalImageStructures(OpenXmlPart part)
    {
        var structures = new Dictionary<int, LocalImageLayout>();

        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = PartXmlReader.CreateVerbatim(stream);

        var structureIndex = 0;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "s")
                continue;

            if (reader.GetAttribute("t") == "_localImage" && ReadLocalImageLayout(reader) is { } layout)
                structures[structureIndex] = layout;

            structureIndex++;
        }

        return structures;
    }

    private static LocalImageLayout? ReadLocalImageLayout(XmlReader reader)
    {
        var relKey = -1;
        var textKey = -1;
        var keyIndex = 0;
        using var structureReader = reader.ReadSubtree();
        while (structureReader.Read())
        {
            if (structureReader.NodeType != XmlNodeType.Element || structureReader.LocalName != "k")
                continue;

            switch (structureReader.GetAttribute("n"))
            {
                // Excel's name, and the one XLibur wrote before #739.
                case "_rvRel:LocalImageIdentifier" or "_rvRel:LocalImageFileUri":
                    relKey = keyIndex;
                    break;
                case "Text" or "text":
                    textKey = keyIndex;
                    break;
            }

            keyIndex++;
        }

        return relKey >= 0 ? new LocalImageLayout(relKey, textKey) : null;
    }

    /// <summary>
    /// Parse richValueRel.xml to get ordered list of relationship IDs.
    /// </summary>
    private static List<string> ParseRichValueRel(OpenXmlPart part)
    {
        var relIds = new List<string>();

        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = PartXmlReader.CreateVerbatim(stream);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "rel")
            {
                // The r:id attribute is in the relationships namespace
                var id = reader.GetAttribute("id", OpenXmlConst.RelationshipsNs);
                if (id is not null)
                    relIds.Add(id);
            }
        }

        return relIds;
    }

    /// <summary>
    /// Parse rdrichvalue.xml. Returns a list indexed by rv index, with null for entries that are
    /// not local images.
    /// </summary>
    private static List<(int RelIndex, string AltText)?> ParseRichValues(
        OpenXmlPart part,
        Dictionary<int, LocalImageLayout> localImageStructures)
    {
        var entries = new List<(int RelIndex, string AltText)?>();

        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = PartXmlReader.CreateVerbatim(stream);

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "rv")
                continue;

            var sAttr = reader.GetAttribute("s");
            if (sAttr is null || !int.TryParse(sAttr, out var structureIndex) ||
                !localImageStructures.TryGetValue(structureIndex, out var layout))
            {
                entries.Add(null);
                continue;
            }

            var values = ReadRvChildValues(reader);
            entries.Add(ReadLocalImage(values, layout));
        }

        return entries;
    }

    private static List<string> ReadRvChildValues(XmlReader reader)
    {
        var values = new List<string>();
        using var rvReader = reader.ReadSubtree();
        while (!rvReader.EOF)
        {
            // ReadElementContentAsString leaves the reader on the node after the element, which is
            // the next <v>. A Read() after it skipped that value, so the alt text never loaded.
            if (rvReader.NodeType == XmlNodeType.Element && rvReader.LocalName == "v")
                values.Add(rvReader.ReadElementContentAsString());
            else
                rvReader.Read();
        }

        return values;
    }

    private static (int RelIndex, string AltText)? ReadLocalImage(List<string> values, LocalImageLayout layout)
    {
        if (layout.RelKey >= values.Count ||
            !int.TryParse(values[layout.RelKey], NumberStyles.Integer, CultureInfo.InvariantCulture, out var relIndex))
            return null;

        var altText = layout.TextKey >= 0 && layout.TextKey < values.Count ? values[layout.TextKey] : string.Empty;
        return (relIndex, altText);
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
                var rvIndex = FindRvIndexFromFutureMetadata(metadata, fmIndex);
                result[vmIndex] = rvIndex;
            }

            vmIndex++;
        }

        return result;
    }

    /// <summary>
    /// Find the rv index from futureMetadata block at the given index for XLRICHVALUE type.
    /// </summary>
    private static int FindRvIndexFromFutureMetadata(Metadata metadata, int blockIndex)
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
        // If no explicit rvb element found, the blockIndex itself is the rv index
        var extList = block.GetFirstChild<ExtensionList>();
        if (extList is null)
            return blockIndex;

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
