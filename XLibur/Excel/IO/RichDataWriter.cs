using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the four rich data XML parts required for in-cell images:
/// rdrichvalue.xml, rdrichvaluestructure.xml, richValueRel.xml, rdRichValueTypes.xml.
/// Also writes image binary parts as children of the richValueRel part.
/// </summary>
internal static class RichDataWriter
{
    // Relationship types
    private const string RichValueRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
    private const string RichValueStructureRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
    private const string RichValueRelRelType = "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel";
    private const string RichValueTypesRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";

    // Content types
    private const string RichValueContentType = "application/vnd.ms-excel.rdrichvalue+xml";
    private const string RichValueStructureContentType = "application/vnd.ms-excel.rdrichvaluestructure+xml";
    private const string RichValueRelContentType = "application/vnd.ms-excel.richValueRel+xml";
    private const string RichValueTypesContentType = "application/vnd.ms-excel.rdrichvaluetypes+xml";

    // XML namespaces
    private const string RvNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private const string RvsNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
    private const string RvrNs = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Entry for a single rich value (one per cell with in-cell image).
    /// </summary>
    internal readonly struct RichValueEntry
    {
        internal readonly int ImageStoreIndex;
        internal readonly string AltText;

        internal RichValueEntry(int imageStoreIndex, string altText)
        {
            ImageStoreIndex = imageStoreIndex;
            AltText = altText;
        }
    }

    /// <summary>
    /// Write all four rich data parts. Returns the list of image part relationship IDs
    /// used to wire up the richValueRel part.
    /// </summary>
    /// <param name="workbookPart">Workbook part to add extended parts to.</param>
    /// <param name="imageStore">Workbook-level image blob store.</param>
    /// <param name="entries">One entry per cell with an in-cell image.</param>
    /// <param name="relIdGenerator">Generator for unique relationship IDs.</param>
    internal static void WriteRichDataParts(
        WorkbookPart workbookPart,
        XLInCellImageStore imageStore,
        IReadOnlyList<RichValueEntry> entries,
        XLWorkbook.RelIdGenerator relIdGenerator)
    {
        // Collect unique image indices in order of first appearance
        var uniqueImageIndices = new List<int>();
        var imageIndexToRelIndex = new Dictionary<int, int>();

        foreach (var entry in entries)
        {
            if (!imageIndexToRelIndex.ContainsKey(entry.ImageStoreIndex))
            {
                imageIndexToRelIndex[entry.ImageStoreIndex] = uniqueImageIndices.Count;
                uniqueImageIndices.Add(entry.ImageStoreIndex);
            }
        }

        // Remove any existing rich data parts (by relationship type)
        RemoveExistingRichDataParts(workbookPart);

        // 1. richValueRel.xml - must be first because image binary parts are children of it
        var richValueRelPart = workbookPart.AddExtendedPart(
            RichValueRelRelType, RichValueRelContentType,
            ".xml", relIdGenerator.GetNext(XLWorkbook.RelType.Workbook));

        // Add image binary parts to the richValueRel part
        var imageRelIds = new string[uniqueImageIndices.Count];
        for (var i = 0; i < uniqueImageIndices.Count; i++)
        {
            var (stream, format) = imageStore.GetImage(uniqueImageIndices[i]);
            var imageRelId = $"rId{i + 1}";
            var partTypeInfo = format.ToOpenXml();
            var imagePart = richValueRelPart.AddNewPart<ImagePart>(partTypeInfo.ContentType, imageRelId);
            stream.Position = 0;
            imagePart.FeedData(stream);
            imageRelIds[i] = imageRelId;
        }

        WriteRichValueRelXml(richValueRelPart, imageRelIds);

        // 2. rdrichvalue.xml
        var richValuePart = workbookPart.AddExtendedPart(
            RichValueRelType, RichValueContentType,
            ".xml", relIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
        WriteRichValueXml(richValuePart, entries, imageIndexToRelIndex);

        // 3. rdrichvaluestructure.xml
        var richValueStructurePart = workbookPart.AddExtendedPart(
            RichValueStructureRelType, RichValueStructureContentType,
            ".xml", relIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
        WriteRichValueStructureXml(richValueStructurePart);

        // 4. rdRichValueTypes.xml
        var richValueTypesPart = workbookPart.AddExtendedPart(
            RichValueTypesRelType, RichValueTypesContentType,
            ".xml", relIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
        WriteRichValueTypesXml(richValueTypesPart);
    }

    private static void RemoveExistingRichDataParts(WorkbookPart workbookPart)
    {
        var richDataRelTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            RichValueRelType,
            RichValueStructureRelType,
            RichValueRelRelType,
            RichValueTypesRelType,
        };

        var partsToRemove = new List<OpenXmlPart>();
        foreach (var idPartPair in workbookPart.Parts)
        {
            if (richDataRelTypes.Contains(idPartPair.OpenXmlPart.RelationshipType))
            {
                partsToRemove.Add(idPartPair.OpenXmlPart);
            }
        }

        foreach (var part in partsToRemove)
            workbookPart.DeletePart(part);
    }

    /// <summary>
    /// richValueRel.xml - maps rel indices to image binary parts.
    /// <code>
    /// &lt;richValueRels xmlns="..." xmlns:r="..."&gt;
    ///   &lt;rel r:id="rId1" /&gt;
    ///   ...
    /// &lt;/richValueRels&gt;
    /// </code>
    /// </summary>
    private static void WriteRichValueRelXml(OpenXmlPart part, string[] imageRelIds)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = XmlWriter.Create(stream, XmlSettings());

        w.WriteStartDocument(true);
        w.WriteStartElement("richValueRels", RvrNs);
        w.WriteAttributeString("xmlns", "r", null, RNs);

        foreach (var relId in imageRelIds)
        {
            w.WriteStartElement("rel", RvrNs);
            w.WriteAttributeString("id", RNs, relId);
            w.WriteEndElement(); // rel
        }

        w.WriteEndElement(); // richValueRels
    }

    /// <summary>
    /// rdrichvalue.xml - one rv entry per cell with in-cell image.
    /// <code>
    /// &lt;rvData xmlns="..." count="N"&gt;
    ///   &lt;rv s="0"&gt;
    ///     &lt;v&gt;relIndex&lt;/v&gt;
    ///     &lt;v&gt;5&lt;/v&gt;
    ///     &lt;v&gt;altText&lt;/v&gt;
    ///   &lt;/rv&gt;
    ///   ...
    /// &lt;/rvData&gt;
    /// </code>
    /// </summary>
    private static void WriteRichValueXml(
        OpenXmlPart part,
        IReadOnlyList<RichValueEntry> entries,
        Dictionary<int, int> imageIndexToRelIndex)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = XmlWriter.Create(stream, XmlSettings());

        w.WriteStartDocument(true);
        w.WriteStartElement("rvData", RvNs);
        w.WriteAttributeString("count", entries.Count.ToString());

        for (var i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            var relIndex = imageIndexToRelIndex[entry.ImageStoreIndex];

            w.WriteStartElement("rv", RvNs);
            w.WriteAttributeString("s", "0"); // structure index 0 = _localImage

            // Key 0: _rvRel:LocalImageFileUri (rel index)
            w.WriteElementString("v", RvNs, relIndex.ToString());
            // Key 1: CalcOrigin
            w.WriteElementString("v", RvNs, "5");
            // Key 2: text (alt text)
            w.WriteElementString("v", RvNs, entry.AltText);

            w.WriteEndElement(); // rv
        }

        w.WriteEndElement(); // rvData
    }

    /// <summary>
    /// rdrichvaluestructure.xml - defines the _localImage structure schema.
    /// </summary>
    private static void WriteRichValueStructureXml(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = XmlWriter.Create(stream, XmlSettings());

        w.WriteStartDocument(true);
        w.WriteStartElement("rvStructures", RvsNs);
        w.WriteAttributeString("count", "1");

        w.WriteStartElement("s", RvsNs);
        w.WriteAttributeString("t", "_localImage");

        // Key 0: _rvRel:LocalImageFileUri
        w.WriteStartElement("k", RvsNs);
        w.WriteAttributeString("n", "_rvRel:LocalImageFileUri");
        w.WriteAttributeString("t", "i");
        w.WriteEndElement(); // k

        // Key 1: CalcOrigin
        w.WriteStartElement("k", RvsNs);
        w.WriteAttributeString("n", "CalcOrigin");
        w.WriteAttributeString("t", "i");
        w.WriteEndElement(); // k

        // Key 2: text (alt text)
        w.WriteStartElement("k", RvsNs);
        w.WriteAttributeString("n", "text");
        w.WriteAttributeString("t", "s");
        w.WriteEndElement(); // k

        w.WriteEndElement(); // s

        w.WriteEndElement(); // rvStructures
    }

    /// <summary>
    /// rdRichValueTypes.xml - global key flags (static content).
    /// </summary>
    private static void WriteRichValueTypesXml(OpenXmlPart part)
    {
        const string rvTypesNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";

        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = XmlWriter.Create(stream, XmlSettings());

        w.WriteStartDocument(true);
        w.WriteStartElement("rvTypesInfo", rvTypesNs);

        w.WriteStartElement("global", rvTypesNs);

        w.WriteStartElement("keyFlags", rvTypesNs);

        // _rvRel:LocalImageFileUri key flag
        w.WriteStartElement("key", rvTypesNs);
        w.WriteAttributeString("name", "_rvRel:LocalImageFileUri");

        w.WriteStartElement("flag", rvTypesNs);
        w.WriteAttributeString("name", "ExcludeFromFile");
        w.WriteAttributeString("value", "1");
        w.WriteEndElement(); // flag

        w.WriteStartElement("flag", rvTypesNs);
        w.WriteAttributeString("name", "ExcludeFromCalcComparison");
        w.WriteAttributeString("value", "1");
        w.WriteEndElement(); // flag

        w.WriteEndElement(); // key

        w.WriteEndElement(); // keyFlags

        w.WriteEndElement(); // global

        w.WriteEndElement(); // rvTypesInfo
    }

    private static XmlWriterSettings XmlSettings() => new()
    {
        Encoding = System.Text.Encoding.UTF8,
        CloseOutput = true,
    };
}
