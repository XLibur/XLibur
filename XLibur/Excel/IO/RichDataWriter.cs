using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the rich data XML parts for in-cell images and for errors newer than Excel 2007:
/// rdrichvalue.xml, rdrichvaluestructure.xml, rdRichValueTypes.xml and, for images, richValueRel.xml
/// with the image binary parts as its children.
/// </summary>
internal static class RichDataWriter
{
    // Relationship types, content types, namespaces and key names are the ones Excel writes for a
    // picture placed in a cell (#739). Excel refuses to open the file if the structure part is in
    // the richdata2 namespace or names its image key otherwise, and shows #FIELD! in the cell if
    // the rich value relationship part is related by its 2017 type.
    internal const string RichValueRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
    internal const string RichValueStructureRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
    internal const string RichValueRelRelType = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
    internal const string RichValueTypesRelType = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";

    /// <summary>
    /// The type XLibur related its rich value relationship part by before #739. Excel does not
    /// read it, but a file written then still names its images through it.
    /// </summary>
    internal const string LegacyRichValueRelRelType = "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel";

    // Content types
    private const string RichValueContentType = "application/vnd.ms-excel.rdrichvalue+xml";
    private const string RichValueStructureContentType = "application/vnd.ms-excel.rdrichvaluestructure+xml";
    private const string RichValueRelContentType = "application/vnd.ms-excel.richvaluerel+xml";
    private const string RichValueTypesContentType = "application/vnd.ms-excel.rdrichvaluetypes+xml";

    // XML namespaces (RvrNs has no OpenXmlConst entry yet)
    private const string RvrNs = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";

    /// <summary>The key of a local image's index into the rich value relationship part.</summary>
    private const string LocalImageIdentifierKey = "_rvRel:LocalImageIdentifier";

    /// <summary>The structure of an error's rich value, [MS-XLSX] 2.3.6.1.3.</summary>
    internal const string ErrorStructureType = "_error";

    /// <summary>
    /// The key of an error's type, numbered as <see cref="XLError"/> is. It is the only key of an error
    /// Excel needs: it opens a file whose <c>_error</c> structure has no other, and shows the error.
    /// </summary>
    internal const string ErrorTypeKey = "errorType";

    /// <summary>
    /// Keys whose flags Excel declares in every rdRichValueTypes part it writes, with the flags
    /// set on each: <c>ExcludeFromFile</c> for <c>_Self</c> only, <c>ExcludeFromCalcComparison</c>
    /// for all.
    /// </summary>
    private static readonly string[] GlobalFlaggedKeys =
    [
        "_Self", "_DisplayString", "_Flags", "_Format", "_SubLabel", "_Attribution", "_Icon", "_Display",
        "_CanonicalPropertyNames", "_ClassificationId",
    ];

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
    /// Write the rich data parts: the rich values of the images, then those of the errors. The
    /// richValueRel part, with the images, is written only when there is an image, as Excel does.
    /// </summary>
    /// <param name="workbookPart">Workbook part to add extended parts to.</param>
    /// <param name="imageStore">Workbook-level image blob store.</param>
    /// <param name="entries">One entry per cell with an in-cell image.</param>
    /// <param name="errors">One entry per error newer than Excel 2007 that a cell holds.</param>
    /// <param name="relIdGenerator">Generator for unique relationship IDs.</param>
    internal static void WriteRichDataParts(
        WorkbookPart workbookPart,
        XLInCellImageStore imageStore,
        IReadOnlyList<RichValueEntry> entries,
        IReadOnlyList<XLError> errors,
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
        if (entries.Count > 0)
        {
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
        }

        // 2. rdrichvalue.xml
        var richValuePart = workbookPart.AddExtendedPart(
            RichValueRelType, RichValueContentType,
            ".xml", relIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
        WriteRichValueXml(richValuePart, entries, errors, imageIndexToRelIndex);

        // 3. rdrichvaluestructure.xml
        var richValueStructurePart = workbookPart.AddExtendedPart(
            RichValueStructureRelType, RichValueStructureContentType,
            ".xml", relIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
        WriteRichValueStructureXml(richValueStructurePart, entries.Count > 0, errors.Count > 0);

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
            LegacyRichValueRelRelType,
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
        using var w = PartXmlWriter.Create(part);

        w.WriteStartDocument(true);
        w.WriteStartElement("richValueRels", RvrNs);
        w.WriteAttributeString("xmlns", "r", null, OpenXmlConst.RelationshipsNs);

        foreach (var relId in imageRelIds)
        {
            w.WriteStartElement("rel", RvrNs);
            w.WriteAttributeString("id", OpenXmlConst.RelationshipsNs, relId);
            w.WriteEndElement(); // rel
        }

        w.WriteEndElement(); // richValueRels
    }

    /// <summary>
    /// rdrichvalue.xml - one rv entry per cell with in-cell image, then one per error. An error is
    /// its <c>errorType</c>, in the <c>_error</c> structure, which follows the image's when there is one.
    /// <code>
    /// &lt;rvData xmlns="..." count="N"&gt;
    ///   &lt;rv s="0"&gt;
    ///     &lt;v&gt;relIndex&lt;/v&gt;
    ///     &lt;v&gt;5&lt;/v&gt;
    ///     &lt;v&gt;altText&lt;/v&gt;
    ///   &lt;/rv&gt;
    ///   &lt;rv s="1"&gt;&lt;v&gt;13&lt;/v&gt;&lt;/rv&gt;
    ///   ...
    /// &lt;/rvData&gt;
    /// </code>
    /// </summary>
    private static void WriteRichValueXml(
        OpenXmlPart part,
        IReadOnlyList<RichValueEntry> entries,
        IReadOnlyList<XLError> errors,
        Dictionary<int, int> imageIndexToRelIndex)
    {
        using var w = PartXmlWriter.Create(part);

        w.WriteStartDocument(true);
        w.WriteStartElement("rvData", OpenXmlConst.RichDataNs);
        w.WriteAttributeString("count", (entries.Count + errors.Count).ToString(CultureInfo.InvariantCulture));

        for (var i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            var relIndex = imageIndexToRelIndex[entry.ImageStoreIndex];

            w.WriteStartElement("rv", OpenXmlConst.RichDataNs);
            w.WriteAttributeString("s", "0"); // structure index 0 = _localImage

            // Key 0: _rvRel:LocalImageIdentifier (rel index)
            w.WriteElementString("v", OpenXmlConst.RichDataNs, relIndex.ToString());
            // Key 1: CalcOrigin
            w.WriteElementString("v", OpenXmlConst.RichDataNs, "5");
            // Key 2: Text (alt text)
            w.WriteElementString("v", OpenXmlConst.RichDataNs, entry.AltText);

            w.WriteEndElement(); // rv
        }

        var errorStructure = (entries.Count > 0 ? 1 : 0).ToString(CultureInfo.InvariantCulture);
        foreach (var error in errors)
        {
            w.WriteStartElement("rv", OpenXmlConst.RichDataNs);
            w.WriteAttributeString("s", errorStructure);
            w.WriteElementString("v", OpenXmlConst.RichDataNs, ((int)error).ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement(); // rv
        }

        w.WriteEndElement(); // rvData
    }

    /// <summary>
    /// rdrichvaluestructure.xml - defines the _localImage structure schema, and the _error one.
    /// Unlike the types part, it is in the <c>richdata</c> namespace, as the rich values are.
    /// <code>
    /// &lt;rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2"&gt;
    ///   &lt;s t="_localImage"&gt;
    ///     &lt;k n="_rvRel:LocalImageIdentifier" t="i"/&gt;&lt;k n="CalcOrigin" t="i"/&gt;&lt;k n="Text" t="s"/&gt;
    ///   &lt;/s&gt;
    ///   &lt;s t="_error"&gt;&lt;k n="errorType" t="i"/&gt;&lt;/s&gt;
    /// &lt;/rvStructures&gt;
    /// </code>
    /// </summary>
    private static void WriteRichValueStructureXml(OpenXmlPart part, bool hasImages, bool hasErrors)
    {
        using var w = PartXmlWriter.Create(part);

        w.WriteStartDocument(true);
        w.WriteStartElement("rvStructures", OpenXmlConst.RichDataNs);
        w.WriteAttributeString("count", ((hasImages ? 1 : 0) + (hasErrors ? 1 : 0)).ToString(CultureInfo.InvariantCulture));

        if (hasImages)
        {
            w.WriteStartElement("s", OpenXmlConst.RichDataNs);
            w.WriteAttributeString("t", "_localImage");

            // Key 0: index of the image in richValueRel.xml
            WriteStructureKey(w, LocalImageIdentifierKey, "i");
            // Key 1: CalcOrigin
            WriteStructureKey(w, "CalcOrigin", "i");
            // Key 2: alt text
            WriteStructureKey(w, "Text", "s");

            w.WriteEndElement(); // s
        }

        if (hasErrors)
        {
            w.WriteStartElement("s", OpenXmlConst.RichDataNs);
            w.WriteAttributeString("t", ErrorStructureType);
            WriteStructureKey(w, ErrorTypeKey, "i");
            w.WriteEndElement(); // s
        }

        w.WriteEndElement(); // rvStructures
    }

    private static void WriteStructureKey(XmlWriter w, string name, string type)
    {
        w.WriteStartElement("k", OpenXmlConst.RichDataNs);
        w.WriteAttributeString("n", name);
        w.WriteAttributeString("t", type);
        w.WriteEndElement(); // k
    }

    /// <summary>
    /// rdRichValueTypes.xml - global key flags (static content), as Excel writes them.
    /// </summary>
    private static void WriteRichValueTypesXml(OpenXmlPart part)
    {
        using var w = PartXmlWriter.Create(part);

        w.WriteStartDocument(true);
        w.WriteStartElement("rvTypesInfo", OpenXmlConst.RichData2Ns);

        w.WriteStartElement("global", OpenXmlConst.RichData2Ns);

        w.WriteStartElement("keyFlags", OpenXmlConst.RichData2Ns);

        foreach (var key in GlobalFlaggedKeys)
        {
            w.WriteStartElement("key", OpenXmlConst.RichData2Ns);
            w.WriteAttributeString("name", key);

            if (key == "_Self")
                WriteKeyFlag(w, "ExcludeFromFile");

            WriteKeyFlag(w, "ExcludeFromCalcComparison");

            w.WriteEndElement(); // key
        }

        w.WriteEndElement(); // keyFlags

        w.WriteEndElement(); // global

        w.WriteEndElement(); // rvTypesInfo
    }

    private static void WriteKeyFlag(XmlWriter w, string name)
    {
        w.WriteStartElement("flag", OpenXmlConst.RichData2Ns);
        w.WriteAttributeString("name", name);
        w.WriteAttributeString("value", "1");
        w.WriteEndElement(); // flag
    }
}
