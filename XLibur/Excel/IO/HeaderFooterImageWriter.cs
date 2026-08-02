using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class HeaderFooterImageWriter
{
    // VML namespaces fixed by the OOXML spec; the writer emits them verbatim.
    private const string VmlNs = "urn:schemas-microsoft-com:vml";
    private const string OfficeNs = "urn:schemas-microsoft-com:office:office";
    private const string ExcelNs = "urn:schemas-microsoft-com:office:excel";

    /// <summary>
    /// Writes header/footer images as a separate VML drawing part with image relationships,
    /// and adds the <c>&lt;legacyDrawingHF&gt;</c> element to the worksheet XML.
    /// </summary>
    internal static void WriteHeaderFooterImages(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        var header = (XLHeaderFooter)xlWorksheet.PageSetup.Header;
        var footer = (XLHeaderFooter)xlWorksheet.PageSetup.Footer;

        if (!header.HasImages && !footer.HasImages)
        {
            // Remove any existing legacyDrawingHF if no images
            worksheet.RemoveAllChildren<LegacyDrawingHeaderFooter>();
            cm.SetElement(XLWorksheetContents.LegacyDrawingHeaderFooter, null);
            return;
        }

        var headerImages = header.CollectImages("H");
        var footerImages = footer.CollectImages("F");
        var allImages = headerImages.Concat(footerImages).ToList();

        if (allImages.Count == 0)
            return;

        // Create a separate VML drawing part for header/footer images.
        // This is distinct from the comments VML part.
        var hfVmlRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>(hfVmlRelId);

        // Add image parts to the VML part and build relId mapping
        var imageRelIds = new Dictionary<XLHFImage, string>();
        foreach (var image in allImages)
        {
            var imgRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            var imagePart = vmlPart.AddImagePart(image.Format.ToOpenXml(), imgRelId);
            image.ImageStream.Position = 0;
            imagePart.FeedData(image.ImageStream);
            imageRelIds[image] = imgRelId;
        }

        // Write the VML content
        WriteVmlContent(vmlPart, allImages, imageRelIds);

        // Add <legacyDrawingHF r:id="..."/> element to worksheet
        worksheet.RemoveAllChildren<LegacyDrawingHeaderFooter>();
        var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.LegacyDrawingHeaderFooter);
        var legacyDrawingHF = new LegacyDrawingHeaderFooter { Id = hfVmlRelId };
        worksheet.InsertAfter(legacyDrawingHF, previousElement);
        cm.SetElement(XLWorksheetContents.LegacyDrawingHeaderFooter,
            worksheet.Elements<LegacyDrawingHeaderFooter>().First());
    }

    private static void WriteVmlContent(
        VmlDrawingPart vmlPart,
        List<XLHFImage> images,
        Dictionary<XLHFImage, string> imageRelIds)
    {
        using var stream = vmlPart.GetStream(FileMode.Create);
        using var writer = new XmlTextWriter(stream, Encoding.UTF8);

        writer.WriteStartElement("xml");
        writer.WriteAttributeString("xmlns", "v", null, VmlNs);
        writer.WriteAttributeString("xmlns", "o", null, OfficeNs);
        writer.WriteAttributeString("xmlns", "x", null, ExcelNs);

        // Shape layout
        writer.WriteStartElement("o", "shapelayout", OfficeNs);
        writer.WriteAttributeString("v", "ext", VmlNs, "edit");
        writer.WriteStartElement("o", "idmap", OfficeNs);
        writer.WriteAttributeString("v", "ext", VmlNs, "edit");
        writer.WriteAttributeString("data", "2");
        writer.WriteEndElement(); // o:idmap
        writer.WriteEndElement(); // o:shapelayout

        // Shape type for pictures
        writer.WriteStartElement("v", "shapetype", VmlNs);
        writer.WriteAttributeString("id", "_x0000_t75");
        writer.WriteAttributeString("coordsize", "21600,21600");
        writer.WriteAttributeString("o", "spt", OfficeNs, "75");
        writer.WriteAttributeString("o", "preferrelative", OfficeNs, "t");
        writer.WriteAttributeString("path", "m@4@5l@4@11@9@11@9@5xe");
        writer.WriteAttributeString("filled", "f");
        writer.WriteAttributeString("stroked", "f");

        writer.WriteStartElement("v", "stroke", VmlNs);
        writer.WriteAttributeString("joinstyle", "miter");
        writer.WriteEndElement();

        writer.WriteStartElement("v", "formulas", VmlNs);
        WriteFormula(writer, "if lineDrawn pixelLineWidth 0");
        WriteFormula(writer, "sum @0 1 0");
        WriteFormula(writer, "sum 0 0 @1");
        WriteFormula(writer, "prod @2 1 2");
        WriteFormula(writer, "prod @3 21600 pixelWidth");
        WriteFormula(writer, "prod @3 21600 pixelHeight");
        WriteFormula(writer, "sum @0 0 1");
        WriteFormula(writer, "prod @6 1 2");
        WriteFormula(writer, "prod @7 21600 pixelWidth");
        WriteFormula(writer, "sum @8 21600 0");
        WriteFormula(writer, "prod @7 21600 pixelHeight");
        WriteFormula(writer, "sum @10 21600 0");
        writer.WriteEndElement(); // v:formulas

        writer.WriteStartElement("v", "path", VmlNs);
        writer.WriteAttributeString("o", "extrusionok", OfficeNs, "f");
        writer.WriteAttributeString("gradientshapeok", "t");
        writer.WriteAttributeString("o", "connecttype", OfficeNs, "rect");
        writer.WriteEndElement();

        writer.WriteStartElement("o", "lock", OfficeNs);
        writer.WriteAttributeString("v", "ext", VmlNs, "edit");
        writer.WriteAttributeString("aspectratio", "t");
        writer.WriteEndElement();

        writer.WriteEndElement(); // v:shapetype

        // Write a shape for each image
        var shapeIndex = 2049; // Start at _x0000_s2049 to avoid conflicts with comment shapes
        var zIndex = 1;
        foreach (var image in images)
        {
            var relId = imageRelIds[image];
            var spid = $"_x0000_s{shapeIndex++}";
            var widthPt = image.WidthPt.ToInvariantString();
            var heightPt = image.HeightPt.ToInvariantString();
            var style = $"position:absolute;margin-left:0;margin-top:0;width:{widthPt}pt;height:{heightPt}pt;z-index:{zIndex++}";

            writer.WriteStartElement("v", "shape", VmlNs);
            writer.WriteAttributeString("id", image.PositionCode!);
            writer.WriteAttributeString("o", "spid", OfficeNs, spid);
            writer.WriteAttributeString("type", "#_x0000_t75");
            writer.WriteAttributeString("style", style);

            writer.WriteStartElement("v", "imagedata", VmlNs);
            writer.WriteAttributeString("o", "relid", OfficeNs, relId);
            writer.WriteAttributeString("o", "title", OfficeNs, "");
            writer.WriteEndElement();

            writer.WriteStartElement("o", "lock", OfficeNs);
            writer.WriteAttributeString("v", "ext", VmlNs, "edit");
            writer.WriteAttributeString("rotation", "t");
            writer.WriteEndElement();

            writer.WriteEndElement(); // v:shape
        }

        writer.WriteEndElement(); // xml
        writer.Flush();
    }

    private static void WriteFormula(XmlTextWriter writer, string eqn)
    {
        writer.WriteStartElement("v", "f", VmlNs);
        writer.WriteAttributeString("eqn", eqn);
        writer.WriteEndElement();
    }
}
