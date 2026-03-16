using XLibur.Excel.ContentManagers;
using XLibur.Excel.Exceptions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using XLibur.Excel.Tables;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class WorksheetPartWriter
{
    internal static void GenerateWorksheetPartContent(
        bool partIsEmpty,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveOptions options,
        SaveContext context)
    {
        var worksheetDom = GetWorksheetDom(partIsEmpty, worksheetPart, xlWorksheet, options, context);
        StreamToPart(worksheetDom, worksheetPart, xlWorksheet, context, options);
    }

    private static Worksheet GetWorksheetDom(
        bool partIsEmpty,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveOptions options,
        SaveContext context)
    {
        if (options.ConsolidateConditionalFormatRanges)
        {
            xlWorksheet.ConditionalFormats.Consolidate();
        }

        #region Worksheet

        Worksheet worksheet;
        if (!partIsEmpty)
        {
            // Accessing the worksheet through worksheetPart.Worksheet creates an attached DOM
            // worksheet that is tracked and later saved automatically to the part.
            // Using the reader, we get a detached DOM.
            // The OpenXmlReader.Create method only reads xml declaration, but doesn't read content.
            using var reader = OpenXmlReader.Create(worksheetPart);
            if (!reader.Read())
            {
                throw new ArgumentException("Worksheet part should contain worksheet xml, but is empty.");
            }

            worksheet = (Worksheet)reader.LoadCurrentElement()!;
        }
        else
        {
            worksheet = new Worksheet();
        }

        if (worksheet.NamespaceDeclarations.All(ns => ns.Value != RelationshipsNs))
            worksheet.AddNamespaceDeclaration("r", RelationshipsNs);

        // We store the x14ac:dyDescent attribute (if set by a xlRow) in a row element. It's an optional attribute and it
        // needs a declared namespace. To avoid writing namespace to each <x:row> element during streaming, write it to
        // every sheet part ahead of time. The namespace has to be marked as ignorable, because OpenXML SDK validator will
        // refuse to validate it because it's an optional extension (see ISO29500 part 3).
        if (worksheet.NamespaceDeclarations.All(ns => ns.Value != X14Ac2009SsNs))
        {
            worksheet.AddNamespaceDeclaration("x14ac", X14Ac2009SsNs);
            worksheet.SetAttribute(new OpenXmlAttribute("mc", "Ignorable", MarkupCompatibilityNs, "x14ac"));
        }

        #endregion Worksheet

        var cm = new XLWorksheetContentManager(worksheet);

        SheetViewWriter.WriteSheetProperties(worksheet, cm, xlWorksheet);
        SheetViewWriter.WriteSheetDimension(worksheet, cm, xlWorksheet);
        SheetViewWriter.WriteSheetViews(worksheet, cm, xlWorksheet);

        var maxOutlineColumn = 0;
        if (xlWorksheet.ColumnCount() > 0)
            maxOutlineColumn = xlWorksheet.GetMaxColumnOutline();

        var maxOutlineRow = 0;
        if (xlWorksheet.RowCount() > 0)
            maxOutlineRow = xlWorksheet.GetMaxRowOutline();

        SheetViewWriter.WriteSheetFormatProperties(worksheet, cm, xlWorksheet,
            maxOutlineColumn, maxOutlineRow, out var worksheetColumnWidth);

        ColumnWriter.WriteColumns(worksheet, cm, xlWorksheet, worksheetColumnWidth, context);

        #region SheetData

        if (!worksheet.Elements<SheetData>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.SheetData);
            worksheet.InsertAfter(new SheetData(), previousElement);
        }

        var sheetData = worksheet.Elements<SheetData>().First();
        cm.SetElement(XLWorksheetContents.SheetData, sheetData);

        // Sheet data is not updated in the Worksheet DOM here, because it is later being streamed directly to the file
        // without an intermediate DOM representation. This is done to save memory, which is especially problematic
        // for large sheets.

        #endregion SheetData

        SheetProtectionWriter.WriteSheetProtection(worksheet, cm, xlWorksheet);
        AutoFilterWriter.WriteAutoFilter(worksheet, cm, xlWorksheet, context);

        WriteMergeCells(worksheet, cm, xlWorksheet);

        ConditionalFormattingWriter.WriteConditionalFormatting(worksheet, cm, xlWorksheet, context);
        ConditionalFormattingWriter.WriteSparklines(worksheet, cm, xlWorksheet);
        DataValidationWriter.WriteDataValidations(worksheet, cm, xlWorksheet, options);
        PageSetupWriter.WriteHyperlinks(worksheet, cm, xlWorksheet, worksheetPart, context);
        PageSetupWriter.WritePrintOptions(worksheet, cm, xlWorksheet);
        PageSetupWriter.WritePageMargins(worksheet, cm, xlWorksheet);
        PageSetupWriter.WritePageSetup(worksheet, cm, xlWorksheet);
        PageSetupWriter.WriteHeaderFooter(worksheet, cm, xlWorksheet);
        PageSetupWriter.WriteRowBreaks(worksheet, cm, xlWorksheet);
        PageSetupWriter.WriteColumnBreaks(worksheet, cm, xlWorksheet);

        PopulateTablePartReferences(xlWorksheet.Tables, worksheet, cm);

        PictureWriter.WriteDrawings(worksheet, cm, xlWorksheet, worksheetPart, context);
        PictureWriter.WriteLegacyDrawing(worksheet, cm, xlWorksheet);
        HeaderFooterImageWriter.WriteHeaderFooterImages(worksheet, cm, xlWorksheet, worksheetPart, context);

        return worksheet;
    }

    private static void WriteMergeCells(Worksheet worksheet, XLWorksheetContentManager cm, XLWorksheet xlWorksheet)
    {
        if ((xlWorksheet).Internals.MergedRanges.Count > 0)
        {
            if (!worksheet.Elements<MergeCells>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.MergeCells);
                worksheet.InsertAfter(new MergeCells(), previousElement);
            }

            var mergeCells = worksheet.Elements<MergeCells>().First();
            cm.SetElement(XLWorksheetContents.MergeCells, mergeCells);
            mergeCells.RemoveAllChildren<MergeCell>();

            foreach (var mergeCell in (xlWorksheet).Internals.MergedRanges
                     .Select(m => m.RangeAddress.FirstAddress + ":" + m.RangeAddress.LastAddress)
                     .Select(merged => new MergeCell { Reference = merged }))
                mergeCells.AppendChild(mergeCell);

            mergeCells.Count = (uint)mergeCells.ChildElements.Count;
        }
        else
        {
            worksheet.RemoveAllChildren<MergeCells>();
            cm.SetElement(XLWorksheetContents.MergeCells, null);
        }
    }

    private static void PopulateTablePartReferences(XLTables xlTables, Worksheet worksheet,
        XLWorksheetContentManager cm)
    {
        var emptyTable = xlTables.FirstOrDefault<XLTable>(t => t.DataRange is null);
        if (emptyTable != null)
            throw new EmptyTableException($"Table '{emptyTable.Name}' should have at least 1 row.");

        TableParts tableParts;
        if (worksheet.Elements<TableParts>().Any())
        {
            tableParts = worksheet.Elements<TableParts>().First();
        }
        else
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.TableParts);
            tableParts = new TableParts();
            worksheet.InsertAfter(tableParts, previousElement);
        }

        cm.SetElement(XLWorksheetContents.TableParts, tableParts);

        xlTables.Deleted.Clear();
        tableParts.RemoveAllChildren();
        foreach (var xlTable in xlTables.Cast<XLTable>())
        {
            tableParts.AppendChild(new TablePart { Id = xlTable.RelId });
        }

        tableParts.Count = (uint)xlTables.Count;
    }

    /// <summary>
    /// Stream detached worksheet DOM to the worksheet part of the stream.
    /// Replaces the content of the part.
    /// </summary>
    private static void StreamToPart(Worksheet worksheet, WorksheetPart worksheetPart, XLWorksheet xlWorksheet,
        SaveContext context, SaveOptions options)
    {
        // Worksheet part might have some data, but the writer truncates everything upon creation.
        using var writer = OpenXmlWriter.Create(worksheetPart);
        using var reader = OpenXmlReader.Create(worksheet);

        writer.WriteStartDocument(true);

        while (reader.Read())
        {
            if (reader.ElementType == typeof(SheetData))
            {
                SheetDataWriter.StreamSheetData(writer, xlWorksheet, context, options);

                // Skip whole SheetData elements from original file, already written
                reader.Skip();
            }

            if (reader.IsStartElement)
            {
                writer.WriteStartElement(reader);
                var canContainText = typeof(OpenXmlLeafTextElement).IsAssignableFrom(reader.ElementType);
                if (canContainText)
                {
                    var text = reader.GetText();
                    if (text.Length > 0)
                    {
                        writer.WriteString(text);
                    }
                }
            }
            else if (reader.IsEndElement)
            {
                writer.WriteEndElement();
            }
        }

        writer.Close();
    }
}
