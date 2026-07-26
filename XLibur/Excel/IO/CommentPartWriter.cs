using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.RichText;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal static class CommentPartWriter
{
    internal static void GenerateWorksheetCommentsPartContent(WorksheetCommentsPart worksheetCommentsPart,
        XLWorksheet xlWorksheet)
    {
        var settings = new XmlWriterSettings
        {
            CloseOutput = true,
            Encoding = XLHelper.NoBomUTF8
        };
        var partStream = worksheetCommentsPart.GetStream(FileMode.Create);
        using var xml = XmlWriter.Create(partStream, settings);

        var entries = CommentWriteSource.Collect(xlWorksheet);
        var authorsDict = new Dictionary<string, int>();
        var hasThreads = false;
        xml.WriteStartElement("x", "comments", Main2006SsNs);

        foreach (var entry in entries)
        {
            if (!authorsDict.ContainsKey(entry.Comment.Author))
                authorsDict.Add(entry.Comment.Author, authorsDict.Count);

            hasThreads |= entry.Thread is not null;
        }

        // The uid a thread's fallback note carries lives in the revision namespace, which older
        // consumers must be able to skip over.
        if (hasThreads)
        {
            xml.WriteAttributeString("xmlns", "mc", null, MarkupCompatibilityNs);
            xml.WriteAttributeString("mc", "Ignorable", MarkupCompatibilityNs, "xr");
            xml.WriteAttributeString("xmlns", "xr", null, RevisionNs);
        }

        xml.WriteStartElement("authors", Main2006SsNs);
        foreach (var author in authorsDict)
            xml.WriteElementString("author", Main2006SsNs, author.Key);

        xml.WriteEndElement(); // authors

        var refBuffer = new char[10];
        xml.WriteStartElement("commentList", Main2006SsNs);
        foreach (var entry in entries)
        {
            var comment = entry.Comment;
            xml.WriteStartElement("comment", Main2006SsNs);

            var refLen = entry.Cell.SheetPoint.Format(refBuffer);
            xml.WriteStartAttribute("ref");
            xml.WriteRaw(refBuffer, 0, refLen);
            xml.WriteEndAttribute(); // ref

            xml.WriteAttribute("authorId", authorsDict[comment.Author]);

            // Excel specifies @guid is optional if the workbook is not shared.
            // Excel ignores the shapeId attribute, and the strict schema the OpenXML SDK validates
            // against does not declare it at all, so it is left out even though Excel writes it.
            // What actually pairs a fallback note with its thread is the uid below, matching the
            // thread root's id, together with the "tc={rootId}" author.
            if (entry.Thread is { } thread)
                xml.WriteAttributeString("xr", "uid", RevisionNs, CommentWriteSource.FormatId(thread.Id));

            xml.WriteStartElement("text", Main2006SsNs);
            var richText = XLImmutableRichText.Create(comment);
            foreach (var run in richText.Runs)
                TextSerializer.WriteRun(xml, richText, run);

            xml.WriteEndElement(); // text
            xml.WriteEndElement(); // comment
        }

        xml.WriteEndElement(); // commentList
        xml.WriteEndElement(); // comments

        xml.Close();
    }
}
