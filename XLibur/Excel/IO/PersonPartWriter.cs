using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes <c>xl/persons/person.xml</c>, the workbook-level list of identities that threaded
/// comments are attributed to.
/// </summary>
internal static class PersonPartWriter
{
    internal const string ThreadedCommentsNs =
        "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

    /// <summary>
    /// The persons that have to appear in <c>person.xml</c>: the workbook's list, plus any author a
    /// thread still references. Removing a person from the list does not rewrite the comments they
    /// authored, and a <c>personId</c> with no matching <c>&lt;person&gt;</c> is a dangling
    /// reference Excel cannot resolve.
    /// </summary>
    internal static List<XLPerson> CollectReferencedPersons(XLWorkbook workbook)
    {
        var persons = new List<XLPerson>();
        var seen = new HashSet<Guid>();

        foreach (var person in workbook.PersonsInternal)
        {
            if (seen.Add(person.Id))
                persons.Add((XLPerson)person);
        }

        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var cell in worksheet.Internals.CellsCollection.GetCells(c => c.HasThreadedComment))
            {
                var root = cell.SliceThreadedComment!;
                AddIfNew(persons, seen, root.AuthorInternal);

                if (root.RepliesInternal is { } replies)
                {
                    foreach (var reply in replies)
                        AddIfNew(persons, seen, reply.AuthorInternal);
                }
            }
        }

        return persons;
    }

    private static void AddIfNew(List<XLPerson> persons, HashSet<Guid> seen, XLPerson person)
    {
        if (seen.Add(person.Id))
            persons.Add(person);
    }

    internal static void GenerateContent(WorkbookPersonPart personPart, XLWorkbook workbook)
    {
        var settings = new XmlWriterSettings
        {
            CloseOutput = true,
            Encoding = XLHelper.NoBomUTF8
        };

        var partStream = personPart.GetStream(FileMode.Create);
        using var xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartElement("personList", ThreadedCommentsNs);
        xml.WriteAttributeString("xmlns", "x", null, OpenXmlConst.Main2006SsNs);

        foreach (var person in CollectReferencedPersons(workbook))
        {
            xml.WriteStartElement("person", ThreadedCommentsNs);
            xml.WriteAttributeString("displayName", person.DisplayName);
            xml.WriteAttributeString("id", CommentWriteSource.FormatId(person.Id));

            // Excel omits both attributes for a person with no identity provider rather than
            // writing them empty.
            if (person.UserId is not null)
                xml.WriteAttributeString("userId", person.UserId);

            if (person.ProviderId is not null)
                xml.WriteAttributeString("providerId", person.ProviderId);

            xml.WriteEndElement(); // person
        }

        xml.WriteEndElement(); // personList
        xml.Close();
    }
}
