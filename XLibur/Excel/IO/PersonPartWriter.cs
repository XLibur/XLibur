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

        foreach (var person in workbook.PersonsInternal)
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
