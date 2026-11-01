using System.IO;
using System.Xml;
using System.Xml.Linq;
using XLibur.Excel.IO;

namespace XLibur.Extensions;

internal static class XDocumentExtensions
{
    public static XDocument? Load(Stream stream)
    {
        using var reader = PartXmlReader.CreateVerbatim(stream);
        try
        {
            return XDocument.Load(reader);
        }
        catch (XmlException)
        {
            return null;
        }
    }
}
