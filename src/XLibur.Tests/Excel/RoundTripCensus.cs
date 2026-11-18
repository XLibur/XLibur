using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Xml;

namespace XLibur.Tests.Excel;

internal sealed class RoundTripCensus
{
    /// <summary>The content type the census gives <c>[Content_Types].xml</c>, which has none of its own.</summary>
    internal const string ContentTypesPart = "[Content_Types]";

    internal const string FormulaValueKey = "f/v";

    internal const string BinaryKeyPrefix = "binary ";

    /// <summary>
    /// Short, document-independent prefixes for the namespaces a workbook usually holds. A producer may
    /// bind any prefix to a namespace, and the save may bind another, so keys use these instead.
    /// </summary>
    private static readonly Dictionary<string, string> Prefixes = new()
    {
        ["http://schemas.openxmlformats.org/spreadsheetml/2006/main"] = "",
        ["http://purl.oclc.org/ooxml/spreadsheetml/main"] = "",
        ["http://schemas.openxmlformats.org/package/2006/relationships"] = "",
        ["http://schemas.openxmlformats.org/package/2006/content-types"] = "",
        ["http://schemas.openxmlformats.org/officeDocument/2006/relationships"] = "r",
        ["http://purl.oclc.org/ooxml/officeDocument/relationships"] = "r",
        ["http://schemas.openxmlformats.org/markup-compatibility/2006"] = "mc",
        ["http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"] = "x14ac",
        ["http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"] = "x14",
        ["http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"] = "x15",
        ["http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"] = "x15ac",
        ["http://schemas.microsoft.com/office/spreadsheetml/2014/revision"] = "xr",
        ["http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"] = "xr2",
        ["http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"] = "xr3",
        ["http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"] = "xr6",
        ["http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"] = "xr10",
        ["http://schemas.microsoft.com/office/spreadsheetml/2017/revision16"] = "xr16",
        ["http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"] = "x16r2",
        ["http://schemas.microsoft.com/office/excel/2006/main"] = "xm",
        ["http://schemas.openxmlformats.org/drawingml/2006/main"] = "a",
        ["http://purl.oclc.org/ooxml/drawingml/main"] = "a",
        ["http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"] = "xdr",
        ["http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"] = "xdr",
        ["http://schemas.openxmlformats.org/drawingml/2006/chart"] = "c",
        ["http://purl.oclc.org/ooxml/drawingml/chart"] = "c",
        ["urn:schemas-microsoft-com:vml"] = "v",
        ["urn:schemas-microsoft-com:office:office"] = "o",
        ["urn:schemas-microsoft-com:office:excel"] = "x",
        ["http://schemas.openxmlformats.org/package/2006/metadata/core-properties"] = "cp",
        ["http://purl.org/dc/elements/1.1/"] = "dc",
        ["http://purl.org/dc/terms/"] = "dcterms",
        ["http://www.w3.org/2001/XMLSchema-instance"] = "xsi",
        ["http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"] = "ep",
        ["http://purl.oclc.org/ooxml/officeDocument/extendedProperties"] = "ep",
        ["http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"] = "op",
        ["http://purl.oclc.org/ooxml/officeDocument/customProperties"] = "op",
        ["http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"] = "vt",
        ["http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"] = "vt",
    };

    private static readonly string[] ContentTypePrefixes =
        ["application/vnd.openxmlformats-officedocument.", "application/vnd.openxmlformats-", "application/vnd.", "application/"];

    private const string Worksheet = "spreadsheetml.worksheet";

    private const string Styles = "spreadsheetml.styles";

    private readonly Dictionary<(string ContentType, string Key), int> _counts;
    private readonly RoundTripCensusCellStyles _cellStyles;
    private readonly RoundTripCensusDxfs _dxfs;

    private RoundTripCensus(Dictionary<(string ContentType, string Key), int> counts, RoundTripCensusCellStyles cellStyles, RoundTripCensusDxfs dxfs)
    {
        _counts = counts;
        _cellStyles = cellStyles;
        _dxfs = dxfs;
    }

    internal static RoundTripCensus Of(Stream package, IReadOnlySet<(string ContentType, string Key)> qualified)
    {
        var counts = new Dictionary<(string, string), int>();
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var contentTypes = ContentTypeMap.Read(zip);
        var parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var dxfReferences = new List<(string, string, string)>();
        byte[]? styles = null;
        foreach (var entry in zip.Entries)
        {
            if (entry.FullName.EndsWith('/'))
                continue;

            var contentType = Shorten(contentTypes.Of(entry.FullName));
            var bytes = ReadAll(entry);
            parts[entry.FullName] = bytes;
            if (contentType == Styles)
                styles = bytes;

            if (!TryCountXml(bytes, contentType, qualified, counts, dxfReferences))
                Add(counts, contentType, BinaryKeyPrefix + Convert.ToHexString(SHA256.HashData(bytes))[..16]);
        }

        return new RoundTripCensus(counts, RoundTripCensusCellStyles.Read(parts), RoundTripCensusDxfs.Resolve(dxfReferences, styles));
    }

    internal static IReadOnlyList<CensusLoss> Losses(RoundTripCensus before, RoundTripCensus after) =>
        before._counts
            .Where(kv => after._counts.GetValueOrDefault(kv.Key) < kv.Value)
            .Select(kv => new CensusLoss(kv.Key.ContentType, kv.Key.Key, kv.Value, after._counts.GetValueOrDefault(kv.Key)))
            .Concat(RoundTripCensusCellStyles.Losses(Worksheet, before._cellStyles, after._cellStyles))
            .Concat(RoundTripCensusDxfs.Losses(before._dxfs, after._dxfs))
            .OrderBy(l => l.ContentType, StringComparer.Ordinal)
            .ThenBy(l => l.Key, StringComparer.Ordinal)
            .ToList();

    internal static string Shorten(string contentType)
    {
        var s = contentType;
        foreach (var prefix in ContentTypePrefixes)
        {
            if (s.StartsWith(prefix, StringComparison.Ordinal))
            {
                s = s[prefix.Length..];
                break;
            }
        }

        return s.EndsWith("+xml", StringComparison.Ordinal) ? s[..^4] : s;
    }

    private static byte[] ReadAll(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static bool TryCountXml(byte[] bytes, string contentType, IReadOnlySet<(string, string)> qualified, Dictionary<(string, string), int> counts,
        List<(string, string, string)> dxfReferences)
    {
        var local = new Dictionary<(string, string), int>();
        var localReferences = new List<(string, string, string)>();
        var attributes = new List<(string Name, string Value)>();
        var previousSibling = new List<string?>();
        try
        {
            using var reader = XmlReader.Create(new MemoryStream(bytes), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                var element = Name(reader.NamespaceURI, reader.LocalName);
                var depth = reader.Depth;
                while (previousSibling.Count <= depth + 1)
                    previousSibling.Add(null);

                var elementKey = element switch
                {
                    "v" when previousSibling[depth] == "f" => FormulaValueKey,
                    "is" => "v",
                    _ => element,
                };
                previousSibling[depth] = element;
                previousSibling[depth + 1] = null;

                attributes.Clear();
                if (reader.MoveToFirstAttribute())
                {
                    do
                    {
                        if (reader.NamespaceURI != "http://www.w3.org/2000/xmlns/")
                            attributes.Add((Name(reader.NamespaceURI, reader.LocalName), reader.Value));
                    } while (reader.MoveToNextAttribute());
                }

                string? qualifiedBy = null;
                foreach (var (name, value) in attributes)
                {
                    if (Qualified(qualified, contentType, $"{element}[@{name}=", value, "]") is { } key)
                    {
                        elementKey = key;
                        qualifiedBy = name;
                        break;
                    }
                }

                Add(local, contentType, elementKey);
                foreach (var (name, value) in attributes)
                {
                    if (RoundTripCensusDxfs.IsReference(name))
                        localReferences.Add((contentType, element + "@" + name, value));

                    if (name == qualifiedBy)
                        continue;

                    var attribute = elementKey + "@" + name;
                    Add(local, contentType, Qualified(qualified, contentType, attribute + "=", value, "") ?? attribute);
                }
            }
        }
        catch (XmlException)
        {
            return false;
        }

        foreach (var (key, count) in local)
            counts[key] = counts.GetValueOrDefault(key) + count;

        dxfReferences.AddRange(localReferences);
        return true;
    }

    /// <summary>
    /// The key <paramref name="prefix"/><paramref name="value"/><paramref name="suffix"/> if
    /// <paramref name="qualified"/> lists it, also when it spells a value of <c>0</c> or <c>1</c> as
    /// <c>false</c> or <c>true</c>; otherwise null.
    /// </summary>
    private static string? Qualified(IReadOnlySet<(string, string)> qualified, string contentType, string prefix, string value, string suffix)
    {
        var key = prefix + value + suffix;
        if (qualified.Contains((contentType, key)))
            return key;

        var spelled = value switch { "0" => "false", "1" => "true", _ => null };
        if (spelled is null)
            return null;

        key = prefix + spelled + suffix;
        return qualified.Contains((contentType, key)) ? key : null;
    }

    private static string Name(string ns, string localName)
    {
        if (ns.Length == 0)
            return localName;

        if (!Prefixes.TryGetValue(ns, out var prefix))
            return "{" + ns + "}" + localName;

        return prefix.Length == 0 ? localName : prefix + ":" + localName;
    }

    private static void Add(Dictionary<(string, string), int> counts, string contentType, string key)
    {
        var k = (contentType, key);
        counts[k] = counts.GetValueOrDefault(k) + 1;
    }

    private sealed class ContentTypeMap
    {
        private readonly Dictionary<string, string> _defaults = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _overrides = new(StringComparer.OrdinalIgnoreCase);

        internal static ContentTypeMap Read(ZipArchive zip)
        {
            var map = new ContentTypeMap();
            var entry = zip.GetEntry("[Content_Types].xml");
            if (entry is null)
                return map;

            using var reader = XmlReader.Create(entry.Open(), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                var type = reader.GetAttribute("ContentType");
                if (type is null)
                    continue;

                if (reader.LocalName == "Default" && reader.GetAttribute("Extension") is { } extension)
                    map._defaults[extension] = type;
                else if (reader.LocalName == "Override" && reader.GetAttribute("PartName") is { } partName)
                    map._overrides[partName.TrimStart('/')] = type;
            }

            return map;
        }

        internal string Of(string entryName)
        {
            if (entryName == "[Content_Types].xml")
                return ContentTypesPart;

            if (_overrides.TryGetValue(entryName, out var type))
                return type;

            var extension = Path.GetExtension(entryName).TrimStart('.');
            return _defaults.TryGetValue(extension, out type) ? type : "(none)/" + extension;
        }
    }
}

internal sealed record CensusLoss(string ContentType, string Key, int Before, int After, string? Example = null)
{
    public override string ToString() => $"{ContentType} {Key}: {Before} -> {After}" + (Example is null ? "" : $" (first: {Example})");
}
