using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using XLibur.Excel;

namespace XLibur.Tests.Excel;

/// <summary>
/// The style each cell resolves to, by sheet name and cell address. A save renumbers style records
/// and drops the index of a cell whose style is the default, so counting <c>c@s</c> cannot see a cell
/// that lost its style. This compares what the index points to instead.
/// </summary>
internal sealed class RoundTripCensusCellStyles
{
    /// <summary>The parts of a resolved style, each counted under its own key, <c>c@s-&gt;font</c> and so on.</summary>
    internal static readonly string[] Components = ["numFmt", "font", "fill", "border", "alignment", "protection", "quotePrefix", "cellStyle"];

    internal const string KeyPrefix = "c@s->";

    private static readonly HashSet<string> MainNamespaces =
        ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"];

    private static readonly HashSet<string> Toggles = ["b", "i", "strike", "outline", "shadow", "condense", "extend"];

    /// <summary>Attribute values that say the same as no attribute.</summary>
    private static readonly HashSet<string> Defaults =
    [
        "alignment@horizontal=general", "alignment@vertical=bottom", "alignment@indent=false", "alignment@textRotation=false",
        "alignment@wrapText=false", "alignment@shrinkToFit=false", "alignment@justifyLastLine=false", "alignment@readingOrder=false",
        "alignment@relativeIndent=false", "protection@locked=true", "protection@hidden=false",
        "patternFill@patternType=none", "fgColor@tint=false", "bgColor@tint=false", "color@tint=false",
    ];

    private static readonly HashSet<string> EmptyMeansNothing = ["alignment", "protection"];

    private readonly string[][] _xfs;
    private readonly Dictionary<string, Sheet> _sheets;

    private RoundTripCensusCellStyles(string[][] xfs, Dictionary<string, Sheet> sheets)
    {
        _xfs = xfs;
        _sheets = sheets;
    }

    internal static RoundTripCensusCellStyles Read(IReadOnlyDictionary<string, byte[]> parts)
    {
        var sheets = new Dictionary<string, Sheet>(StringComparer.OrdinalIgnoreCase);
        var workbook = Relationships(parts, "").FirstOrDefault(r => r.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target;
        if (workbook is null || Load(parts, workbook) is not { } workbookXml)
            return new RoundTripCensusCellStyles([], sheets);

        var rels = Relationships(parts, workbook).ToList();
        var stylesPart = rels.FirstOrDefault(r => r.Type.EndsWith("/styles", StringComparison.Ordinal)).Target;
        var xfs = stylesPart is not null && Load(parts, stylesPart) is { } styles ? Xfs(styles) : [];
        var targets = rels.DistinctBy(r => r.Id).ToDictionary(r => r.Id, r => r.Target);
        foreach (var sheet in workbookXml.Descendants().Where(e => e.Name.LocalName == "sheet"))
        {
            var id = sheet.Attributes().FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.NamespaceName.EndsWith("relationships", StringComparison.Ordinal))?.Value;
            if ((string?)sheet.Attribute("name") is { } name && id is not null && targets.TryGetValue(id, out var target) && parts.TryGetValue(target, out var bytes))
                sheets[name] = Sheet.Read(bytes);
        }

        return new RoundTripCensusCellStyles(xfs, sheets);
    }

    /// <summary>
    /// For each style component, how many cells of <paramref name="before"/> resolve to the same value in
    /// <paramref name="after"/>, with the first cell that does not.
    /// </summary>
    internal static IEnumerable<CensusLoss> Losses(string contentType, RoundTripCensusCellStyles before, RoundTripCensusCellStyles after)
    {
        var total = 0;
        var kept = new int[Components.Length];
        var examples = new string?[Components.Length];
        foreach (var (name, sheet) in before._sheets)
        {
            after._sheets.TryGetValue(name, out var afterSheet);
            foreach (var (cell, style) in sheet.Cells)
            {
                total++;
                var was = before.Resolve(style);
                var now = afterSheet is null ? null : after.Resolve(afterSheet.Effective(cell));
                for (var i = 0; i < Components.Length; i++)
                {
                    if (now is not null && was[i] == now[i])
                        kept[i]++;
                    else
                        examples[i] ??= $"{name}!{Address(cell)} {was[i]} -> {(now is null ? "(no sheet)" : now[i])}";
                }
            }
        }

        for (var i = 0; i < Components.Length; i++)
        {
            if (kept[i] < total)
                yield return new CensusLoss(contentType, KeyPrefix + Components[i], total, kept[i], examples[i]);
        }
    }

    private string[] Resolve(int xf) =>
        xf >= 0 && xf < _xfs.Length ? _xfs[xf] : Enumerable.Repeat($"(xf {xf} out of range)", Components.Length).ToArray();

    private static string[][] Xfs(XDocument styles)
    {
        var root = styles.Root!;
        IReadOnlyList<XElement> Children(string list, string item) =>
            root.Elements().FirstOrDefault(e => e.Name.LocalName == list)?.Elements().Where(e => e.Name.LocalName == item).ToList() ?? [];

        var numFmts = Children("numFmts", "numFmt").ToDictionary(n => (string?)n.Attribute("numFmtId") ?? "", n => (string?)n.Attribute("formatCode") ?? "");
        var fonts = Children("fonts", "font");
        var fills = Children("fills", "fill");
        var borders = Children("borders", "border");
        var named = Children("cellStyles", "cellStyle").GroupBy(s => (string?)s.Attribute("xfId") ?? "").ToDictionary(g => g.Key, g => (string?)g.First().Attribute("name") ?? "");

        string Pick(IReadOnlyList<XElement> list, XElement xf, string attribute) =>
            int.TryParse((string?)xf.Attribute(attribute) ?? "0", out var i) && i >= 0 && i < list.Count ? Canonical(list[i]) : $"({attribute} out of range)";

        string Child(XElement xf, string name) =>
            xf.Elements().FirstOrDefault(e => e.Name.LocalName == name) is { } child ? Canonical(child) : "";

        return Children("cellXfs", "xf").Select(xf =>
        {
            var numFmtId = (string?)xf.Attribute("numFmtId") ?? "0";
            var format = numFmts.TryGetValue(numFmtId, out var code) && code.Length > 0 ? code : "#" + numFmtId;
            var xfId = (string?)xf.Attribute("xfId") ?? "0";
            var cellStyle = named.TryGetValue(xfId, out var styleName) ? styleName : xfId == "0" ? "Normal" : "";
            return new[]
            {
                format, Pick(fonts, xf, "fontId"), Pick(fills, xf, "fillId"), Pick(borders, xf, "borderId"),
                Child(xf, "alignment"), Child(xf, "protection"), Bool((string?)xf.Attribute("quotePrefix")), cellStyle,
            };
        }).ToArray();
    }

    private static string Canonical(XElement element)
    {
        var name = element.Name.LocalName;
        var attributes = element.Attributes()
            .Where(a => !a.IsNamespaceDeclaration)
            .Select(a => $"{a.Name.LocalName}={Bool(a.Value)}")
            .Where(a => !Defaults.Contains($"{name}@{a}"))
            .ToList();
        if (Toggles.Contains(name))
        {
            if (attributes.Count == 0)
                attributes.Add("val=true");
            else if (attributes.Contains("val=false"))
                return "";
        }

        if (attributes.Count == 0 && !element.HasElements && EmptyMeansNothing.Contains(name))
            return "";

        // Automatic text is black, and the save writes one where the source had the other.
        if (name == "color" && element.Parent?.Name.LocalName == "font" && attributes is ["rgb=FF000000"])
            return "";

        // A solid fill shows only its foreground colour, and the save rewrites the background one.
        var children = element.Elements();
        if (name == "patternFill" && attributes.Contains("patternType=solid"))
            children = children.Where(c => c.Name.LocalName != "bgColor");

        attributes.Sort(StringComparer.Ordinal);
        return $"<{name} {string.Join(' ', attributes)}>{string.Concat(children.Select(Canonical))}";
    }

    private static string Bool(string? value) => value switch { "1" => "true", "0" or null => "false", _ => value };

    private static XDocument? Load(IReadOnlyDictionary<string, byte[]> parts, string name)
    {
        if (!parts.TryGetValue(name, out var bytes))
            return null;

        try
        {
            using var reader = XmlReader.Create(new MemoryStream(bytes), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            return XDocument.Load(reader);
        }
        catch (XmlException)
        {
            return null;
        }
    }

    private static IEnumerable<(string Id, string Type, string Target)> Relationships(IReadOnlyDictionary<string, byte[]> parts, string source)
    {
        var directory = source.Contains('/') ? source[..(source.LastIndexOf('/') + 1)] : "";
        var fileName = source[directory.Length..];
        if (Load(parts, $"{directory}_rels/{fileName}.rels") is not { } rels)
            yield break;

        foreach (var rel in rels.Root!.Elements().Where(e => e.Name.LocalName == "Relationship"))
        {
            if ((string?)rel.Attribute("TargetMode") == "External")
                continue;

            yield return ((string?)rel.Attribute("Id") ?? "", (string?)rel.Attribute("Type") ?? "", Resolve(directory, (string?)rel.Attribute("Target") ?? ""));
        }
    }

    private static string Resolve(string directory, string target)
    {
        var path = target.StartsWith('/') ? target[1..] : directory + target;
        var segments = new List<string>();
        foreach (var segment in path.Split('/'))
        {
            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);
            }
            else if (segment is not ("." or ""))
            {
                segments.Add(segment);
            }
        }

        return string.Join('/', segments);
    }

    private static string Address(long cell) => XLHelper.GetColumnLetterFromNumber((int)(cell % ColumnSpan)) + (cell / ColumnSpan);

    private const long ColumnSpan = 1 << 15;

    /// <summary>
    /// A cell's style index, and the styles a cell with no element of its own takes: the row's when
    /// the row has customFormat, otherwise the column's.
    /// </summary>
    private sealed class Sheet
    {
        internal Dictionary<long, int> Cells { get; } = [];

        private readonly Dictionary<int, int> _rows = [];
        private readonly List<(int Min, int Max, int Style)> _columns = [];

        internal int Effective(long cell)
        {
            if (Cells.TryGetValue(cell, out var style))
                return style;

            if (_rows.TryGetValue((int)(cell / ColumnSpan), out style))
                return style;

            var column = (int)(cell % ColumnSpan);
            foreach (var (min, max, columnStyle) in _columns)
            {
                if (column >= min && column <= max)
                    return columnStyle;
            }

            return 0;
        }

        internal static Sheet Read(byte[] bytes)
        {
            var sheet = new Sheet();
            var row = 0;
            var column = 0;
            try
            {
                using var reader = XmlReader.Create(new MemoryStream(bytes), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element || !MainNamespaces.Contains(reader.NamespaceURI))
                        continue;

                    switch (reader.LocalName)
                    {
                        case "col":
                            sheet._columns.Add((Int(reader, "min"), Int(reader, "max"), Int(reader, "style")));
                            break;
                        case "row":
                            row = reader.GetAttribute("r") is { } r ? int.Parse(r) : row + 1;
                            column = 0;
                            if (Bool(reader.GetAttribute("customFormat")) == "true")
                                sheet._rows[row] = Int(reader, "s");
                            break;
                        case "c":
                            if (reader.GetAttribute("r") is { } reference)
                            {
                                var (cellRow, cellColumn) = Parse(reference);
                                row = cellRow;
                                column = cellColumn;
                            }
                            else
                            {
                                column++;
                            }

                            sheet.Cells[row * ColumnSpan + column] = Int(reader, "s");
                            break;
                    }
                }
            }
            catch (XmlException)
            {
                // A sheet that is not well-formed XML has no cells to compare.
            }

            return sheet;
        }

        private static int Int(XmlReader reader, string attribute) =>
            int.TryParse(reader.GetAttribute(attribute), out var value) ? value : 0;

        private static (int Row, int Column) Parse(string reference)
        {
            var letters = 0;
            while (letters < reference.Length && char.IsLetter(reference[letters]))
                letters++;

            return (int.Parse(reference.AsSpan(letters)), XLHelper.GetColumnNumberFromLetter(reference[..letters]));
        }
    }
}
