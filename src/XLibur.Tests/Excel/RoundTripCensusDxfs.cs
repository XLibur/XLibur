using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace XLibur.Tests.Excel;

/// <summary>
/// The differential format each <c>dxfId</c> reference resolves to: a conditional format rule, a table
/// or table column, a pivot format, a table or slicer style element. A save prunes, merges and
/// renumbers the dxfs, so counting the styles part cannot see a rule whose format lost its fill. This
/// counts, per reference and dxf component, how many references resolve to each canonical value, and
/// a value that has fewer references after the save is lost.
/// </summary>
internal sealed class RoundTripCensusDxfs
{
    internal const string Arrow = "->";

    private static readonly HashSet<string> Toggles = ["b", "i", "strike", "outline", "shadow", "condense", "extend"];

    /// <summary>Attribute values that say the same as no attribute.</summary>
    private static readonly HashSet<string> Defaults = ["fgColor@tint=false", "bgColor@tint=false", "color@tint=false"];

    private readonly Dictionary<(string ContentType, string Key), Dictionary<string, int>> _resolved;

    private RoundTripCensusDxfs(Dictionary<(string ContentType, string Key), Dictionary<string, int>> resolved) => _resolved = resolved;

    internal static bool IsReference(string attribute) =>
        attribute == "dxfId" || attribute.EndsWith("DxfId", StringComparison.Ordinal);

    /// <param name="references">Each reference as its part's content type, <c>element@attribute</c> and the dxf index.</param>
    /// <param name="styles">The styles part, or null when the package has none.</param>
    internal static RoundTripCensusDxfs Resolve(IEnumerable<(string ContentType, string Reference, string DxfId)> references, byte[]? styles)
    {
        var dxfs = styles is null ? [] : Read(styles);
        var resolved = new Dictionary<(string, string), Dictionary<string, int>>();
        foreach (var (contentType, reference, dxfId) in references)
        {
            var components = int.TryParse(dxfId, out var i) && i >= 0 && i < dxfs.Count ? dxfs[i] : [("dxf", $"(dxf {dxfId} out of range)")];
            foreach (var (component, value) in components)
            {
                var key = (contentType, reference + Arrow + component);
                if (!resolved.TryGetValue(key, out var values))
                    resolved[key] = values = new Dictionary<string, int>(StringComparer.Ordinal);

                values[value] = values.GetValueOrDefault(value) + 1;
            }
        }

        return new RoundTripCensusDxfs(resolved);
    }

    /// <summary>
    /// For each reference key, how many of the references in <paramref name="before"/> still resolve to
    /// a value <paramref name="after"/> has as often, with the first value that lost a reference.
    /// </summary>
    internal static IEnumerable<CensusLoss> Losses(RoundTripCensusDxfs before, RoundTripCensusDxfs after)
    {
        foreach (var (key, values) in before._resolved)
        {
            after._resolved.TryGetValue(key, out var afterValues);
            var total = 0;
            var lost = 0;
            string? example = null;
            foreach (var (value, count) in values)
            {
                total += count;
                var missing = count - (afterValues?.GetValueOrDefault(value) ?? 0);
                if (missing > 0)
                {
                    lost += missing;
                    example ??= value;
                }
            }

            if (lost > 0)
            {
                var gained = afterValues?.FirstOrDefault(v => v.Value > values.GetValueOrDefault(v.Key)).Key;
                yield return new CensusLoss(key.ContentType, key.Key, total, total - lost, $"{example} -> {gained ?? "(none)"}");
            }
        }
    }

    private static List<(string Component, string Value)[]> Read(byte[] styles)
    {
        XDocument document;
        try
        {
            using var reader = XmlReader.Create(new MemoryStream(styles), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            document = XDocument.Load(reader);
        }
        catch (XmlException)
        {
            return [];
        }

        var dxfs = document.Root?.Elements().FirstOrDefault(e => e.Name.LocalName == "dxfs")?.Elements().Where(e => e.Name.LocalName == "dxf") ?? [];
        return dxfs.Select(dxf => dxf.Elements()
            .Select(c => (Component: c.Name.LocalName, Value: Canonical(c)))
            .Where(c => c.Value != Empty(c.Component))
            .ToArray()).ToList();
    }

    /// <summary>
    /// An element as text, with its attributes and children sorted. Unlike a cell format, a dxf states
    /// only what it overrides: a toggle with val="0" turns the property off, and a pattern fill with no
    /// patternType is solid, so neither is folded into its absence.
    /// </summary>
    private static string Canonical(XElement element)
    {
        var name = element.Name.LocalName;
        var attributes = element.Attributes()
            .Where(a => !a.IsNamespaceDeclaration)
            .Select(a => $"{a.Name.LocalName}={Value(a)}")
            .Where(a => !Defaults.Contains($"{name}@{a}"))
            .ToList();
        if (Toggles.Contains(name) && !attributes.Any(a => a.StartsWith("val=", StringComparison.Ordinal)))
            attributes.Add("val=true");

        // A custom format is renumbered when the formats are pruned; its code says what it is.
        if (name == "numFmt" && attributes.Any(a => a.StartsWith("formatCode=", StringComparison.Ordinal)))
            attributes.RemoveAll(a => a.StartsWith("numFmtId=", StringComparison.Ordinal));

        var children = element.Elements();
        if (name == "border")
            children = children.Where(side => side.HasAttributes || side.HasElements);

        if (name == "patternFill" && element.HasElements)
        {
            if (!attributes.Any(a => a.StartsWith("patternType=", StringComparison.Ordinal)))
                attributes.Add("patternType=solid");

            // A solid dxf fill shows its background colour, and the save adds an automatic foreground one.
            if (attributes.Contains("patternType=solid"))
                children = children.Where(c => !(c.Name.LocalName == "fgColor" && Bool((string?)c.Attribute("auto")) == "true"));
        }

        attributes.Sort(StringComparer.Ordinal);
        return $"<{name} {string.Join(' ', attributes)}>{string.Concat(children.Select(Canonical).Order(StringComparer.Ordinal))}";
    }

    /// <summary>The text of a component that overrides nothing.</summary>
    private static string Empty(string name) => $"<{name} >" + (name == "fill" ? "<patternFill >" : "");

    private static string Value(XAttribute attribute) =>
        attribute.Name.LocalName == "tint" && double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var tint)
            ? Bool(tint.ToString("R", CultureInfo.InvariantCulture))
            : Bool(attribute.Value);

    private static string Bool(string? value) => value switch { "1" => "true", "0" or null => "false", _ => value };
}
