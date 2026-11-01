using System;
using System.IO;
using System.Text.RegularExpressions;

namespace XLibur.Tests.Utils;

/// <summary>
/// Hand-edits the stylesheet of a saved package, <c>xl/styles.xml</c>, whose elements XLibur writes
/// with the <c>x:</c> prefix: tests put fonts, style formats, cell formats and cell styles in it
/// that XLibur would never write itself.
/// </summary>
internal static class StylesPart
{
    private const string PartName = "xl/styles.xml";

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);

    /// <summary>
    /// Replaces each named element of the stylesheet in <paramref name="package"/>, with all it holds,
    /// by the XML given for it, e.g. <c>("cellXfs", "&lt;x:cellXfs count=\"1\"&gt;...")</c>. Returns
    /// the same stream, rewound, for chaining.
    /// </summary>
    internal static MemoryStream RewriteStyles(this MemoryStream package, params (string Element, string Xml)[] replacements) =>
        package.RewritePart(PartName, xml => ReplaceElements(xml, replacements));

    /// <summary>
    /// Replaces each named element of <paramref name="xml"/>, with all it holds, by the XML given for
    /// it. An element is named in full, e.g. <c>fonts</c>, never by a prefix of another's name.
    /// </summary>
    internal static string ReplaceElements(string xml, params (string Element, string Xml)[] replacements)
    {
        foreach (var (element, replacement) in replacements)
            xml = Regex.Replace(xml, $"<x:{element}.*?</x:{element}>", replacement, RegexOptions.None, RegexTimeout);

        return xml;
    }

    /// <summary>The first element named <paramref name="element"/> in the stylesheet a save wrote, e.g. the first <c>font</c>.</summary>
    internal static string FirstElement(Stream saved, string element) =>
        Regex.Match(saved.ReadPart(PartName), $"<x:{element}>.*?</x:{element}>", RegexOptions.None, RegexTimeout).Value;
}
