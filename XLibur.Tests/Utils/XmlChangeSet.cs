using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace XLibur.Tests.Utils;

/// <summary>What happened to one node between the two documents.</summary>
internal enum XmlChangeKind
{
    /// <summary>The node exists only in the later document.</summary>
    Added,

    /// <summary>The node exists only in the earlier document.</summary>
    Removed,

    /// <summary>The node exists in both, but its attributes or its own text differ.</summary>
    Modified,

    /// <summary>The node exists in both, and the children it kept changed their relative order.</summary>
    Reordered,
}

/// <summary>One entry of a change set: a kind, the node it happened to, and what differs.</summary>
internal sealed record XmlChange(XmlChangeKind Kind, string Path, string Detail)
{
    public override string ToString()
    {
        var kind = Kind switch
        {
            XmlChangeKind.Added => "added",
            XmlChangeKind.Removed => "removed",
            XmlChangeKind.Modified => "modified",
            XmlChangeKind.Reordered => "reordered",
            _ => throw new ArgumentOutOfRangeException(nameof(Kind), Kind, "Unknown change kind."),
        };

        return Detail.Length == 0 ? $"{kind} {Path}" : $"{kind} {Path} {Detail}";
    }
}

/// <summary>
/// The difference between one XML part before an operation and the same part after it, expressed as
/// the exact set of nodes that changed. A test asserts that set is precisely what the operation
/// promised — which is stronger than asserting the new value is present, because it also states that
/// nothing else was added, dropped or reordered.
/// </summary>
/// <remarks>
/// <para>
/// Comparison is by meaning, not by bytes. Namespace prefixes, attribute order, the XML declaration,
/// comments, and insignificant whitespace are all absorbed, so a part whose DOM was merely
/// materialised and re-serialized with no model edits produces an empty change set. That is what
/// makes the instrument usable against a save path: loading a part's DOM re-serializes it (see the
/// comment in <c>PictureWriter.RemoveEmptyDrawingPart</c>), so byte comparison reports noise the
/// moment anything is edited.
/// </para>
/// <para>
/// A node is identified by its position: <c>/c:chartSpace[1]/c:chart[1]/c:plotArea[1]</c>, indexed
/// among siblings of the same name and always written with an index, since the path is a diff key
/// rather than a display string. Reordering siblings of <em>different</em> names, which is what the
/// DrawingML sequence rules are about, is reported as <see cref="XmlChangeKind.Reordered"/> on the
/// parent.
/// </para>
/// <para>
/// Positional identity has one cost, and it is worth stating precisely. Anything that changes how
/// many same-named siblings precede a node renumbers that node: swapping two of them reads as
/// modifications to both rather than as a move, and inserting or removing one ahead of the others
/// reads as each of the rest taking on its predecessor's content, plus the addition or removal at
/// the end of the run. A sibling added or removed <em>after</em> the others costs nothing, which is
/// the case that actually arises when a drawing is appended to a part.
/// </para>
/// <para>
/// What that cost is <em>not</em> is a hole. It can make a change set louder than the edit was; it
/// cannot make one empty. An empty change set means every path in one document is present in the
/// other carrying the same attributes, the same text and the same child order — which is what it
/// means for the two to be the same document. An instrument that over-reports costs a reader some
/// time working out which entries were the edit; one that under-reports would let a refactor claim
/// it changed nothing when it had, and that is the failure this exists to prevent.
/// </para>
/// <para>
/// Aligning same-named runs before numbering them would trade that noise away, at the price of a
/// real sequence diff. It has deliberately not been done: the alternative of keying identity on a
/// child's value — a series' <c>c:idx</c>, say — would put one schema's knowledge into an instrument
/// that also has to diff drawings and shapes, and would key identity on a value an edit is allowed
/// to change. <c>XmlChangeSetTests</c> pins all four behaviours described above.
/// </para>
/// <para>
/// An added or removed subtree is reported once, at its root, rather than once per descendant. The
/// entry is suffixed <c>(subtree)</c> when the node has children, so an expectation cannot silently
/// mistake a whole grafted branch for a single empty element.
/// </para>
/// </remarks>
internal sealed class XmlChangeSet
{
    private XmlChangeSet(IReadOnlyList<XmlChange> changes) => Changes = changes;

    /// <summary>
    /// The changes, in the document order of the later document, with removals appended in the
    /// document order of the earlier one.
    /// </summary>
    internal IReadOnlyList<XmlChange> Changes { get; }

    /// <summary>Whether the two documents mean the same thing.</summary>
    internal bool IsEmpty => Changes.Count == 0;

    /// <summary>
    /// The change set as one line per change, which is the shape to assert against: comparing an
    /// ordered list of strings is what makes a failure legible.
    /// </summary>
    internal IReadOnlyList<string> Describe() => Changes.Select(change => change.ToString()).ToList();

    /// <summary>
    /// The change set as a block of text, one change per line. This is the form to assert on, paired
    /// with <see cref="Expect"/>.
    /// </summary>
    public override string ToString() => Expect([.. Describe()]);

    /// <summary>
    /// The block of text a change set would produce if it held exactly these changes, and nothing
    /// else. Pass no arguments to expect no change at all.
    /// </summary>
    /// <remarks>
    /// Asserting on text rather than on a collection is deliberate. A collection assertion that
    /// fails on the count reports only that the counts differ, which is the least useful thing it
    /// could say about a stray mutation; comparing the blocks names the line that was not expected.
    /// That failure message is the whole point of the instrument.
    /// </remarks>
    internal static string Expect(params string[] changes) =>
        changes.Length == 0 ? "(no changes)" : string.Join(Environment.NewLine, changes);

    /// <summary>Computes the change set between two XML documents.</summary>
    internal static XmlChangeSet Between(string beforeXml, string afterXml)
    {
        var before = Flatten(beforeXml);
        var after = Flatten(afterXml);
        var beforeByPath = before.ToDictionary(node => node.Path, StringComparer.Ordinal);
        var afterByPath = after.ToDictionary(node => node.Path, StringComparer.Ordinal);

        var changes = new List<XmlChange>();

        var grafted = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in after)
        {
            if (!beforeByPath.TryGetValue(node.Path, out var earlier))
            {
                // Every path in the grafted region joins the set, whether or not it is reported, so
                // that the descendants of a reported root recognise themselves as already covered.
                // Document order guarantees a parent is seen before its children.
                var alreadyCovered = HasReportedAncestor(node.Path, grafted);
                grafted.Add(node.Path);

                if (!alreadyCovered)
                    changes.Add(new XmlChange(XmlChangeKind.Added, node.Path, SubtreeSuffix(node)));

                continue;
            }

            var difference = DescribeDifference(earlier, node);
            if (difference.Length > 0)
                changes.Add(new XmlChange(XmlChangeKind.Modified, node.Path, difference));

            var reorder = DescribeReorder(earlier, node);
            if (reorder != null)
                changes.Add(new XmlChange(XmlChangeKind.Reordered, node.Path, reorder));
        }

        var pruned = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in before)
        {
            if (afterByPath.ContainsKey(node.Path))
                continue;

            var alreadyCovered = HasReportedAncestor(node.Path, pruned);
            pruned.Add(node.Path);

            if (!alreadyCovered)
                changes.Add(new XmlChange(XmlChangeKind.Removed, node.Path, SubtreeSuffix(node)));
        }

        return new XmlChangeSet(changes);
    }

    // ── Canonicalization ────────────────────────────────────────────────

    /// <summary>One element, reduced to everything the comparison treats as meaningful.</summary>
    private sealed class Node
    {
        internal Node(string path) => Path = path;

        internal string Path { get; }

        /// <summary>Attributes by prefixed name, namespace declarations excluded, order discarded.</summary>
        internal SortedDictionary<string, string> Attributes { get; } = new(StringComparer.Ordinal);

        /// <summary>The element's own text, or <c>null</c> when it has none that counts.</summary>
        internal string? Text { get; set; }

        /// <summary>The paths of the child elements, in document order.</summary>
        internal List<string> ChildPaths { get; } = [];
    }

    /// <summary>
    /// Flattens a document to its elements in document order.
    /// </summary>
    /// <remarks>
    /// Parsed with <see cref="LoadOptions.PreserveWhitespace"/> on purpose. The default would drop
    /// insignificant whitespace for us, but then the rule would be the parser's rather than this
    /// class's, and <c>xml:space="preserve"</c> would not be honoured.
    /// </remarks>
    private static List<Node> Flatten(string xml)
    {
        var root = XDocument.Parse(xml, LoadOptions.PreserveWhitespace).Root
                   ?? throw new ArgumentException("The document has no root element.", nameof(xml));

        var sink = new List<Node>();
        Visit(root, $"/{Prefixed(root.Name)}[1]", SpaceIsPreserved(root, inherited: false), sink);
        return sink;
    }

    private static void Visit(XElement element, string path, bool preserveSpace, List<Node> sink)
    {
        var node = new Node(path) { Text = DirectText(element, preserveSpace) };
        foreach (var attribute in element.Attributes())
        {
            if (attribute.IsNamespaceDeclaration)
                continue;

            node.Attributes[Prefixed(attribute.Name)] = attribute.Value;
        }

        sink.Add(node);

        var seen = new Dictionary<XName, int>();
        foreach (var child in element.Elements())
        {
            seen.TryGetValue(child.Name, out var count);
            seen[child.Name] = ++count;

            var childPath = $"{path}/{Prefixed(child.Name)}[{count}]";
            node.ChildPaths.Add(childPath);
            Visit(child, childPath, SpaceIsPreserved(child, preserveSpace), sink);
        }
    }

    /// <summary>
    /// The element's own text nodes, concatenated. Whitespace-only text is insignificant — it is how
    /// a document is indented — unless <c>xml:space="preserve"</c> is in scope, which is how Excel
    /// marks the whitespace inside an <c>a:t</c> that has to survive. Text that is not purely
    /// whitespace is kept verbatim, so a leading space in a run is never trimmed away.
    /// </summary>
    private static string? DirectText(XElement element, bool preserveSpace)
    {
        string? text = null;
        foreach (var child in element.Nodes().OfType<XText>())
        {
            if (!preserveSpace && string.IsNullOrWhiteSpace(child.Value))
                continue;

            text = text is null ? child.Value : text + child.Value;
        }

        return text;
    }

    private static bool SpaceIsPreserved(XElement element, bool inherited)
    {
        var attribute = element.Attribute(XNamespace.Xml + "space");
        return attribute is null ? inherited : attribute.Value == "preserve";
    }

    // ── Differences ─────────────────────────────────────────────────────

    private static string DescribeDifference(Node earlier, Node later)
    {
        var parts = new List<string>();

        foreach (var name in earlier.Attributes.Keys.Union(later.Attributes.Keys, StringComparer.Ordinal)
                     .OrderBy(name => name, StringComparer.Ordinal))
        {
            var had = earlier.Attributes.TryGetValue(name, out var before);
            var has = later.Attributes.TryGetValue(name, out var after);

            if (had && has)
            {
                if (!string.Equals(before, after, StringComparison.Ordinal))
                    parts.Add($"@{name}: '{before}' -> '{after}'");
            }
            else if (has)
                parts.Add($"+@{name}='{after}'");
            else
                parts.Add($"-@{name}='{before}'");
        }

        if (!string.Equals(earlier.Text, later.Text, StringComparison.Ordinal))
            parts.Add($"text: {Quote(earlier.Text)} -> {Quote(later.Text)}");

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Whether the children the node kept changed their relative order. Only children present on
    /// both sides are considered, so inserting or dropping a child never reads as a move; a genuine
    /// move does.
    /// </summary>
    private static string? DescribeReorder(Node earlier, Node later)
    {
        var kept = new HashSet<string>(later.ChildPaths, StringComparer.Ordinal);
        kept.IntersectWith(earlier.ChildPaths);
        if (kept.Count < 2)
            return null;

        var before = earlier.ChildPaths.Where(kept.Contains).ToList();
        var after = later.ChildPaths.Where(kept.Contains).ToList();
        if (before.SequenceEqual(after, StringComparer.Ordinal))
            return null;

        return $"{string.Join(", ", before.Select(LastSegment))} -> {string.Join(", ", after.Select(LastSegment))}";
    }

    private static bool HasReportedAncestor(string path, HashSet<string> region)
    {
        var parent = ParentPath(path);
        return parent is not null && region.Contains(parent);
    }

    private static string SubtreeSuffix(Node node) => node.ChildPaths.Count > 0 ? "(subtree)" : "";

    private static string? ParentPath(string path)
    {
        var slash = path.LastIndexOf('/');
        return slash <= 0 ? null : path[..slash];
    }

    private static string LastSegment(string path) => path[(path.LastIndexOf('/') + 1)..];

    private static string Quote(string? text) => text is null ? "(none)" : $"'{text}'";

    // ── Names ───────────────────────────────────────────────────────────

    /// <summary>
    /// The prefixes paths are written with. They are this class's own, not the document's: a part
    /// that declares <c>chart</c> where another declares <c>c</c> has to produce the same path, or
    /// the instrument would compare prefixes rather than meaning.
    /// </summary>
    private static readonly Dictionary<string, string> WellKnownPrefixes = new(StringComparer.Ordinal)
    {
        ["http://schemas.openxmlformats.org/drawingml/2006/main"] = "a",
        ["http://schemas.openxmlformats.org/drawingml/2006/chart"] = "c",
        ["http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"] = "xdr",
        ["http://schemas.openxmlformats.org/drawingml/2006/picture"] = "pic",
        ["http://schemas.openxmlformats.org/spreadsheetml/2006/main"] = "x",
        ["http://schemas.openxmlformats.org/officeDocument/2006/relationships"] = "r",
        ["http://schemas.openxmlformats.org/markup-compatibility/2006"] = "mc",
        ["http://www.w3.org/XML/1998/namespace"] = "xml",
        ["http://schemas.microsoft.com/office/drawing/2007/8/2/chart"] = "c14",
        ["http://schemas.microsoft.com/office/drawing/2012/chart"] = "c15",
        ["http://schemas.microsoft.com/office/drawing/2014/chart"] = "c16",
        ["http://schemas.microsoft.com/office/drawing/2014/chartex"] = "cx",
        ["http://schemas.microsoft.com/office/drawing/2010/main"] = "a14",
        ["http://schemas.microsoft.com/office/spreadsheetml/2014/revision"] = "xr",
    };

    /// <summary>
    /// A name as it appears in a path: <c>c:spPr</c> for a well-known namespace, the bare local name
    /// for no namespace, and Clark notation for anything unrecognised, so an unexpected namespace is
    /// visible in the failure rather than silently collapsed onto another.
    /// </summary>
    private static string Prefixed(XName name)
    {
        if (name.NamespaceName.Length == 0)
            return name.LocalName;

        return WellKnownPrefixes.TryGetValue(name.NamespaceName, out var prefix)
            ? $"{prefix}:{name.LocalName}"
            : $"{{{name.NamespaceName}}}{name.LocalName}";
    }
}
