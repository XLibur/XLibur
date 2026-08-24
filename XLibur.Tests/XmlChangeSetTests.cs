using System.Threading.Tasks;
using XLibur.Tests.Utils;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Tests;

/// <summary>
/// The change-set instrument proving what it claims. Every case here is a property some later test
/// leans on: that serialization noise is absorbed, that a real edit is reported, and — the one that
/// matters most — that an extra mutation nobody asked for shows up beside the intended one.
/// </summary>
public class XmlChangeSetTests
{
    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>A minimal chart part whose <c>c:spPr</c> body is the argument.</summary>
    private static string ChartWith(string shapeProperties) => $"""
        <c:chartSpace xmlns:c="{ChartNs}" xmlns:a="{DrawingNs}">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:idx val="0"/>
                  <c:spPr>{shapeProperties}</c:spPr>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        """;

    private const string ShapeProperties =
        "/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:barChart[1]/c:ser[1]/c:spPr[1]";

    /// <summary>Asserts that the difference between the two documents is exactly these changes.</summary>
    private static async Task AssertChanges(string before, string after, params string[] expected) =>
        await Assert.That(XmlChangeSet.Between(before, after).ToString())
            .IsEqualTo(XmlChangeSet.Expect(expected));

    // ── Serialization noise is absorbed ─────────────────────────────────

    [Test]
    public async Task An_unchanged_document_has_an_empty_change_set()
    {
        var xml = ChartWith("<a:solidFill><a:srgbClr val=\"ED7D31\"/></a:solidFill>");

        await Assert.That(XmlChangeSet.Between(xml, xml).IsEmpty).IsTrue();
        await AssertChanges(xml, xml);
    }

    [Test]
    public async Task Prefixes_attribute_order_indentation_and_comments_are_not_changes()
    {
        const string before = $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="{ChartNs}" xmlns:a="{DrawingNs}">
              <!-- authored by hand -->
              <c:chart>
                <c:plotArea><c:layout/></c:plotArea>
              </c:chart>
            </c:chartSpace>
            """;

        // The same document: a different prefix, the namespace declared where it is first used, no
        // indentation and no comment.
        const string after =
            $"""<chart:chartSpace xmlns:chart="{ChartNs}"><chart:chart><chart:plotArea>""" +
            """<chart:layout/></chart:plotArea></chart:chart></chart:chartSpace>""";

        await AssertChanges(before, after);
    }

    /// <summary>
    /// Acceptance criterion 5: a part whose DOM was materialised and re-serialized with no model
    /// edits canonicalizes to an empty change set. This is what makes the instrument usable against
    /// a save path at all — merely loading a part's DOM rewrites its bytes, so a byte comparison
    /// would report that rewrite as though it were the edit under test.
    /// </summary>
    [Test]
    public async Task A_dom_round_trip_with_no_model_edits_changes_nothing()
    {
        const string authored = $"""
            <c:chartSpace xmlns:c="{ChartNs}" xmlns:a="{DrawingNs}">
              <c:chart>
                <c:title>
                  <c:tx>
                    <c:rich>
                      <a:bodyPr rot="0" vert="horz"/>
                      <a:lstStyle/>
                      <a:p>
                        <a:r>
                          <a:rPr lang="en-US" sz="1600" b="1"/>
                          <a:t>Units and price</a:t>
                        </a:r>
                      </a:p>
                    </c:rich>
                  </c:tx>
                  <c:overlay val="0"/>
                </c:title>
              </c:chart>
            </c:chartSpace>
            """;

        var chartSpace = new C.ChartSpace(authored);

        // Touching a child is what materialises the DOM. Until something does, the SDK holds the
        // part as the raw text it was handed and gives it straight back, which is precisely the
        // asymmetry that makes byte comparison unreliable: an unread part keeps its bytes and a read
        // one does not.
        _ = chartSpace.ChildElements.Count;
        var reserialized = chartSpace.OuterXml;

        await Assert.That(reserialized).IsNotEqualTo(authored)
            .Because("If materialising the DOM did not rewrite the bytes, this test would prove nothing.");
        await AssertChanges(authored, reserialized);
    }

    // ── Whitespace ──────────────────────────────────────────────────────

    [Test]
    public async Task Whitespace_a_run_asked_to_keep_is_significant()
    {
        const string before = $"""<a:t xmlns:a="{DrawingNs}" xml:space="preserve"> Units </a:t>""";
        const string after = $"""<a:t xmlns:a="{DrawingNs}" xml:space="preserve">Units</a:t>""";

        await AssertChanges(before, after, "modified /a:t[1] text: ' Units ' -> 'Units'");
    }

    [Test]
    public async Task Text_that_is_not_only_whitespace_keeps_its_edges()
    {
        const string before = $"""<a:t xmlns:a="{DrawingNs}"> Units </a:t>""";
        const string after = $"""<a:t xmlns:a="{DrawingNs}">Units</a:t>""";

        await AssertChanges(before, after, "modified /a:t[1] text: ' Units ' -> 'Units'");
    }

    // ── Real edits are reported, and reported once ──────────────────────

    [Test]
    public async Task A_replaced_colour_is_one_modification()
    {
        await AssertChanges(
            ChartWith("<a:solidFill><a:srgbClr val=\"ED7D31\"/></a:solidFill>"),
            ChartWith("<a:solidFill><a:srgbClr val=\"00B050\"/></a:solidFill>"),
            $"modified {ShapeProperties}/a:solidFill[1]/a:srgbClr[1] @val: 'ED7D31' -> '00B050'");
    }

    [Test]
    public async Task An_attribute_that_appears_and_one_that_vanishes_are_both_named()
    {
        await AssertChanges(
            ChartWith("<a:ln w=\"19050\" cap=\"rnd\"/>"),
            ChartWith("<a:ln w=\"38100\" algn=\"ctr\"/>"),
            $"modified {ShapeProperties}/a:ln[1] +@algn='ctr', -@cap='rnd', @w: '19050' -> '38100'");
    }

    [Test]
    public async Task A_grafted_subtree_is_reported_at_its_root_and_marked_as_one()
    {
        await AssertChanges(
            ChartWith("<a:solidFill><a:srgbClr val=\"ED7D31\"/></a:solidFill>"),
            ChartWith("<a:solidFill><a:srgbClr val=\"ED7D31\"/></a:solidFill>" +
                      "<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"203864\"/></a:solidFill></a:ln>"),
            $"added {ShapeProperties}/a:ln[1] (subtree)");
    }

    [Test]
    public async Task A_pruned_subtree_is_reported_at_its_root()
    {
        await AssertChanges(
            ChartWith("<a:gradFill><a:gsLst><a:gs pos=\"0\"/></a:gsLst></a:gradFill>"),
            ChartWith(""),
            $"removed {ShapeProperties}/a:gradFill[1] (subtree)");
    }

    // ── Order ───────────────────────────────────────────────────────────

    /// <summary>
    /// The property the DrawingML sequence rules turn on: <c>a:ln</c> after the fill, not before it.
    /// A patch that writes the right elements in the wrong order is still a wrong patch.
    /// </summary>
    [Test]
    public async Task Moving_a_child_past_a_differently_named_sibling_is_a_reorder()
    {
        await AssertChanges(
            ChartWith("<a:solidFill/><a:ln w=\"19050\"/>"),
            ChartWith("<a:ln w=\"19050\"/><a:solidFill/>"),
            $"reordered {ShapeProperties} a:solidFill[1], a:ln[1] -> a:ln[1], a:solidFill[1]");
    }

    [Test]
    public async Task Inserting_a_child_is_an_addition_and_not_a_reorder()
    {
        await AssertChanges(
            ChartWith("<a:solidFill/><a:effectLst/>"),
            ChartWith("<a:solidFill/><a:ln w=\"19050\"/><a:effectLst/>"),
            $"added {ShapeProperties}/a:ln[1]");
    }

    // ── The point of the instrument ─────────────────────────────────────

    /// <summary>
    /// The gate spec 16 task 1 is measured by. A test that only asserts the intended edit landed
    /// passes whether or not the operation also did something nobody asked for; a change set states
    /// the whole difference, so the stray mutation has nowhere to hide.
    /// </summary>
    [Test]
    public async Task A_stray_mutation_alongside_the_intended_one_is_reported_too()
    {
        // The intended edit is the fill colour. The outline losing its cap is the stray mutation.
        await AssertChanges(
            ChartWith("<a:solidFill><a:srgbClr val=\"ED7D31\"/></a:solidFill><a:ln w=\"19050\" cap=\"rnd\"/>"),
            ChartWith("<a:solidFill><a:srgbClr val=\"00B050\"/></a:solidFill><a:ln w=\"19050\"/>"),
            $"modified {ShapeProperties}/a:solidFill[1]/a:srgbClr[1] @val: 'ED7D31' -> '00B050'",
            $"modified {ShapeProperties}/a:ln[1] -@cap='rnd'");
    }

    [Test]
    public async Task An_unrecognised_namespace_is_named_rather_than_collapsed()
    {
        const string vendor = "urn:example:vendor";

        await AssertChanges(
            $"""<c:chartSpace xmlns:c="{ChartNs}"><c:chart/></c:chartSpace>""",
            $"""<c:chartSpace xmlns:c="{ChartNs}"><c:chart/><v:ext xmlns:v="{vendor}"/></c:chartSpace>""",
            $"added /c:chartSpace[1]/{{{vendor}}}ext[1]");
    }
}
