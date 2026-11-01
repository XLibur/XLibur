using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// Every holder of a conditional format reads its formulas relative to one cell, the anchor
/// (<see cref="XLConditionalFormat.AnchorOf"/>): the top-left of the rectangle bounding its range
/// (issue #499). A rule set with <c>SetRanges</c> can list an area first that is not at that corner,
/// which is where a holder reading the first area's first cell instead goes wrong.
/// </summary>
public class ConditionalFormatAnchorTests
{
    /// <summary>
    /// A save with the default <see cref="SaveOptions"/> consolidates conditional formats, rewriting a
    /// rule's range and re-expressing its formulas for the new one. <c>B1:B5 A3:A5</c> with
    /// <c>A1&gt;0</c> has its anchor at <c>A1</c>, so every cell tests itself. Consolidation used to
    /// re-express the formula for the first consolidated area's first cell, <c>B1</c>, which made every
    /// cell test its right-hand neighbour, and the live model kept that, so each further save moved it
    /// another column.
    /// </summary>
    [Test]
    public async Task Consolidation_keeps_a_formula_on_its_anchor_through_repeated_saves()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("B1:B5").AddConditionalFormat()
            .SetRanges(new[] { ws.Range("B1:B5"), ws.Range("A3:A5") })
            .WhenIsTrue("=A1>0").Fill.SetBackgroundColor(XLColor.Red);

        // Both streams stay open: a save reads the package the one before it wrote.
        using var firstSave = new MemoryStream();
        using var secondSave = new MemoryStream();
        var first = SavedRule(wb, firstSave);
        var second = SavedRule(wb, secondSave);

        await Assert.That(first).IsEqualTo("A1 anchors A1>0");
        await Assert.That(second).IsEqualTo(first);
    }

    /// <summary>
    /// Copying part of a rule re-expresses its formulas for the copy's range. The rule over
    /// <c>D1:E4 B2:B5</c> with <c>B1&gt;0</c> has its anchor at <c>B1</c>, the corner of the rectangle
    /// bounding both areas, where neither area starts; so every cell tests itself. Copying
    /// <c>D1:E4</c> to <c>G1</c> lands it on <c>G1:H4</c>, where every cell must still test itself:
    /// <c>G1&gt;0</c>. Read from an area's first cell, <c>D1</c> or <c>B2</c>, rather than from the
    /// anchor, the formula comes out pointing a row or two columns away.
    /// </summary>
    [Test]
    public async Task Copying_part_of_a_rule_keeps_each_cell_reading_the_same_cells()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        source.Range("D1:E4").AddConditionalFormat()
            .SetRanges(new[] { source.Range("D1:E4"), source.Range("B2:B5") })
            .WhenIsTrue("=B1>0").Fill.SetBackgroundColor(XLColor.Red);
        var target = wb.AddWorksheet("Target");

        source.Range("D1:E4").CopyTo(target.Cell("G1"));

        var copy = (XLConditionalFormat)target.ConditionalFormats.Single();
        await Assert.That(string.Join(" ", copy.Areas)).IsEqualTo("G1:H4");
        await Assert.That(copy.Values[1].Value).IsEqualTo("G1>0");
    }

    /// <summary>
    /// The blank, error, text and date rules have a formula the save builds from a template, and Excel
    /// writes that formula relative to the anchor too. Measured over COM: Excel saves a "blanks" rule
    /// on <c>B1:B5,A3:A5</c> as <c>LEN(TRIM(A1))=0</c>, an "errors" rule on <c>F3:F5,E1:E2</c> as
    /// <c>ISERROR(E1)</c>, and a "does not contain" rule on <c>J2:J5,H4:H6</c> as
    /// <c>ISERROR(SEARCH("ab",H2))</c>, <c>H2</c> being in neither area. The save used to write each
    /// template for the first area's first cell (<c>B1</c>, <c>F3</c>, <c>J2</c>), so Excel tested
    /// every cell's neighbour.
    /// </summary>
    [Test]
    public async Task A_template_formula_is_written_for_the_anchor()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        AddRule(ws, "B1:B5", "A3:A5").WhenIsBlank().Fill.SetBackgroundColor(XLColor.Red);
        AddRule(ws, "F3:F5", "E1:E2").WhenIsError().Fill.SetBackgroundColor(XLColor.Red);
        AddRule(ws, "J2:J5", "H4:H6").WhenNotContains("ab").Fill.SetBackgroundColor(XLColor.Red);
        AddRule(ws, "M3:M5", "L1:L2").WhenDateIs(XLTimePeriod.Today).Fill.SetBackgroundColor(XLColor.Red);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        var formulas = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .Descendants<S.Formula>()
            .Select(f => f.Text)
            .ToList();

        await Assert.That(formulas).IsEquivalentTo(new[]
        {
            "LEN(TRIM(A1))=0",
            "ISERROR(E1)",
            "ISERROR(SEARCH(\"ab\",H2))",
            "FLOOR(L1,1)=TODAY()",
        });

        static IXLConditionalFormat AddRule(IXLWorksheet ws, string first, string second)
            => ws.Range(first).AddConditionalFormat().SetRanges(new[] { ws.Range(first), ws.Range(second) });
    }

    /// <summary>
    /// Saves <paramref name="wb"/> to <paramref name="ms"/> with the default options and reads its one
    /// rule back as <c>anchor anchors formula</c>, the anchor being the top-left of the saved range.
    /// </summary>
    private static string SavedRule(XLWorkbook wb, MemoryStream ms)
    {
        wb.SaveAs(ms);
        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single();
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
        var formatting = worksheet.Elements<S.ConditionalFormatting>().Single();
        var areas = (formatting.SequenceOfReferences?.InnerText ?? string.Empty)
            .Split(' ')
            .Select(a => Area.Parse(a))
            .ToList();
        var anchor = new Area(new Point(areas.Min(a => a.TopRow), areas.Min(a => a.LeftColumn)));
        return $"{anchor} anchors {formatting.Descendants<S.Formula>().Single().Text}";
    }
}
