using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Spec 28: the same style XML is decoded by two implementations chosen by which element it came
/// from, and they have diverged. These tests pin the divergences.
/// <para>
/// Two of the three fail on the tree at c569b95a and are made to pass by spec 28 tasks 3 and 4:
/// the conditional-format font losing its name, family and charset, and the diagonal border flags
/// decoding differently from the two paths. The third — dxf table growth — passes on c569b95a,
/// which disproves that premise; see the remarks on
/// <see cref="Round_tripping_does_not_grow_the_dxf_table"/> for why.
/// </para>
/// </summary>
public class StyleDecoderTests
{
    /// <summary>
    /// OpenXmlHelper.LoadFont takes an untyped OpenXmlElement and looks for RunFont, FontFamily and
    /// no charset at all — the &lt;x:rPr&gt; spellings. A dxf hands it a &lt;x:font&gt;, whose
    /// corresponding children are the unrelated types FontName, FontFamilyNumbering and FontCharSet.
    /// The writer emits all three (WorkbookStylesPartWriter.AppendFontScalarElements), so they reach
    /// the file and are dropped on the way back.
    /// </summary>
    [Test]
    public async Task A_conditional_format_font_keeps_its_name_family_and_charset()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var cf = ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5);
            cf.Font.FontName = "Arial";
            cf.Font.FontCharSet = XLFontCharSet.Arabic;
            cf.Font.FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var format = reloaded.Worksheet("Sheet1").ConditionalFormats.Single();

        await Assert.That(format.Style.Font.FontName).IsEqualTo("Arial");
        await Assert.That(format.Style.Font.FontCharSet).IsEqualTo(XLFontCharSet.Arabic);
        await Assert.That(format.Style.Font.FontFamilyNumbering)
            .IsEqualTo(XLFontFamilyNumberingValues.Swiss);
    }

    /// <summary>
    /// Spec 28 predicted this would fail: the writer's reuse-map decode
    /// (<c>WorkbookStylesPartWriter.FillDifferentialFormatsCollection</c>) reads four of the six
    /// children a dxf may carry, while the pivot reader reads five, so an alignment-bearing pivot
    /// dxf was expected never to match its own map entry and to be appended again on every save.
    /// <para>
    /// <b>The premise is disproved, and measured flat at 1, 1, 1, 1 across four saves.</b>
    /// <c>AddDifferentialFormats</c> calls <c>differentialFormats.RemoveAllChildren()</c> on the
    /// line immediately before <c>FillDifferentialFormatsCollection</c>, so that method always
    /// iterates an empty collection and the reuse map is always empty. The decoder mismatch is
    /// real, but it cannot produce growth because the decode side never runs on any input. The
    /// dxf table is rebuilt from the live object model on every save instead.
    /// </para>
    /// <para>
    /// The test is kept as a regression guard rather than deleted: it is what would catch the
    /// growth if the <c>RemoveAllChildren</c> call were ever moved or removed, which would put the
    /// reuse map back into service and make the two decodes have to agree exactly. The fixture is
    /// not vacuous — the pivot format loads with <c>Alignment.Horizontal == Center</c> and a
    /// non-default style value, which is precisely the input the premise needed.
    /// </para>
    /// </summary>
    [Test]
    public async Task Round_tripping_does_not_grow_the_dxf_table()
    {
        var bytes = BuildWorkbookWithAnAlignedPivotFormat();
        var counts = new List<int> { CountDxfs(bytes) };

        for (var i = 0; i < 3; i++)
        {
            bytes = ReSave(bytes);
            counts.Add(CountDxfs(bytes));
        }

        await Assert.That(counts.Distinct().Count())
            .IsEqualTo(1)
            .Because($"dxf count per round trip: {string.Join(", ", counts)}");
    }

    /// <summary>
    /// The fixture the growth test rests on: the alignment-bearing pivot dxf must actually reach
    /// <see cref="XLPivotFormat.DxfStyleValue"/>, or a flat dxf count would prove nothing. Pinned
    /// separately so that a future change which silently stops loading pivot formats shows up here
    /// rather than as a growth test that passes for the wrong reason.
    /// </summary>
    [Test]
    public async Task An_alignment_bearing_pivot_dxf_reaches_the_pivot_format()
    {
        var bytes = BuildWorkbookWithAnAlignedPivotFormat();

        using var input = new MemoryStream(bytes, writable: false);
        using var wb = new XLWorkbook(input);
        var pt = (XLPivotTable)wb.Worksheet("Pivots").PivotTables.Single();

        await Assert.That(pt.Formats.Count).IsEqualTo(1);
        await Assert.That(pt.Formats[0].DxfStyleValue.Alignment.Horizontal)
            .IsEqualTo(XLAlignmentHorizontalValues.Center);
        await Assert.That(pt.Formats[0].DxfStyleValue.Equals(XLStyleValue.Default)).IsFalse();
    }

    /// <summary>
    /// A &lt;border diagonalUp="1"/&gt; with no &lt;diagonal&gt; child decodes one way through
    /// BorderToXLibur (flags read only inside the &lt;diagonal&gt; guard) and another through
    /// LoadBorder (flags read unconditionally). One of the two is wrong per ECMA-376 CT_Border;
    /// spec 28 task 3 decides which.
    /// </summary>
    [Test]
    [Arguments(true, false)]
    [Arguments(false, true)]
    [Arguments(true, true)]
    public async Task The_diagonal_flags_decode_the_same_from_both_paths(bool up, bool down)
    {
        var border = new Border
        {
            DiagonalUp = up,
            DiagonalDown = down,
        };

        var throughKeyPath = OpenXmlHelper.BorderToXLibur(border, XLBorderValue.Default.Key);

        var mutated = XLStyle.CreateEmptyStyle();
        OpenXmlHelper.LoadBorder(border, mutated.Border);
        var throughMutatingPath = ((XLBorder)mutated.Border).Key.Normalize();

        await Assert.That(throughKeyPath.DiagonalUp).IsEqualTo(throughMutatingPath.DiagonalUp);
        await Assert.That(throughKeyPath.DiagonalDown).IsEqualTo(throughMutatingPath.DiagonalDown);
    }

    /// <summary>
    /// A workbook whose pivot table carries one <c>&lt;format&gt;</c> pointing at a dxf that states
    /// only <c>&lt;alignment&gt;</c>. The pivot reader decodes that alignment into
    /// <see cref="XLPivotFormat.DxfStyleValue"/>; the writer's reuse-map decode does not read
    /// <c>&lt;alignment&gt;</c> at all, so the two do not meet.
    /// </summary>
    private static byte[] BuildWorkbookWithAnAlignedPivotFormat()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            data.FirstCell().InsertData(new object[]
            {
                ("Pastry", "Sold"),
                ("Waffle", 3),
                ("Donut", 5),
            });

            var pivots = wb.AddWorksheet("Pivots");
            var pt = pivots.PivotTables.Add("pvt", pivots.Cell("A1"), data.Range("A1:B3"));
            pt.RowLabels.Add("Pastry");
            pt.Values.Add("Sold");

            wb.SaveAs(ms);
        }

        var bytes = ms.ToArray();
        var patched = new MemoryStream();
        patched.Write(bytes, 0, bytes.Length);
        patched.Position = 0;

        using (var doc = SpreadsheetDocument.Open(patched, true))
        {
            var stylesPart = doc.WorkbookPart!.WorkbookStylesPart!;
            var stylesheet = stylesPart.Stylesheet!;
            var dxfs = stylesheet.DifferentialFormats;
            if (dxfs is null)
            {
                dxfs = new DifferentialFormats();
                stylesheet.DifferentialFormats = dxfs;
            }

            var dxfId = dxfs.ChildElements.Count;
            dxfs.AppendChild(new DifferentialFormat(
                new Alignment { Horizontal = HorizontalAlignmentValues.Center }));
            dxfs.Count = (uint)dxfs.ChildElements.Count;
            stylesheet.Save();

            var pivotPart = doc.WorkbookPart.WorksheetParts
                .SelectMany(wsp => wsp.GetPartsOfType<PivotTablePart>())
                .Single();
            var definition = pivotPart.PivotTableDefinition!;
            definition.Formats = new Formats(
                new Format(new PivotArea { Outline = false, FieldPosition = 0U })
                {
                    FormatId = (uint)dxfId,
                })
            {
                Count = 1U,
            };
            definition.Save();
        }

        return patched.ToArray();
    }

    private static byte[] ReSave(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.ToArray();
    }

    private static int CountDxfs(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var doc = SpreadsheetDocument.Open(input, false);
        var dxfs = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats;
        return dxfs?.ChildElements.Count ?? 0;
    }
}
