using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;
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
    /// Spec 28 task 2 moved four key decoders out of <c>OpenXmlHelper</c> into
    /// <see cref="StyleDecoder"/> and wrote <see cref="StyleDecoder.FillKey"/> as a key-returning
    /// port of the mutating <c>OpenXmlHelper.LoadFill</c>. These pin the port: the same
    /// <c>&lt;fill&gt;</c> through both paths must give the same key. They pass from the moment
    /// the port lands — they are the proof that task 2 moved code without changing it.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(FillCases))]
    public async Task FillKey_agrees_with_the_mutating_fill_decoder(Fill fill, bool differential)
    {
        var throughKeyPath = StyleDecoder.FillKey(fill, differential, XLFillValue.Default.Key);

        var mutated = new XLFill();
        OpenXmlHelper.LoadFill(fill, mutated, differential);

        await Assert.That(throughKeyPath).IsEqualTo(mutated.Key);
    }

    public static IEnumerable<Func<(Fill Fill, bool Differential)>> FillCases()
    {
        yield return () => (new Fill(), false);
        yield return () => (new Fill(new PatternFill()), false);
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.None }), false);
        yield return () => (new Fill(new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" },
        }), false);
        yield return () => (new Fill(new PatternFill
        {
            PatternType = PatternValues.Solid,
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        }), true);
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.Solid }), false);
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.Solid }), true);
        yield return () => (new Fill(new PatternFill
        {
            PatternType = PatternValues.DarkGrid,
            ForegroundColor = new ForegroundColor { Indexed = 12U },
            BackgroundColor = new BackgroundColor { Indexed = 13U },
        }), false);
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.DarkGrid }), false);
    }

    /// <summary>
    /// The nine font fields both decoder families already handled must come out the same. The
    /// other three — name, family numbering and charset — are the divergence
    /// <see cref="A_conditional_format_font_keeps_its_name_family_and_charset"/> pins, and are
    /// deliberately not compared here: the mutating decoder cannot read them off a
    /// <c>&lt;x:font&gt;</c> at all.
    /// </summary>
    [Test]
    public async Task FontKey_agrees_with_the_mutating_font_decoder_on_the_nine_shared_fields()
    {
        var font = new Font(
            new Bold(),
            new Italic(),
            new Shadow(),
            new Strike(),
            new Underline { Val = UnderlineValues.Double },
            new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript },
            new FontSize { Val = 14.5D },
            new Color { Rgb = "FF112233" },
            new FontScheme { Val = FontSchemeValues.Major });

        var throughKeyPath = StyleDecoder.FontKey(font, XLFontValue.Default.Key);

        var mutated = XLStyle.CreateEmptyStyle();
        OpenXmlHelper.LoadFont(font, mutated.Font);
        var throughMutatingPath = ((XLFont)mutated.Font).Key;

        await Assert.That(throughKeyPath.Bold).IsEqualTo(throughMutatingPath.Bold);
        await Assert.That(throughKeyPath.Italic).IsEqualTo(throughMutatingPath.Italic);
        await Assert.That(throughKeyPath.Shadow).IsEqualTo(throughMutatingPath.Shadow);
        await Assert.That(throughKeyPath.Strikethrough).IsEqualTo(throughMutatingPath.Strikethrough);
        await Assert.That(throughKeyPath.Underline).IsEqualTo(throughMutatingPath.Underline);
        await Assert.That(throughKeyPath.VerticalAlignment)
            .IsEqualTo(throughMutatingPath.VerticalAlignment);
        await Assert.That(throughKeyPath.FontSize).IsEqualTo(throughMutatingPath.FontSize);
        await Assert.That(throughKeyPath.FontColor).IsEqualTo(throughMutatingPath.FontColor);
        await Assert.That(throughKeyPath.FontScheme).IsEqualTo(throughMutatingPath.FontScheme);
    }

    /// <summary>
    /// The alignment decoders agree on every attribute, provided the indent is one the mutating
    /// path's <c>IXLAlignment.Indent</c> setter tolerates. Where it does not they diverge, because
    /// that setter rewrites the horizontal alignment and throws for some legal files; spec 28 task
    /// 4 is where that stops mattering, since nothing decodes through the setter afterwards.
    /// </summary>
    [Test]
    public async Task AlignmentKey_agrees_with_the_mutating_alignment_decoder()
    {
        var alignment = new Alignment
        {
            Horizontal = HorizontalAlignmentValues.Left,
            Vertical = VerticalAlignmentValues.Top,
            Indent = 3U,
            ReadingOrder = 2U,
            WrapText = true,
            TextRotation = 45U,
            ShrinkToFit = false,
            RelativeIndent = 1,
            JustifyLastLine = true,
        };

        var throughKeyPath = StyleDecoder.AlignmentKey(alignment, XLAlignmentValue.Default.Key);

        var mutated = XLStyle.CreateEmptyStyle();
        OpenXmlHelper.LoadAlignment(alignment, mutated.Alignment);

        await Assert.That(throughKeyPath).IsEqualTo(mutated.Key.Alignment);
    }

    /// <summary>
    /// The inline <c>&lt;numFmt&gt;</c> decoder agrees with the mutating one, including the
    /// built-in-id branch that discards the format code.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(InlineNumberFormatCases))]
    public async Task NumberFormatKey_agrees_with_the_mutating_number_format_decoder(
        NumberingFormat numberingFormat)
    {
        var throughKeyPath =
            StyleDecoder.NumberFormatKey(numberingFormat, XLNumberFormatValue.Default.Key);

        var mutated = XLStyle.CreateEmptyStyle();
        OpenXmlHelper.LoadNumberFormat(numberingFormat, mutated.NumberFormat);

        await Assert.That(throughKeyPath).IsEqualTo(((XLNumberFormat)mutated.NumberFormat).Key);
    }

    public static IEnumerable<Func<NumberingFormat>> InlineNumberFormatCases()
    {
        yield return () => new NumberingFormat { NumberFormatId = 5U, FormatCode = "0.00" };
        yield return () => new NumberingFormat { NumberFormatId = 164U, FormatCode = "0.000" };
        yield return () => new NumberingFormat { NumberFormatId = 200U, FormatCode = "#,##0" };
        yield return () => new NumberingFormat { FormatCode = "yyyy-mm-dd" };
        yield return () => new NumberingFormat { NumberFormatId = 14U };
    }

    /// <summary>
    /// A &lt;border diagonalUp="1"/&gt; with no &lt;diagonal&gt; child decoded one way through
    /// BorderToXLibur (flags read only inside the &lt;diagonal&gt; guard) and another through
    /// LoadBorder (flags read unconditionally).
    /// <para>
    /// Spec 28 task 3 settled it towards the schema: ECMA-376 Part 1 §18.8.4 declares
    /// <c>diagonalUp</c> and <c>diagonalDown</c> as attributes of the <c>border</c> element, not
    /// as part of its <c>diagonal</c> child, so the unconditional read is correct and the guard was
    /// the bug. Only that one rule now exists — this test compares the surviving mutating decoder
    /// against <see cref="StyleDecoder.BorderKey"/> and they must agree.
    /// </para>
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

        var throughKeyPath = StyleDecoder.BorderKey(border, XLBorderValue.Default.Key);

        var mutated = XLStyle.CreateEmptyStyle();
        OpenXmlHelper.LoadBorder(border, mutated.Border);
        var throughMutatingPath = ((XLBorder)mutated.Border).Key.Normalize();

        await Assert.That(throughKeyPath.DiagonalUp).IsEqualTo(throughMutatingPath.DiagonalUp);
        await Assert.That(throughKeyPath.DiagonalDown).IsEqualTo(throughMutatingPath.DiagonalDown);

        // ...and they must agree on the value the file actually stated, not merely with each other.
        await Assert.That(throughKeyPath.DiagonalUp).IsEqualTo(up);
        await Assert.That(throughKeyPath.DiagonalDown).IsEqualTo(down);
    }

    /// <summary>
    /// The evidence the diagonal decision rests on, pinned rather than asserted in prose: the SDK's
    /// element model is generated from the ECMA-376 schema, and it serializes <c>diagonalUp</c> and
    /// <c>diagonalDown</c> as attributes on <c>&lt;border&gt;</c> itself. An attribute of
    /// <c>border</c> is a sibling of the <c>&lt;diagonal&gt;</c> child element, so nothing ties
    /// reading the flags to that child being present. See the remarks on
    /// <see cref="StyleDecoder.BorderKey"/>.
    /// </summary>
    [Test]
    public async Task The_diagonal_flags_are_attributes_of_border_not_of_its_diagonal_child()
    {
        var border = new Border { DiagonalUp = true, DiagonalDown = true };

        // <x:border diagonalUp="1" diagonalDown="1" /> - both on the border element itself, and no
        // <diagonal> child in sight.
        await Assert.That(border.OuterXml).Contains("<x:border diagonalUp=\"1\" diagonalDown=\"1\"");
        await Assert.That(border.Elements<DiagonalBorder>().Any()).IsFalse();
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
