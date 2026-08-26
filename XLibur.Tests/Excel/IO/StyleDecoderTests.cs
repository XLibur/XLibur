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
internal class StyleDecoderTests
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
    /// <see cref="StyleDecoder.FillKey"/> is a key-returning port of the mutating fill decoder spec
    /// 28 deleted. These pinned the port when both existed side by side; now that only one decoder
    /// remains they state the expected key outright, which is the stronger form of the same test.
    /// <para>
    /// Note the asymmetry, carried over deliberately: the <c>None</c> pattern touches no colour,
    /// while <c>Solid</c> and the patterned branch both default a missing background to index 64.
    /// </para>
    /// </summary>
    [Test]
    [MethodDataSource(nameof(FillCases))]
    public async Task FillKey_decodes_a_fill_element(Fill fill, bool differential, XLFillKey expected)
    {
        var actual = StyleDecoder.FillKey(fill, differential, XLFillValue.Default.Key);

        await Assert.That(actual).IsEqualTo(expected);
    }

    public static IEnumerable<Func<(Fill Fill, bool Differential, XLFillKey Expected)>> FillCases()
    {
        var defaults = XLFillValue.Default.Key;
        var transparent = XLColor.FromIndex(64).Key;

        // No <patternFill> at all: nothing is stated, so nothing changes.
        yield return () => (new Fill(), false, defaults);

        // <patternFill> with no patternType attribute is read as solid.
        yield return () => (new Fill(new PatternFill()), false,
            defaults with { PatternType = XLFillPatternValues.Solid, BackgroundColor = transparent });

        // pattern="none" leaves both colours alone.
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.None }), false,
            defaults with { PatternType = XLFillPatternValues.None });

        // A non-differential solid fill takes its background from fgColor.
        yield return () => (new Fill(new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" },
        }), false, defaults with
        {
            PatternType = XLFillPatternValues.Solid,
            BackgroundColor = XLColor.FromArgb(0xFF, 0x00, 0x00).Key,
        });

        // A differential solid fill takes its background from bgColor.
        yield return () => (new Fill(new PatternFill
        {
            PatternType = PatternValues.Solid,
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        }), true, defaults with
        {
            PatternType = XLFillPatternValues.Solid,
            BackgroundColor = XLColor.FromArgb(0x00, 0xFF, 0x00).Key,
        });

        // Solid with no colour stated at all defaults to transparent, either way round.
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.Solid }), false,
            defaults with { PatternType = XLFillPatternValues.Solid, BackgroundColor = transparent });
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.Solid }), true,
            defaults with { PatternType = XLFillPatternValues.Solid, BackgroundColor = transparent });

        // A real pattern reads fgColor as the pattern colour and bgColor as the background.
        yield return () => (new Fill(new PatternFill
        {
            PatternType = PatternValues.DarkGrid,
            ForegroundColor = new ForegroundColor { Indexed = 12U },
            BackgroundColor = new BackgroundColor { Indexed = 13U },
        }), false, defaults with
        {
            PatternType = XLFillPatternValues.DarkGrid,
            PatternColor = XLColor.FromIndex(12).Key,
            BackgroundColor = XLColor.FromIndex(13).Key,
        });

        // A pattern with neither colour keeps the default pattern colour but still defaults the
        // background to transparent.
        yield return () => (new Fill(new PatternFill { PatternType = PatternValues.DarkGrid }), false,
            defaults with
            {
                PatternType = XLFillPatternValues.DarkGrid,
                BackgroundColor = transparent,
            });
    }

    /// <summary>
    /// <c>&lt;x:font&gt;</c> decodes through the typed <c>CT_Font</c> children — all twelve of
    /// them. The last three are the divergence
    /// <see cref="A_conditional_format_font_keeps_its_name_family_and_charset"/> pins: the deleted
    /// untyped decoder looked for the <c>&lt;x:rPr&gt;</c> spellings and so could not read them off
    /// a font at all.
    /// </summary>
    [Test]
    public async Task FontKey_decodes_every_font_child_including_name_family_and_charset()
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
            new FontScheme { Val = FontSchemeValues.Major },
            new FontName { Val = "Arial" },
            new FontFamilyNumbering { Val = 2 },
            new FontCharSet { Val = 178 });

        var key = StyleDecoder.FontKey(font, XLFontValue.Default.Key);

        await Assert.That(key.Bold).IsTrue();
        await Assert.That(key.Italic).IsTrue();
        await Assert.That(key.Shadow).IsTrue();
        await Assert.That(key.Strikethrough).IsTrue();
        await Assert.That(key.Underline).IsEqualTo(XLFontUnderlineValues.Double);
        await Assert.That(key.VerticalAlignment)
            .IsEqualTo(XLFontVerticalTextAlignmentValues.Superscript);
        await Assert.That(key.FontSize).IsEqualTo(14.5D);
        await Assert.That(key.FontColor).IsEqualTo(XLColor.FromArgb(0x11, 0x22, 0x33).Key);
        await Assert.That(key.FontScheme).IsEqualTo(XLFontScheme.Major);
        await Assert.That(key.FontName).IsEqualTo("Arial");
        await Assert.That(key.FontFamilyNumbering).IsEqualTo(XLFontFamilyNumberingValues.Swiss);
        await Assert.That(key.FontCharSet).IsEqualTo(XLFontCharSet.Arabic);
    }

    /// <summary>
    /// The rich-text counterpart, over the <c>&lt;x:rPr&gt;</c> spellings. It reads
    /// <c>&lt;charset&gt;</c> too, which the decoder it replaced never looked for on this path
    /// either — the same omission as on the dxf path, and for the same reason.
    /// </summary>
    [Test]
    public async Task RunFontKey_decodes_the_rich_text_spellings_including_charset()
    {
        var runProperties = new RunProperties(
            new Bold(),
            new RunFont { Val = "Arial" },
            new FontFamily { Val = 2 },
            new RunPropertyCharSet { Val = 178 },
            new FontSize { Val = 11D });

        var key = StyleDecoder.RunFontKey(runProperties, XLFontValue.Default.Key);

        await Assert.That(key.Bold).IsTrue();
        await Assert.That(key.FontName).IsEqualTo("Arial");
        await Assert.That(key.FontFamilyNumbering).IsEqualTo(XLFontFamilyNumberingValues.Swiss);
        await Assert.That(key.FontCharSet).IsEqualTo(XLFontCharSet.Arabic);
        await Assert.That(key.FontSize).IsEqualTo(11D);
    }

    /// <summary>
    /// Every <c>&lt;alignment&gt;</c> attribute decodes into the corresponding key field.
    /// </summary>
    [Test]
    public async Task AlignmentKey_decodes_every_alignment_attribute()
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

        var key = StyleDecoder.AlignmentKey(alignment, XLAlignmentValue.Default.Key);

        await Assert.That(key.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Left);
        await Assert.That(key.Vertical).IsEqualTo(XLAlignmentVerticalValues.Top);
        await Assert.That(key.Indent).IsEqualTo(3);
        await Assert.That(key.ReadingOrder).IsEqualTo(XLAlignmentReadingOrderValues.RightToLeft);
        await Assert.That(key.WrapText).IsTrue();
        await Assert.That(key.TextRotation).IsEqualTo(45);
        await Assert.That(key.ShrinkToFit).IsFalse();
        await Assert.That(key.RelativeIndent).IsEqualTo(1);
        await Assert.That(key.JustifyLastLine).IsTrue();
    }

    /// <summary>
    /// An <c>&lt;alignment&gt;</c> that states an indent alongside a centred horizontal alignment
    /// is legal OOXML, and now decodes as written.
    /// <para>
    /// Before spec 28 the pivot path decoded dxf alignments by writing through
    /// <c>IXLAlignment</c>, whose <c>Indent</c> setter rewrites a <c>General</c> horizontal
    /// alignment to <c>Left</c> and <b>throws</b> <see cref="ArgumentException"/> for any indent
    /// above zero on a horizontal alignment that is not left, right or distributed. Loading a
    /// workbook whose pivot dxf carried <c>&lt;alignment horizontal="center" indent="2"/&gt;</c>
    /// therefore failed outright. Decoding to a key touches no setter, so the file loads and the
    /// value survives.
    /// </para>
    /// </summary>
    [Test]
    public async Task An_indent_with_a_centred_horizontal_alignment_no_longer_throws()
    {
        var alignment = new Alignment
        {
            Horizontal = HorizontalAlignmentValues.Center,
            Indent = 2U,
        };

        var key = StyleDecoder.AlignmentKey(alignment, XLAlignmentValue.Default.Key);

        await Assert.That(key.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
        await Assert.That(key.Indent).IsEqualTo(2);
    }

    /// <summary>
    /// The same, for the quieter half of that setter: an indent with no horizontal alignment stated
    /// used to come back as <c>Left</c> because the setter rewrote it. It now stays
    /// <c>General</c>, which is what the file said.
    /// </summary>
    [Test]
    public async Task An_indent_alone_no_longer_forces_the_horizontal_alignment_to_left()
    {
        var alignment = new Alignment { Indent = 2U };

        var key = StyleDecoder.AlignmentKey(alignment, XLAlignmentValue.Default.Key);

        await Assert.That(key.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.General);
        await Assert.That(key.Indent).IsEqualTo(2);
    }

    /// <summary>
    /// The inline <c>&lt;numFmt&gt;</c> a dxf states. An id below
    /// <c>XLConstants.NumberOfBuiltInStyles</c> (164) wins over any format code present and clears
    /// the format — that exclusivity is not arbitrary, it is what the <c>IXLNumberFormat</c>
    /// setters enforced, and this overload reproduces it.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(InlineNumberFormatCases))]
    public async Task NumberFormatKey_decodes_an_inline_numFmt(
        NumberingFormat numberingFormat, int expectedId, string expectedFormat)
    {
        var key = StyleDecoder.NumberFormatKey(numberingFormat, XLNumberFormatValue.Default.Key);

        await Assert.That(key.NumberFormatId).IsEqualTo(expectedId);
        await Assert.That(key.Format).IsEqualTo(expectedFormat);
    }

    public static IEnumerable<Func<(NumberingFormat Format, int ExpectedId, string ExpectedFormat)>>
        InlineNumberFormatCases()
    {
        // A built-in id wins and the format code is discarded.
        yield return () => (new NumberingFormat { NumberFormatId = 5U, FormatCode = "0.00" },
            5, string.Empty);
        yield return () => (new NumberingFormat { NumberFormatId = 14U }, 14, string.Empty);

        // At and above 164 the format code wins and the key is marked custom (-1).
        yield return () => (new NumberingFormat { NumberFormatId = 164U, FormatCode = "0.000" },
            XLNumberFormatKey.CustomFormatNumberId, "0.000");
        yield return () => (new NumberingFormat { NumberFormatId = 200U, FormatCode = "#,##0" },
            XLNumberFormatKey.CustomFormatNumberId, "#,##0");

        // No id at all: the format code is all there is.
        yield return () => (new NumberingFormat { FormatCode = "yyyy-mm-dd" },
            XLNumberFormatKey.CustomFormatNumberId, "yyyy-mm-dd");
    }

    /// <summary>
    /// The id form of the resolver, over the workbook's declared custom formats. A declared id
    /// yields the custom key; anything else is treated as a built-in id.
    /// </summary>
    [Test]
    [Arguments(164, XLNumberFormatKey.CustomFormatNumberId, "0.000")]
    [Arguments(14, 14, "")]
    [Arguments(999, 999, "")]
    public async Task NumberFormatKey_resolves_a_numFmtId_against_the_declared_custom_formats(
        int numberFormatId, int expectedId, string expectedFormat)
    {
        var styles = new StylesheetData(
            Stylesheet: null,
            NumberingFormats: new NumberingFormats(
                new NumberingFormat { NumberFormatId = 164U, FormatCode = "0.000" }),
            Fills: null,
            Borders: null,
            Fonts: null,
            DifferentialFormats: new Dictionary<int, DifferentialFormat>());

        var key = StyleDecoder.NumberFormatKey(numberFormatId, styles,
            XLNumberFormatValue.Default.Key);

        await Assert.That(key.NumberFormatId).IsEqualTo(expectedId);
        await Assert.That(key.Format).IsEqualTo(expectedFormat);
    }

    /// <summary>
    /// The built-in branch clears the format string rather than inheriting it. The two decoders
    /// spec 28 unified disagreed here — the cell path left it inherited, the pivot path wrote the
    /// empty string — and the suite passes either way, so this pins the choice that was made
    /// rather than leaving it to be reversed by accident. The empty string is what a built-in id
    /// means: <c>NumberFormatId</c> is <c>-1</c> exactly when the format is custom, so any other id
    /// says the format lives in <c>XLPredefinedFormat</c> and no literal belongs beside it.
    /// </summary>
    [Test]
    public async Task A_built_in_numFmtId_clears_an_inherited_custom_format_string()
    {
        var styles = new StylesheetData(null, null, null, null, null,
            new Dictionary<int, DifferentialFormat>());
        var inherited = XLNumberFormatKey.ForFormat("0.000");

        var key = StyleDecoder.NumberFormatKey(14, styles, inherited);

        await Assert.That(key.NumberFormatId).IsEqualTo(14);
        await Assert.That(key.Format).IsEqualTo(string.Empty);
    }

    /// <summary>
    /// A <c>&lt;numFmt&gt;</c> declaring an id the workbook already declared keeps the first, which
    /// is what the linear scan this replaced did with its <c>FirstOrDefault</c>. Worth pinning
    /// because the per-load dictionary that also went away used <c>Add</c>, so such a file threw.
    /// </summary>
    [Test]
    public async Task A_duplicated_numFmtId_keeps_the_first_declaration_and_does_not_throw()
    {
        var styles = new StylesheetData(
            Stylesheet: null,
            NumberingFormats: new NumberingFormats(
                new NumberingFormat { NumberFormatId = 164U, FormatCode = "first" },
                new NumberingFormat { NumberFormatId = 164U, FormatCode = "second" }),
            Fills: null,
            Borders: null,
            Fonts: null,
            DifferentialFormats: new Dictionary<int, DifferentialFormat>());

        var key = StyleDecoder.NumberFormatKey(164, styles, XLNumberFormatValue.Default.Key);

        await Assert.That(key.Format).IsEqualTo("first");
    }

    /// <summary>
    /// A &lt;numFmt&gt; with an id but no format code is not admitted to the map, and the lookup
    /// falls through to the built-in branch by missing — the same place the scan reached by
    /// finding the element and then seeing its format code was empty.
    /// </summary>
    [Test]
    public async Task A_declared_numFmt_with_no_format_code_falls_through_to_the_built_in_branch()
    {
        var styles = new StylesheetData(
            Stylesheet: null,
            NumberingFormats: new NumberingFormats(new NumberingFormat { NumberFormatId = 200U }),
            Fills: null,
            Borders: null,
            Fonts: null,
            DifferentialFormats: new Dictionary<int, DifferentialFormat>());

        var key = StyleDecoder.NumberFormatKey(200, styles, XLNumberFormatValue.Default.Key);

        await Assert.That(key.NumberFormatId).IsEqualTo(200);
        await Assert.That(key.Format).IsEqualTo(string.Empty);
    }

    /// <summary>
    /// A &lt;border diagonalUp="1"/&gt; with no &lt;diagonal&gt; child decoded one way through
    /// BorderToXLibur (flags read only inside the &lt;diagonal&gt; guard) and another through
    /// LoadBorder (flags read unconditionally).
    /// <para>
    /// Spec 28 task 3 settled it towards the schema: ECMA-376 Part 1 §18.8.4 declares
    /// <c>diagonalUp</c> and <c>diagonalDown</c> as attributes of the <c>border</c> element, not
    /// as part of its <c>diagonal</c> child, so the unconditional read is correct and the guard was
    /// the bug. Only that one rule now exists, so the test asserts the decoded value against what
    /// the element stated rather than comparing two implementations.
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

        var key = StyleDecoder.BorderKey(border, XLBorderValue.Default.Key);

        await Assert.That(key.DiagonalUp).IsEqualTo(up);
        await Assert.That(key.DiagonalDown).IsEqualTo(down);
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
