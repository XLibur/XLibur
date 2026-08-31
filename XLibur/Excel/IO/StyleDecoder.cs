using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Extensions;
using XLibur.Utils;

// S4136 wants the three Decode overloads adjacent. They are ordered by layer instead: the two
// entry points a caller reaches from a style index come first, then the element decoders those
// call. Making the overloads adjacent would move an entry point below the machinery it delegates
// to, for an ordering nothing else in the file follows.
#pragma warning disable S4136

namespace XLibur.Excel.IO;

/// <summary>
/// The single decoder from OOXML style XML to XLibur style keys.
/// </summary>
/// <remarks>
/// Before spec 28 the same XML was decoded by two families chosen by provenance: a mutating one for
/// <c>&lt;dxfs&gt;</c> that wrote through <c>IXLFontBase</c> and friends, and a key-returning one for
/// <c>&lt;cellXfs&gt;</c>. They had diverged — a dxf font lost its name, family and charset, and the
/// diagonal border flags were read under different conditions. One implementation cannot diverge
/// from itself.
/// </remarks>
internal static class StyleDecoder
{
    /// <summary>
    /// Decodes the <c>&lt;xf&gt;</c> at <paramref name="styleIndex"/> in <c>&lt;cellXfs&gt;</c>.
    /// A workbook with no stylesheet, or one whose stylesheet declares no cell formats, leaves
    /// <paramref name="defaults"/> untouched.
    /// </summary>
    internal static XLStyleKey Decode(int styleIndex, StylesheetData styles, XLStyleKey defaults)
    {
        if (styles.Stylesheet is not { CellFormats: not null } s)
            return defaults; // No stylesheet, no styles.

        return Decode((CellFormat)s.CellFormats.ElementAt(styleIndex), styles, defaults);
    }

    /// <summary>
    /// Resolves a <c>&lt;cellXfs&gt;</c> style index to an interned <see cref="XLStyleValue"/>
    /// without creating an <see cref="XLStyle"/> wrapper or writing to any slice.
    /// </summary>
    internal static XLStyleValue ResolveStyleValue(int styleIndex, StylesheetData styles)
    {
        var key = Decode(styleIndex, styles, XLStyle.Default.Key);
        return XLStyleValue.FromKey(ref key);
    }

    /// <summary>
    /// Decodes the <c>&lt;xf&gt;</c> at <paramref name="styleIndex"/> and applies it to
    /// <paramref name="xlStylized"/>.
    /// </summary>
    internal static void ApplyStyle(IXLStylized xlStylized, int styleIndex, StylesheetData styles)
    {
        var xlStyleKey = Decode(styleIndex, styles, XLStyle.Default.Key);

        // When loading columns, we must propagate the style to each column but not deeper. In other cases we do not propagate at all.
        if (xlStylized is IXLColumns columns)
        {
            columns.Cast<XLColumn>().ForEach(col => col.InnerStyle = new XLStyle(col, xlStyleKey));
        }
        else
        {
            xlStylized.InnerStyle = new XLStyle(xlStylized, xlStyleKey);
        }
    }

    /// <summary>
    /// Decodes one <c>&lt;xf&gt;</c> from <c>&lt;cellXfs&gt;</c>. Each aspect is decoded only when
    /// the <c>&lt;xf&gt;</c> states an index or a child for it, so an unstated aspect keeps
    /// whatever <paramref name="defaults"/> carried.
    /// </summary>
    internal static XLStyleKey Decode(CellFormat cellFormat, StylesheetData styles, XLStyleKey defaults)
    {
        var key = defaults with
        {
            IncludeQuotePrefix = OpenXmlHelper.GetBooleanValueAsBool(cellFormat.QuotePrefix, false),
        };

        if (cellFormat.ApplyProtection != null)
        {
            var protection = cellFormat.Protection;
            var protectionKey = XLProtectionValue.Default.Key;
            if (protection is not null)
                protectionKey = ProtectionKey(protection, protectionKey);

            key = key with { Protection = protectionKey };
        }

        if (UInt32HasValue(cellFormat.FillId))
        {
            var fill = (Fill)styles.Fills!.ElementAt((int)cellFormat.FillId!.Value);

            // Unlike the other aspects, the fill does not inherit from what the key already holds:
            // a cellXf's fillId points at a complete <fill> definition rather than an override, so
            // it is decoded against the default fill. Protection below is the same shape. This is
            // what the decoder replaced here did, by mutating a fresh XLFill.
            if (fill.PatternFill is not null)
                key = key with { Fill = FillKey(fill, differential: false, XLFillValue.Default.Key) };
        }

        if (cellFormat.Alignment is { } alignment)
            key = key with { Alignment = AlignmentKey(alignment, key.Alignment) };

        if (UInt32HasValue(cellFormat.BorderId))
        {
            var border = (Border)styles.Borders!.ElementAt((int)cellFormat.BorderId!.Value);
            key = key with { Border = BorderKey(border, key.Border) };
        }

        if (UInt32HasValue(cellFormat.FontId))
        {
            var font = (Font)styles.Fonts!.ElementAt((int)cellFormat.FontId!.Value);
            key = key with { Font = FontKey(font, key.Font) };
        }

        if (UInt32HasValue(cellFormat.NumberFormatId))
        {
            key = key with
            {
                NumberFormat = NumberFormatKey((int)cellFormat.NumberFormatId!.Value, styles,
                    key.NumberFormat),
            };
        }

        return key;
    }

    /// <summary>
    /// Decodes one <c>&lt;dxf&gt;</c>. Differential formats state only what they override, so every
    /// absent child leaves the corresponding part of <paramref name="defaults"/> in place.
    /// </summary>
    /// <remarks>
    /// Reads all six children <c>CT_Dxf</c> permits. Before spec 28 the three callers read three
    /// different subsets of them — the conditional-format reader four, the pivot reader five, the
    /// writer's reuse map four — and none read <c>&lt;protection&gt;</c> at all.
    /// </remarks>
    internal static XLStyleKey Decode(DifferentialFormat dxf, XLStyleKey defaults)
    {
        var key = defaults;

        if (dxf.Font is { } font)
            key = key with { Font = FontKey(font, key.Font) };

        if (dxf.Fill is { } fill)
            key = key with { Fill = FillKey(fill, differential: true, key.Fill) };

        if (dxf.Border is { } border)
            key = key with { Border = BorderKey(border, key.Border) };

        if (dxf.NumberingFormat is { } numberingFormat)
            key = key with { NumberFormat = NumberFormatKey(numberingFormat, key.NumberFormat) };

        if (dxf.Alignment is { } alignment)
            key = key with { Alignment = AlignmentKey(alignment, key.Alignment) };

        if (dxf.Protection is { } protection)
            key = key with { Protection = ProtectionKey(protection, key.Protection) };

        return key;
    }

    /// <summary>
    /// Decodes an <c>&lt;alignment&gt;</c>. Every attribute the element omits falls back to
    /// <paramref name="defaultAlignment"/>.
    /// </summary>
    /// <remarks>
    /// Decoding to a key rather than writing through <c>IXLAlignment</c> is what lets an indent
    /// load as written. That interface's <c>Indent</c> setter rewrites a <c>General</c> horizontal
    /// alignment to <c>Left</c> and throws for an indent on a centred one — a reasonable guard on
    /// the public API, and the wrong rule for a reader reproducing what a file states.
    /// </remarks>
    internal static XLAlignmentKey AlignmentKey(Alignment alignment, XLAlignmentKey defaultAlignment)
    {
        return new XLAlignmentKey
        {
            Indent = checked((int?)alignment.Indent?.Value) ?? defaultAlignment.Indent,
            Horizontal = alignment.Horizontal.ToXLiburOrNull() ?? defaultAlignment.Horizontal,
            Vertical = alignment.Vertical.ToXLiburOrNull() ?? defaultAlignment.Vertical,
            ReadingOrder = alignment.ReadingOrder?.Value.ToXLibur() ?? defaultAlignment.ReadingOrder,
            WrapText = alignment.WrapText?.Value ?? defaultAlignment.WrapText,
            TextRotation = alignment.TextRotation is not null
                ? OpenXmlHelper.GetXLiburTextRotation(alignment)
                : defaultAlignment.TextRotation,
            ShrinkToFit = alignment.ShrinkToFit?.Value ?? defaultAlignment.ShrinkToFit,
            RelativeIndent = alignment.RelativeIndent?.Value ?? defaultAlignment.RelativeIndent,
            JustifyLastLine = alignment.JustifyLastLine?.Value ?? defaultAlignment.JustifyLastLine,
        };
    }

    /// <summary>
    /// Decodes a <c>&lt;border&gt;</c>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>The diagonal flags are read unconditionally, independent of the <c>&lt;diagonal&gt;</c>
    /// child.</b> ECMA-376 Part 1, §18.8.4 (<c>border</c>, <c>CT_Border</c>) declares
    /// <c>diagonalUp</c> and <c>diagonalDown</c> as <em>attributes of the <c>border</c> element</em>,
    /// alongside <c>outline</c>, while <c>diagonal</c> is one of the nine <c>CT_BorderPr</c> child
    /// elements in the type's sequence. An attribute of <c>border</c> is a sibling of the
    /// <c>diagonal</c> child, not a dependent of it, so nothing in the schema makes reading the
    /// flags conditional on the child being present.
    /// </para>
    /// <para>
    /// Before spec 28 the two decoders disagreed here: the key form read the flags only inside the
    /// <c>&lt;diagonal&gt;</c> guard and the mutating form read them unconditionally, so a
    /// <c>&lt;border diagonalUp="1"/&gt;</c> with no <c>&lt;diagonal&gt;</c> child produced two
    /// different keys for one element. The unconditional read is the one that matches the schema,
    /// so that is the single rule implemented here; the guarded read is gone rather than kept
    /// behind a flag.
    /// </para>
    /// </remarks>
    internal static XLBorderKey BorderKey(Border b, XLBorderKey defaultBorder)
    {
        var nb = defaultBorder;

        if (b.DiagonalBorder is { } diagonalBorder)
            nb = ApplyBorderStyleAndColor(nb, diagonalBorder,
                (key, style) => key with { DiagonalBorder = style },
                (key, color) => key with { DiagonalBorderColor = color });

        if (b.DiagonalUp is not null)
            nb = nb with { DiagonalUp = b.DiagonalUp.Value };
        if (b.DiagonalDown is not null)
            nb = nb with { DiagonalDown = b.DiagonalDown.Value };

        if (b.LeftBorder is not null)
            nb = ApplyBorderStyleAndColor(nb, b.LeftBorder,
                (key, style) => key with { LeftBorder = style },
                (key, color) => key with { LeftBorderColor = color });

        if (b.RightBorder is not null)
            nb = ApplyBorderStyleAndColor(nb, b.RightBorder,
                (key, style) => key with { RightBorder = style },
                (key, color) => key with { RightBorderColor = color });

        if (b.TopBorder is not null)
            nb = ApplyBorderStyleAndColor(nb, b.TopBorder,
                (key, style) => key with { TopBorder = style },
                (key, color) => key with { TopBorderColor = color });

        if (b.BottomBorder is not null)
            nb = ApplyBorderStyleAndColor(nb, b.BottomBorder,
                (key, style) => key with { BottomBorder = style },
                (key, color) => key with { BottomBorderColor = color });

        // A file is free to state a colour for an edge it gives no style - the two attributes are
        // independent in the schema - so normalize on the way in. Otherwise such a key would compare
        // unequal to the interned form of the same border, and BordersAreEqual would write a
        // duplicate <border> for one already in the stylesheet.
        return nb.Normalize();
    }

    /// <summary>
    /// Applies one edge's <c>style</c> and <c>color</c> to the key, each only if the edge states
    /// it. The two are independent attributes, which is why <see cref="XLBorderKey.Normalize"/> is
    /// needed afterwards.
    /// </summary>
    private static XLBorderKey ApplyBorderStyleAndColor(
        XLBorderKey nb,
        BorderPropertiesType border,
        Func<XLBorderKey, XLBorderStyleValues, XLBorderKey> applyStyle,
        Func<XLBorderKey, XLColorKey, XLBorderKey> applyColor)
    {
        if (border.Style is not null)
            nb = applyStyle(nb, border.Style.Value.ToXLibur());
        if (border.Color is not null)
            nb = applyColor(nb, border.Color.ToXLiburColor().Key);
        return nb;
    }

    /// <summary>
    /// Decodes a <c>&lt;fill&gt;</c>.
    /// </summary>
    /// <param name="fill">The fill element.</param>
    /// <param name="differential">
    /// Differential fills store background in <c>bgColor</c> and pattern in <c>fgColor</c>, which is
    /// the sane reading. Ordinary fills store the background in <c>fgColor</c> when the pattern is
    /// solid. The flag selects between them; it is not a style choice.
    /// </param>
    /// <param name="defaults">The fill to fall back to for anything the element does not state.</param>
    internal static XLFillKey FillKey(Fill fill, bool differential, XLFillKey defaults)
    {
        if (fill.PatternFill is not { } patternFill)
            return defaults;

        var patternType = patternFill.PatternType is not null
            ? patternFill.PatternType.Value.ToXLibur()
            : XLFillPatternValues.Solid;

        var key = defaults with { PatternType = patternType };

        // The None branch touches no colour while the other two default a missing background to
        // index 64 (transparent). That asymmetry is carried over from the decoder this replaced.
        return patternType switch
        {
            XLFillPatternValues.None => key,
            XLFillPatternValues.Solid => key with
            {
                BackgroundColor = SolidFillBackground(patternFill, differential),
            },
            _ => key with
            {
                PatternColor = patternFill.ForegroundColor is not null
                    ? patternFill.ForegroundColor.ToXLiburColor().Key
                    : key.PatternColor,
                BackgroundColor = patternFill.BackgroundColor is not null
                    ? patternFill.BackgroundColor.ToXLiburColor().Key
                    : XLColor.FromIndex(64).Key,
            },
        };
    }

    /// <summary>
    /// The background of a solid fill, which an ordinary fill stores in <c>fgColor</c> and a
    /// differential one in <c>bgColor</c>. A fill stating neither is transparent (index 64).
    /// </summary>
    private static XLColorKey SolidFillBackground(PatternFill patternFill, bool differential)
    {
        // yes, for a non-differential solid fill the source is the foreground!
        ColorType? source = differential ? patternFill.BackgroundColor : patternFill.ForegroundColor;
        return source is not null ? source.ToXLiburColor().Key : XLColor.FromIndex(64).Key;
    }

    /// <summary>
    /// Decodes a <c>&lt;font&gt;</c> through the typed <c>CT_Font</c> children — including
    /// <c>name</c>, <c>family</c> and <c>charset</c>, which the decoder this replaced could not
    /// read off a font at all because it looked for the <c>&lt;x:rPr&gt;</c> spellings.
    /// </summary>
    /// <remarks>
    /// Bold, italic, shadow and strikethrough are assigned unconditionally, so a font element that
    /// omits them decodes to <c>false</c> rather than inheriting from <paramref name="nf"/>. Both
    /// decoder families behaved this way before spec 28, and it is harmless for a
    /// <c>&lt;dxf&gt;</c> because a dxf is never decoded against a cell's font — every caller
    /// passes a default-derived key, where those four are already <c>false</c>. See
    /// <c>A_colour_only_conditional_format_leaves_a_bold_cell_bold</c>.
    /// </remarks>
    internal static XLFontKey FontKey(Font f, XLFontKey nf)
    {
        nf = nf with
        {
            Bold = OpenXmlHelper.GetBoolean(f.Bold),
            Italic = OpenXmlHelper.GetBoolean(f.Italic),
            Shadow = OpenXmlHelper.GetBoolean(f.Shadow),
            Strikethrough = OpenXmlHelper.GetBoolean(f.Strike),
        };

        var underline = f.Underline;
        if (underline is not null)
        {
            var value = underline.Val?.Value.ToXLibur() ??
                        XLFontUnderlineValues.Single;
            nf = nf with { Underline = value };
        }

        var verticalTextAlignment = f.VerticalTextAlignment;
        if (verticalTextAlignment is not null)
        {
            var value = verticalTextAlignment.Val?.Value.ToXLibur() ??
                        XLFontVerticalTextAlignmentValues.Baseline;
            nf = nf with { VerticalAlignment = value };
        }

        var fontSize = f.FontSize?.Val;
        if (fontSize is not null)
            nf = nf with { FontSize = fontSize.Value };

        var color = f.Color;
        if (color is not null)
            nf = nf with { FontColor = color.ToXLiburColor().Key };

        var fontName = f.FontName?.Val?.Value ?? string.Empty;
        if (!string.IsNullOrEmpty(fontName))
            nf = nf with { FontName = fontName };

        var fontFamilyNumbering = f.FontFamilyNumbering?.Val?.Value;
        if (fontFamilyNumbering is not null)
            nf = nf with { FontFamilyNumbering = (XLFontFamilyNumberingValues)fontFamilyNumbering };

        var fontCharSet = f.FontCharSet?.Val?.Value;
        if (fontCharSet is not null)
            nf = nf with { FontCharSet = (XLFontCharSet)fontCharSet };

        var fontScheme = f.FontScheme;
        if (fontScheme is not null)
            nf = nf with { FontScheme = fontScheme.Val?.Value.ToXLibur() ?? XLFontScheme.None };
        return nf;
    }

    /// <summary>
    /// Decodes a <c>&lt;protection&gt;</c>. Before spec 28 no <c>&lt;dxf&gt;</c> caller read this
    /// element at all, so a conditional or pivot format's protection was dropped on load.
    /// </summary>
    internal static XLProtectionKey ProtectionKey(Protection protection, XLProtectionKey p)
    {
        // OI29500, hidden default is false, locked default is true.
        if (protection.Hidden is not null)
            p = p with { Hidden = protection.Hidden.Value };

        if (protection.Locked is not null)
            p = p with { Locked = protection.Locked.Value };

        return p;
    }

    /// <summary>
    /// Resolves a <c>numFmtId</c>. A workbook-declared custom format wins; anything else is taken to
    /// be a built-in id, including an id at or above <see cref="XLConstants.NumberOfBuiltInStyles"/>
    /// that no <c>&lt;numFmt&gt;</c> declares — such a file is malformed, and this is what both the
    /// cell and pivot paths did before spec 28 unified them.
    /// </summary>
    /// <remarks>
    /// The two paths this replaces disagreed about the format string on the built-in branch: the
    /// cell path left it inherited, the pivot path wrote the empty string. This takes the empty
    /// string, which is what a built-in id means — <see cref="XLNumberFormatKey.NumberFormatId"/>
    /// is <c>-1</c> exactly when the format is custom, so any other id says the format lives in
    /// <c>XLPredefinedFormat</c> and no literal belongs in the key beside it.
    /// </remarks>
    internal static XLNumberFormatKey NumberFormatKey(int numberFormatId, StylesheetData styles,
        XLNumberFormatKey defaults)
    {
        if (styles.CustomNumberFormats.TryGetValue(numberFormatId, out var formatCode))
            return XLNumberFormatKey.ForFormat(formatCode);

        return defaults with { NumberFormatId = numberFormatId, Format = string.Empty };
    }

    /// <summary>
    /// Reads a <c>&lt;numFmt&gt;</c> stated inline, as a dxf states it.
    /// </summary>
    /// <remarks>
    /// The id branch clearing <see cref="XLNumberFormatKey.Format"/> is not a change of behaviour:
    /// <c>IXLNumberFormat.NumberFormatId</c>'s setter resets <c>Format</c> to
    /// <c>XLNumberFormatValue.Default.Format</c> (the empty string), so the mutating decoder this
    /// replaced already produced exactly this key.
    /// </remarks>
    internal static XLNumberFormatKey NumberFormatKey(NumberingFormat inline, XLNumberFormatKey defaults)
    {
        if (inline.NumberFormatId is { Value: var id } && id < XLConstants.NumberOfBuiltInStyles)
            return new XLNumberFormatKey { NumberFormatId = (int)id, Format = string.Empty };

        if (inline.FormatCode?.Value is { Length: > 0 } code)
            return XLNumberFormatKey.ForFormat(code);

        return defaults;
    }

    /// <summary>
    /// Decodes a rich-text run's <c>&lt;x:rPr&gt;</c>. Separate from <see cref="FontKey"/> on
    /// purpose: <c>CT_RPrElt</c> and <c>CT_Font</c> spell three children with different CLR types
    /// (<c>rFont</c>/<c>name</c>, and two each for <c>family</c> and <c>charset</c>), so one
    /// element-typed decoder cannot serve both. Conflating them is what dropped three fields from
    /// every dxf font before spec 28.
    /// </summary>
    /// <remarks>
    /// This reads one field more than the decoder it replaces: <c>&lt;charset&gt;</c>
    /// (<c>RunPropertyCharSet</c>) was never looked for on the rich-text path either, for the same
    /// reason it was dropped on the dxf path — an untyped <c>OpenXmlElement</c> let one function
    /// pretend to serve two schemas.
    /// </remarks>
    internal static XLFontKey RunFontKey(RunProperties runProperties, XLFontKey nf)
    {
        nf = nf with
        {
            Bold = OpenXmlHelper.GetBoolean(runProperties.Elements<Bold>().FirstOrDefault()),
            Italic = OpenXmlHelper.GetBoolean(runProperties.Elements<Italic>().FirstOrDefault()),
            Shadow = OpenXmlHelper.GetBoolean(runProperties.Elements<Shadow>().FirstOrDefault()),
            Strikethrough = OpenXmlHelper.GetBoolean(runProperties.Elements<Strike>().FirstOrDefault()),
        };

        var fontColor = runProperties.Elements<Color>().FirstOrDefault();
        if (fontColor is not null)
            nf = nf with { FontColor = fontColor.ToXLiburColor().Key };

        var fontFamily = runProperties.Elements<FontFamily>().FirstOrDefault();
        if (fontFamily?.Val is not null)
            nf = nf with { FontFamilyNumbering = (XLFontFamilyNumberingValues)fontFamily.Val.Value };

        var runFont = runProperties.Elements<RunFont>().FirstOrDefault();
        if (runFont?.Val?.Value is { } runFontName)
            nf = nf with { FontName = runFontName };

        var charSet = runProperties.Elements<RunPropertyCharSet>().FirstOrDefault();
        if (charSet?.Val is not null)
            nf = nf with { FontCharSet = (XLFontCharSet)charSet.Val.Value };

        var fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
        if (fontSize?.Val is not null)
            nf = nf with { FontSize = fontSize.Val.Value };

        var underline = runProperties.Elements<Underline>().FirstOrDefault();
        if (underline is not null)
        {
            nf = nf with
            {
                Underline = underline.Val is not null
                    ? underline.Val.Value.ToXLibur()
                    : XLFontUnderlineValues.Single,
            };
        }

        var verticalTextAlignment = runProperties.Elements<VerticalTextAlignment>().FirstOrDefault();
        if (verticalTextAlignment is not null)
        {
            nf = nf with
            {
                VerticalAlignment = verticalTextAlignment.Val is not null
                    ? verticalTextAlignment.Val.Value.ToXLibur()
                    : XLFontVerticalTextAlignmentValues.Baseline,
            };
        }

        var fontScheme = runProperties.Elements<FontScheme>().FirstOrDefault();
        if (fontScheme is not null)
        {
            nf = nf with
            {
                FontScheme = fontScheme.Val is not null
                    ? fontScheme.Val.Value.ToXLibur()
                    : XLFontScheme.None,
            };
        }

        return nf;
    }

    /// <summary>
    /// Decodes a rich-text run's <c>&lt;x:rPr&gt;</c> and writes it through an
    /// <see cref="IXLFontBase"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A thin applier over <see cref="RunFontKey"/> rather than a second decoder. The rich-text
    /// call sites hold an <c>XLRichString</c>, which is an <see cref="IXLFontBase"/> with no
    /// <c>InnerStyle</c> to assign a key to, so the fields are written through one at a time. One
    /// decode, one shape.
    /// </para>
    /// <para>
    /// The writes are gated and ordered exactly as the decoder this replaced gated and ordered
    /// them, and that is load-bearing rather than cosmetic. A rich run is part of its shared
    /// string's identity, so every property write on one dereferences that string's
    /// shared-string-table entry and interns a new one. Since
    /// <c>SharedStringTable.GetConsecutiveMap</c> emits entries in insertion order, a different
    /// set or order of intermediate writes reorders <c>sharedStrings.xml</c> for a file whose
    /// content is unchanged. The four booleans are written unconditionally, as before; everything
    /// else is written only when the run states the element.
    /// </para>
    /// </remarks>
    internal static void ApplyRunFont(RunProperties? runProperties, IXLFontBase fontBase)
    {
        if (runProperties is null)
            return;

        var key = RunFontKey(runProperties, XLFont.GenerateKey(fontBase));

        fontBase.Bold = key.Bold;

        if (runProperties.Elements<Color>().Any())
        {
            var fontColor = key.FontColor;
            fontBase.FontColor = XLColor.FromKey(ref fontColor);
        }

        if (runProperties.Elements<FontFamily>().Any(f => f.Val is not null))
            fontBase.FontFamilyNumbering = key.FontFamilyNumbering;

        if (runProperties.Elements<RunFont>().Any(f => f.Val is not null))
            fontBase.FontName = key.FontName;

        if (runProperties.Elements<FontSize>().Any(f => f.Val is not null))
            fontBase.FontSize = key.FontSize;

        fontBase.Italic = key.Italic;
        fontBase.Shadow = key.Shadow;
        fontBase.Strikethrough = key.Strikethrough;

        if (runProperties.Elements<Underline>().Any())
            fontBase.Underline = key.Underline;

        if (runProperties.Elements<VerticalTextAlignment>().Any())
            fontBase.VerticalAlignment = key.VerticalAlignment;

        if (runProperties.Elements<FontScheme>().Any())
            fontBase.FontScheme = key.FontScheme;

        // New in spec 28: the charset was never read on this path. It goes last so that adding it
        // cannot disturb the intermediate states the writes above produce for a run that has none.
        if (runProperties.Elements<RunPropertyCharSet>().Any(c => c.Val is not null))
            fontBase.FontCharSet = key.FontCharSet;
    }

    /// <summary>
    /// Whether a <c>&lt;cellXf&gt;</c> index attribute is both present and carries a value. Each
    /// such guard suppresses a decode, so an absent index leaves that aspect at its default.
    /// </summary>
    private static bool UInt32HasValue(UInt32Value? value)
    {
        return value != null && value.HasValue;
    }
}
