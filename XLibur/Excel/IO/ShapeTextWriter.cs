using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes a shape's text model into its <c>xdr:txBody</c> at the grain the model changed: a run's
/// text into its <c>a:t</c>, a paragraph's alignment into its <c>a:pPr</c>, a new run or paragraph
/// after the ones the file has (spec 15, design rule 8). Only a body whose text was replaced
/// wholesale is rebuilt.
/// </summary>
/// <remarks>
/// A new run takes the run properties of the run before it in its paragraph, else of the body's
/// first run, else of the empty paragraph Excel writes for an empty body, else Excel's default for a
/// new shape: <c>lang="en-US" sz="1100"</c>. So text added to a loaded shape looks like the text
/// around it. Every edit goes through <see cref="TextBodyPrimitives"/>.
/// </remarks>
internal static class ShapeTextWriter
{
    /// <param name="body">The <c>xdr:txBody</c> to write into.</param>
    /// <param name="model">The text model.</param>
    /// <param name="alignNewParagraphs">
    /// Whether every new paragraph states its alignment, as Excel does in a preset shape, whose text
    /// a theme can otherwise centre. A text box states it only when it was assigned.
    /// </param>
    internal static void Write(Xdr.TextBody body, XLTextBody model, bool alignNewParagraphs)
    {
        WriteBodyProperties(body, model);

        var paragraphs = body.Elements<A.Paragraph>().ToList();
        var template = TemplateOf(paragraphs);
        if (model.IsRebuilt)
        {
            // The new paragraphs take the first one's properties, as the new runs take its first
            // run's, so a centred shape stays centred.
            var paragraphTemplate = paragraphs.FirstOrDefault()?.ParagraphProperties;
            foreach (var paragraph in paragraphs)
                paragraph.Remove();

            foreach (var paragraph in model.Paragraphs)
                body.Append(NewParagraph(paragraph, template, alignNewParagraphs, paragraphTemplate));

            if (!body.Elements<A.Paragraph>().Any())
                body.Append(EmptyParagraph(template, alignNewParagraphs, paragraphTemplate));

            return;
        }

        // The empty paragraph of an empty body takes the first paragraph added, so the body does not
        // keep a blank line in front of it.
        var placeholder = model.HasPlaceholderParagraph ? paragraphs.FirstOrDefault() : null;
        foreach (var paragraph in model.Paragraphs)
        {
            if (paragraph.DomIndex is { } index)
            {
                if (index < paragraphs.Count)
                    PatchParagraph(paragraphs[index], paragraph, template);

                continue;
            }

            if (placeholder is not null)
            {
                PatchParagraph(placeholder, paragraph, template);
                placeholder = null;
                continue;
            }

            body.Append(NewParagraph(paragraph, template, alignNewParagraphs));
        }

        EnsureParagraph(body, template, alignNewParagraphs);
    }

    /// <summary>
    /// Gives a body with no paragraph the empty one Excel writes: <c>CT_TextBody</c> needs at least
    /// one, and an <c>a:p</c> with a zero-length run is not how Excel says "no text".
    /// </summary>
    private static void EnsureParagraph(Xdr.TextBody body, A.RunProperties template, bool align)
    {
        if (!body.Elements<A.Paragraph>().Any())
            body.Append(EmptyParagraph(template, align));
    }

    /// <summary>The <c>a:bodyPr</c> attributes that were assigned: the vertical alignment and the wrap.</summary>
    private static void WriteBodyProperties(Xdr.TextBody body, XLTextBody model)
    {
        if (!model.IsVerticalAlignmentAssigned && !model.IsWrapTextAssigned)
            return;

        var properties = body.BodyProperties ?? body.PrependChild(new A.BodyProperties());
        if (model.IsVerticalAlignmentAssigned)
            properties.Anchor = ToAnchor(model.VerticalAlignment);

        if (model.IsWrapTextAssigned)
            properties.Wrap = model.WrapText ? A.TextWrappingValues.Square : A.TextWrappingValues.None;
    }

    private static void PatchParagraph(A.Paragraph element, XLTextParagraph paragraph, A.RunProperties template)
    {
        if (paragraph.IsAlignmentAssigned)
            TextBodyPrimitives.EnsureParagraphProperties(element).Alignment = ToAlignment(paragraph.Alignment);

        // Taken before any run is added, so each new run starts from the paragraph as the file had
        // it, and not from the formatting given to the new run before it.
        var runs = element.Elements<A.Run>().ToList();
        var paragraphTemplate = runs.LastOrDefault()?.RunProperties is { } previous
            ? (A.RunProperties)previous.CloneNode(true)
            : element.Elements<A.EndParagraphRunProperties>().FirstOrDefault() is { } end
                ? CopyCharacterProperties<A.RunProperties>(end)
                : template;

        foreach (var run in paragraph.Runs)
        {
            if (run.DomIndex is { } index)
            {
                if (index < runs.Count)
                    PatchRun(runs[index], run);

                continue;
            }

            TextBodyPrimitives.AppendRun(element, NewRun(run, (A.RunProperties)paragraphTemplate.CloneNode(true)));
        }
    }

    private static void PatchRun(A.Run element, XLTextRun run)
    {
        if ((run.Assigned & XLTextRunProperties.Text) != 0)
            TextBodyPrimitives.SetRunText(element, run.Text);

        if ((run.Assigned & ~XLTextRunProperties.Text) != 0)
            WriteFont(TextBodyPrimitives.EnsureRunProperties(element), run, run.Assigned);
    }

    private static A.Paragraph NewParagraph(XLTextParagraph paragraph, A.RunProperties template, bool align,
        A.ParagraphProperties? paragraphTemplate = null)
    {
        var element = new A.Paragraph();
        if (paragraphTemplate is not null)
        {
            var properties = (A.ParagraphProperties)paragraphTemplate.CloneNode(true);
            if (paragraph.IsAlignmentAssigned)
                properties.Alignment = ToAlignment(paragraph.Alignment);

            element.Append(properties);
        }
        else if (align || paragraph.IsAlignmentAssigned)
        {
            element.Append(new A.ParagraphProperties { Alignment = ToAlignment(paragraph.Alignment) });
        }

        foreach (var run in paragraph.Runs)
            element.Append(NewRun(run, (A.RunProperties)template.CloneNode(true)));

        if (paragraph.Runs.Count == 0)
            element.Append(CopyCharacterProperties<A.EndParagraphRunProperties>(template));

        return element;
    }

    /// <summary>The one paragraph of an empty body, as Excel writes it: no run, and an <c>a:endParaRPr</c>.</summary>
    private static A.Paragraph EmptyParagraph(A.RunProperties template, bool align,
        A.ParagraphProperties? paragraphTemplate = null)
    {
        var element = new A.Paragraph();
        if (paragraphTemplate is not null)
            element.Append(paragraphTemplate.CloneNode(true));
        else if (align)
            element.Append(new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Left });

        element.Append(CopyCharacterProperties<A.EndParagraphRunProperties>(template));
        return element;
    }

    /// <summary>A run of <paramref name="run"/>'s text in <paramref name="runProperties"/>, with every font property it was given.</summary>
    private static A.Run NewRun(XLTextRun run, A.RunProperties runProperties)
    {
        WriteFont(runProperties, run, run.Assigned);
        return TextBodyPrimitives.NewRun(runProperties, run.Text);
    }

    private static void WriteFont(A.RunProperties properties, XLTextRun run, XLTextRunProperties assigned)
    {
        if (Has(assigned, XLTextRunProperties.Bold))
            properties.Bold = run.Bold is { } bold ? bold : null;

        if (Has(assigned, XLTextRunProperties.Italic))
            properties.Italic = run.Italic is { } italic ? italic : null;

        if (Has(assigned, XLTextRunProperties.Underline))
            properties.Underline = run.Underline is { } underline ? ToUnderline(underline) : null;

        if (Has(assigned, XLTextRunProperties.Strikethrough))
            properties.Strike = run.Strikethrough is { } strike
                ? strike ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike
                : null;

        if (Has(assigned, XLTextRunProperties.FontSize))
            properties.FontSize = run.FontSize is { } size ? (int)System.Math.Round(size * 100) : null;

        if (Has(assigned, XLTextRunProperties.FontColor))
            TextBodyPrimitives.SetRunFill(properties, run.FontColor);

        if (Has(assigned, XLTextRunProperties.FontName))
            TextBodyPrimitives.SetLatinFont(properties, run.FontName);
    }

    /// <summary>The run properties new runs start from, when the run before them has none to give.</summary>
    private static A.RunProperties TemplateOf(List<A.Paragraph> paragraphs)
    {
        foreach (var paragraph in paragraphs)
        {
            if (paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties is { } first)
                return (A.RunProperties)first.CloneNode(true);
        }

        foreach (var paragraph in paragraphs)
        {
            if (paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault() is { } end)
                return CopyCharacterProperties<A.RunProperties>(end);
        }

        return new A.RunProperties { Language = "en-US", FontSize = 1100 };
    }

    /// <summary>
    /// A copy of a <c>CT_TextCharacterProperties</c> element under another name: <c>a:rPr</c> and
    /// <c>a:endParaRPr</c> are the one schema type.
    /// </summary>
    private static T CopyCharacterProperties<T>(OpenXmlCompositeElement source) where T : OpenXmlCompositeElement, new()
    {
        var copy = new T();
        copy.SetAttributes(source.GetAttributes());
        foreach (var child in source.ChildElements)
            copy.Append(child.CloneNode(true));

        return copy;
    }

    internal static A.TextAlignmentTypeValues ToAlignment(XLTextAlignment alignment) => alignment switch
    {
        XLTextAlignment.Center => A.TextAlignmentTypeValues.Center,
        XLTextAlignment.Right => A.TextAlignmentTypeValues.Right,
        XLTextAlignment.Justified => A.TextAlignmentTypeValues.Justified,
        XLTextAlignment.Distributed => A.TextAlignmentTypeValues.Distributed,
        _ => A.TextAlignmentTypeValues.Left,
    };

    internal static A.TextAnchoringTypeValues ToAnchor(XLTextVerticalAlignment alignment) => alignment switch
    {
        XLTextVerticalAlignment.Center => A.TextAnchoringTypeValues.Center,
        XLTextVerticalAlignment.Bottom => A.TextAnchoringTypeValues.Bottom,
        _ => A.TextAnchoringTypeValues.Top,
    };

    private static A.TextUnderlineValues ToUnderline(XLShapeTextUnderline underline) => underline switch
    {
        XLShapeTextUnderline.Single => A.TextUnderlineValues.Single,
        XLShapeTextUnderline.Double => A.TextUnderlineValues.Double,
        _ => A.TextUnderlineValues.None,
    };

    private static bool Has(XLTextRunProperties assigned, XLTextRunProperties any) => (assigned & any) != 0;
}
