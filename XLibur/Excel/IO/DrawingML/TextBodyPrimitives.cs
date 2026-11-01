using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Edits DrawingML text bodies, <c>a:CT_TextBody</c>: a chart title's <c>c:rich</c>, an extended
/// chart's <c>cx:rich</c> and a shape's <c>xdr:txBody</c>, which all hold the same paragraphs.
/// </summary>
/// <remarks>
/// <para>Three rules shape everything here, as <see cref="ShapePropertiesWriter"/>'s do for <c>spPr</c>:</para>
/// <list type="number">
/// <item>
/// <c>a:endParaRPr</c> closes a paragraph and stays last: a run goes in front of it, never after.
/// </item>
/// <item>
/// A run's text is rewritten in its <c>a:t</c>, and its <c>a:rPr</c> — a theme font, a language,
/// effects XLibur does not model — is left as it was.
/// </item>
/// <item>
/// <c>CT_TextCharacterProperties</c> and <c>CT_TextParagraph</c> are sequences, and the SDK does not
/// order children built by hand, so each insertion goes in front of the first child that must
/// follow it.
/// </item>
/// </list>
/// </remarks>
internal static class TextBodyPrimitives
{
    /// <summary>Whether an element is one of the children of <c>a:p</c> that carry text: a run, a line break or a field.</summary>
    internal static bool IsRunLevel(OpenXmlElement element) => element is A.Run or A.Break or A.Field;

    /// <summary>
    /// Replaces the text of a body with <paramref name="text"/> in a single paragraph and run, keeping
    /// the body's properties and the run properties of its first run.
    /// </summary>
    /// <remarks>A body spread over several paragraphs collapses into its first one.</remarks>
    internal static void SetSingleParagraphText(OpenXmlCompositeElement body, string text)
    {
        var paragraphs = body.Elements<A.Paragraph>().ToList();
        for (var i = 1; i < paragraphs.Count; i++)
            paragraphs[i].Remove();

        var paragraph = paragraphs.FirstOrDefault();
        if (paragraph == null)
        {
            paragraph = new A.Paragraph();
            body.Append(paragraph);
        }

        ReplaceParagraphText(paragraph, text);
    }

    /// <summary>
    /// Replaces every run, break and field of <paramref name="paragraph"/> with one run holding
    /// <paramref name="text"/>, in the run properties of the first run it had, or a bare
    /// <c>lang="en-US"</c> when it had none. The paragraph's own properties and <c>a:endParaRPr</c> stay.
    /// </summary>
    internal static void ReplaceParagraphText(A.Paragraph paragraph, string text)
    {
        var runProperties = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
        foreach (var existing in paragraph.ChildElements.Where(IsRunLevel).ToList())
            existing.Remove();

        AppendRun(paragraph, NewRun((A.RunProperties?)runProperties ?? new A.RunProperties { Language = "en-US" }, text));
    }

    /// <summary>A run of <paramref name="text"/> in <paramref name="runProperties"/>, which must not have a parent.</summary>
    internal static A.Run NewRun(A.RunProperties runProperties, string text) => new(runProperties, new A.Text(text));

    /// <summary>Puts <paramref name="run"/> after the paragraph's last run, break or field, ahead of <c>a:endParaRPr</c>.</summary>
    internal static void AppendRun(A.Paragraph paragraph, OpenXmlElement run)
    {
        var endProperties = paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
        if (endProperties != null)
            paragraph.InsertBefore(run, endProperties);
        else
            paragraph.Append(run);
    }

    /// <summary>Rewrites the text of <paramref name="run"/>, leaving its run properties alone.</summary>
    internal static void SetRunText(A.Run run, string text)
    {
        if (run.Text is { } existing)
            existing.Text = text;
        else
            run.Append(new A.Text(text));
    }

    /// <summary>The paragraph's <c>a:pPr</c>, added as its first child when it has none.</summary>
    internal static A.ParagraphProperties EnsureParagraphProperties(A.Paragraph paragraph) =>
        paragraph.ParagraphProperties ?? paragraph.PrependChild(new A.ParagraphProperties());

    /// <summary>The run's <c>a:rPr</c>, added as its first child when it has none.</summary>
    internal static A.RunProperties EnsureRunProperties(A.Run run) =>
        run.RunProperties ?? run.PrependChild(new A.RunProperties());

    /// <summary>
    /// Sets the run's fill to a solid <paramref name="color"/>, or removes the fill when it is
    /// <c>null</c>. A gradient or other fill goes too: the fills are one choice group.
    /// </summary>
    /// <remarks>In <c>CT_TextCharacterProperties</c> the fill follows only <c>a:ln</c>.</remarks>
    internal static void SetRunFill(A.RunProperties runProperties, XLColor? color)
    {
        foreach (var existing in runProperties.ChildElements.Where(IsFill).ToList())
            existing.Remove();

        if (color is not { HasValue: true })
            return;

        var fill = new A.SolidFill(ShapePropertiesWriter.BuildColor(color));
        if (runProperties.GetFirstChild<A.Outline>() is { } outline)
            runProperties.InsertAfter(fill, outline);
        else
            runProperties.PrependChild(fill);
    }

    /// <summary>Sets the run's Latin typeface, <c>a:latin</c>, or removes it when <paramref name="typeface"/> is <c>null</c>.</summary>
    internal static void SetLatinFont(A.RunProperties runProperties, string? typeface)
    {
        if (typeface is null)
        {
            runProperties.GetFirstChild<A.LatinFont>()?.Remove();
            return;
        }

        if (runProperties.GetFirstChild<A.LatinFont>() is { } latin)
        {
            latin.Typeface = typeface;
            return;
        }

        var font = new A.LatinFont { Typeface = typeface };
        var after = runProperties.ChildElements.FirstOrDefault(IsAfterLatin);
        if (after != null)
            runProperties.InsertBefore(font, after);
        else
            runProperties.Append(font);
    }

    private static bool IsFill(OpenXmlElement element) =>
        element is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    /// <summary>The children of <c>CT_TextCharacterProperties</c> that follow <c>a:latin</c>.</summary>
    private static bool IsAfterLatin(OpenXmlElement element) =>
        element is A.EastAsianFont or A.ComplexScriptFont or A.SymbolFont or A.HyperlinkOnClick
            or A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList;
}
