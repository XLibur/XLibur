using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// The title of a chart — <c>c:title</c> on a standard chart, <c>cx:title</c> on an extended one —
/// and the rich text block both of them carry the literal text in.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Apply"/> covers both a chart being created and a chart loaded from a file. Creating the
/// element is the branch it takes when there is none, so a new chart and a loaded chart reach the
/// same XML from the same model rather than through two functions that had to agree by hand.
/// </para>
/// <para>
/// Only the text is ever rewritten. An existing title keeps its layout, overlay, shape and text
/// properties, and the run properties of the run that held the old text, so a chart that came out of
/// a file with a styled title keeps that styling.
/// </para>
/// </remarks>
internal static class ChartTitleXml
{
    /// <summary>
    /// Writes an assigned chart title into <paramref name="chart"/>, adding, editing or removing the
    /// <c>c:title</c> child as the model requires. A chart nobody retitled is not modified.
    /// </summary>
    /// <remarks>
    /// <c>c:autoTitleDeleted</c> is kept in step, because Excel hides the title outright while it is
    /// set and would otherwise ignore whatever was written here.
    /// </remarks>
    internal static void Apply(C.Chart chart, XLChart xlChart)
    {
        if (!xlChart.TitleAssigned)
            return;

        var title = chart.Elements<C.Title>().FirstOrDefault();

        if (xlChart.Title == null)
        {
            title?.Remove();
            SetAutoTitleDeleted(chart, deleted: true);
            return;
        }

        if (title == null)
        {
            title = new C.Title(LiteralText(xlChart.Title), new C.Overlay { Val = false });
            ChartElementOrder.InsertOrdered(chart, title, ChartElementOrder.ChartChildOrder);
        }
        else
        {
            SetTitleText(title, xlChart.Title);
        }

        SetAutoTitleDeleted(chart, deleted: false);
    }

    /// <summary>
    /// Writes an assigned title into the <c>cx:chart</c> of an extended chart part.
    /// </summary>
    /// <remarks>
    /// An extended chart has none of the other properties the patchers write — its series live in the
    /// cx namespace and carry no formatting XLibur models — so the title is the one edit that can be
    /// carried back into one. As with a standard chart, only the text is replaced, and there is no
    /// <c>autoTitleDeleted</c> to keep in step: removing <c>cx:title</c> is all a null title needs.
    /// </remarks>
    internal static void ApplyExtended(Cx.Chart chart, XLChart xlChart)
    {
        if (!xlChart.TitleAssigned)
            return;

        var title = chart.Elements<Cx.ChartTitle>().FirstOrDefault();

        if (xlChart.Title == null)
        {
            title?.Remove();
            return;
        }

        if (title == null)
        {
            // cx:title opens CT_Chart, ahead of cx:plotArea and cx:legend.
            chart.InsertAt(ExtendedTitle(xlChart.Title), 0);
            return;
        }

        SetExtendedTitleText(title, xlChart.Title);
    }

    /// <summary>
    /// The <c>c:tx</c> of a chart or axis title: a rich text run holding the literal text. An axis
    /// title is the same block under a different parent, which is why this is shared rather than
    /// written twice.
    /// </summary>
    internal static C.ChartText LiteralText(string text) =>
        new(new C.RichText(new A.BodyProperties(), new A.ListStyle(), TitleParagraph(text)));

    private static void SetAutoTitleDeleted(C.Chart chart, bool deleted)
    {
        var existing = chart.Elements<C.AutoTitleDeleted>().FirstOrDefault();
        if (existing != null)
        {
            existing.Val = deleted;
            return;
        }

        ChartElementOrder.InsertOrdered(chart, new C.AutoTitleDeleted { Val = deleted },
            ChartElementOrder.ChartChildOrder);
    }

    /// <summary>
    /// Puts <paramref name="text"/> into an existing <c>c:title</c>, reusing its rich text block when
    /// it has one.
    /// </summary>
    private static void SetTitleText(C.Title title, string text)
    {
        var chartText = title.Elements<C.ChartText>().FirstOrDefault();
        var rich = chartText?.Elements<C.RichText>().FirstOrDefault();
        if (rich != null)
        {
            SetRichText(rich, text);
            return;
        }

        // Either the title carried no text at all, or its text came from a cell through a
        // c:strRef — which a literal title replaces rather than edits.
        chartText?.Remove();
        ChartElementOrder.InsertOrdered(title, LiteralText(text), ChartElementOrder.TitleChildOrder);
    }

    /// <summary>
    /// Replaces the text of a rich text block with a single run, keeping the block's body and list
    /// properties and the run properties of the run that was there.
    /// </summary>
    /// <param name="rich">
    /// A <c>c:rich</c> or its extended-chart counterpart <c>cx:rich</c>. The two are different
    /// elements holding the same DrawingML paragraphs.
    /// </param>
    /// <param name="text">The text the block is left holding.</param>
    private static void SetRichText(OpenXmlCompositeElement rich, string text)
    {
        var paragraphs = rich.Elements<A.Paragraph>().ToList();

        // A title spread over several paragraphs collapses into one; the formatting that survives is
        // the first run's, which is the run the caller was looking at.
        for (var i = 1; i < paragraphs.Count; i++)
            paragraphs[i].Remove();

        var paragraph = paragraphs.FirstOrDefault();
        if (paragraph == null)
        {
            paragraph = new A.Paragraph();
            rich.Append(paragraph);
        }

        var runProperties = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
        foreach (var existing in paragraph.ChildElements.Where(IsRunLevel).ToList())
            existing.Remove();

        var run = new A.Run(
            runProperties ?? new A.RunProperties { Language = "en-US" },
            new A.Text(text));

        // a:endParaRPr closes the paragraph and has to stay last.
        var endProperties = paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
        if (endProperties != null)
            paragraph.InsertBefore(run, endProperties);
        else
            paragraph.Append(run);
    }

    /// <summary>Whether an element is one of the run-level children of <c>a:p</c> that carry text.</summary>
    private static bool IsRunLevel(OpenXmlElement element) => element is A.Run or A.Break or A.Field;

    private static A.Paragraph TitleParagraph(string text) =>
        new(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text)));

    private static void SetExtendedTitleText(Cx.ChartTitle title, string text)
    {
        var chartText = title.Elements<Cx.Text>().FirstOrDefault();

        var rich = chartText?.Elements<Cx.RichTextBody>().FirstOrDefault();
        if (rich != null)
        {
            SetRichText(rich, text);
            return;
        }

        // A cx:tx may hold plain text in a cx:txData/cx:v instead of a rich body.
        var value = chartText?.Elements<Cx.TextData>().FirstOrDefault()
            ?.Elements<Cx.VXsdstring>().FirstOrDefault();
        if (value != null)
        {
            value.Text = text;
            return;
        }

        chartText?.Remove();
        title.InsertAt(ExtendedTitleText(text), 0);
    }

    /// <summary>
    /// The <c>cx:title</c> of an extended chart: centred above the plot, holding a rich text run with
    /// the literal text.
    /// </summary>
    private static Cx.ChartTitle ExtendedTitle(string text)
    {
        var title = new Cx.ChartTitle
        {
            Pos = Cx.SidePos.T,
            Align = Cx.PosAlign.Ctr,
            Overlay = false
        };
        title.AppendChild(ExtendedTitleText(text));
        return title;
    }

    private static Cx.Text ExtendedTitleText(string text)
    {
        var chartText = new Cx.Text();
        chartText.AppendChild(new Cx.RichTextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            TitleParagraph(text)));
        return chartText;
    }
}
