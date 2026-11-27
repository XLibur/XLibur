using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// #1027. A load and save keeps what a note's file says about it beyond its text and its anchor:
/// the comments part's <c>commentPr</c>, the alignment in its VML text box, and the fill, line
/// colour and arrow a LibreOffice note gives itself.
/// </summary>
public class NotePropertiesRoundTripTests
{
    private const string LibreOffice = @"TryToLoad\LO\xlsx\universal-content.xlsx";
    private const string LibreOfficeStrict = @"TryToLoad\LO\xlsx\universal-content-strict.xlsx";

    [Test]
    [Arguments(LibreOffice)]
    [Arguments(LibreOfficeStrict)]
    public async Task An_untouched_note_keeps_its_comment_properties(string fixture)
    {
        var loaded = ReadCommentProperties(Load(fixture));
        await Assert.That(loaded).IsNotEmpty();

        var saved = ReadCommentProperties(Save(fixture, _ => { }));

        await Assert.That(saved).IsEquivalentTo(loaded);
    }

    [Test]
    [Arguments(LibreOffice)]
    [Arguments(LibreOfficeStrict)]
    public async Task A_note_saved_with_its_comment_properties_passes_validation(string fixture)
    {
        using var saved = Save(fixture, _ => { });
        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, isEditable: false);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        var errors = document.WorkbookPart!.WorksheetParts
            .Select(wsp => wsp.WorksheetCommentsPart).OfType<WorksheetCommentsPart>()
            .SelectMany(part => validator.Validate(part))
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task A_resized_note_drops_the_comment_properties_that_no_longer_describe_it()
    {
        var saved = ReadCommentProperties(Save(LibreOffice, wb =>
        {
            var note = wb.Worksheet(1).Cell("H3").GetComment();
            note.Style.Size.SetWidth(note.Style.Size.Width * 2);
        }));

        await Assert.That(saved).IsEmpty();
    }

    [Test]
    public async Task A_realigned_note_drops_the_comment_properties_that_no_longer_describe_it()
    {
        var saved = ReadCommentProperties(Save(LibreOffice, wb =>
            wb.Worksheet(1).Cell("H3").GetComment().Style.Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Center)));

        await Assert.That(saved).IsEmpty();
    }

    [Test]
    public async Task A_note_keeps_its_text_box_alignment()
    {
        var saved = ReadNoteShapes(Save(@"TryToLoad\CommentAndButton.xlsx", _ => { }));

        await Assert.That(saved).Count().IsEqualTo(1);
        await Assert.That(TextBoxDivStyle(saved[0])).IsEqualTo("text-align:right");
    }

    [Test]
    [Arguments(XLDrawingHorizontalAlignment.Left, "text-align:left")]
    [Arguments(XLDrawingHorizontalAlignment.Center, "text-align:center")]
    [Arguments(XLDrawingHorizontalAlignment.Justify, "text-align:justify")]
    public async Task A_new_note_states_its_alignment_in_its_text_box(XLDrawingHorizontalAlignment alignment, string style)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var note = wb.AddWorksheet().Cell("B2").GetComment();
            note.AddText("note");
            note.Style.Alignment.SetHorizontal(alignment);
            wb.SaveAs(saved);
        }

        var shapes = ReadNoteShapes(saved);

        await Assert.That(shapes).Count().IsEqualTo(1);
        await Assert.That(TextBoxDivStyle(shapes[0])).IsEqualTo(style);
    }

    [Test]
    public async Task A_LibreOffice_note_loads_the_fill_and_line_colour_it_gives_itself()
    {
        using var source = Load(LibreOffice);
        source.Position = 0;
        using var wb = new XLWorkbook(source);

        var colors = wb.Worksheet(1).Cell("H3").GetComment().Style.ColorsAndLines;

        await Assert.That(colors.FillColor.Color.ToArgb()).IsEqualTo(unchecked((int)0xFFFFFFC0));
        await Assert.That(colors.LineColor.Color.ToArgb()).IsEqualTo(unchecked((int)0xFF808080));
    }

    [Test]
    [Arguments(LibreOffice, "block,medium,medium")]
    [Arguments(LibreOfficeStrict, "block,,")]
    public async Task A_LibreOffice_note_keeps_its_start_arrow(string fixture, string arrow)
    {
        var shapes = ReadNoteShapes(Save(fixture, _ => { }));

        await Assert.That(shapes).Count().IsEqualTo(1);
        var stroke = shapes[0].Element(Vml + "stroke")!;
        await Assert.That(string.Join(",", stroke.Attribute("startarrow")?.Value,
            stroke.Attribute("startarrowlength")?.Value, stroke.Attribute("startarrowwidth")?.Value)).IsEqualTo(arrow);
    }

    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace ExcelVml = "urn:schemas-microsoft-com:office:excel";

    private static List<XElement> ReadNoteShapes(MemoryStream ms) =>
        ReadParts<VmlDrawingPart>(ms)
            .SelectMany(d => d.Descendants(Vml + "shape"))
            .Where(s => s.Element(ExcelVml + "ClientData")?.Attribute("ObjectType")?.Value == "Note")
            .ToList();

    private static string? TextBoxDivStyle(XElement shape) =>
        shape.Element(Vml + "textbox")?.Elements().SingleOrDefault(e => e.Name.LocalName == "div")?.Attribute("style")?.Value;

    private static MemoryStream Load(string fixture)
    {
        var ms = new MemoryStream();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(fixture)))
            stream.CopyTo(ms);
        return ms;
    }

    private static MemoryStream Save(string fixture, Action<XLWorkbook> edit)
    {
        var saved = new MemoryStream();
        using var source = Load(fixture);
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit(wb);
        wb.SaveAs(saved);
        return saved;
    }

    /// <summary>
    /// Each note's <c>commentPr</c> by its cell, as its attributes and its anchor's, then the
    /// anchor's corners. Names are local, so a strict file compares with the transitional save.
    /// </summary>
    private static Dictionary<string, string> ReadCommentProperties(MemoryStream ms)
    {
        var properties = new Dictionary<string, string>();
        foreach (var comment in ReadParts<WorksheetCommentsPart>(ms).SelectMany(d => d.Descendants().Where(e => e.Name.LocalName == "comment")))
        {
            var commentPr = comment.Descendants().SingleOrDefault(e => e.Name.LocalName == "commentPr");
            if (commentPr is null)
                continue;

            var anchor = commentPr.Elements().Single(e => e.Name.LocalName == "anchor");
            properties.Add(comment.Attribute("ref")!.Value, string.Join(" | ",
                Attributes(commentPr),
                Attributes(anchor),
                string.Join(";", anchor.Elements().Select(corner =>
                    corner.Name.LocalName + ":" + string.Join(",", corner.Elements().Select(e => e.Name.LocalName + "=" + e.Value))))));
        }

        return properties;
    }

    private static string Attributes(XElement element) =>
        string.Join(",", element.Attributes().Where(a => !a.IsNamespaceDeclaration)
            .Select(a => a.Name.LocalName + "=" + a.Value).Order(StringComparer.Ordinal));

    private static List<XDocument> ReadParts<TPart>(MemoryStream ms) where TPart : OpenXmlPart
    {
        ms.Position = 0;
        using var ssd = SpreadsheetDocument.Open(ms, isEditable: false);
        return ssd.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<TPart>())
            .Select(part =>
            {
                using var stream = part.GetStream();
                return XDocument.Load(stream);
            })
            .ToList();
    }
}
