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
