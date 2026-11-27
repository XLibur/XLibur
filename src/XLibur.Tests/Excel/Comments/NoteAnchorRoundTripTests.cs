using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// #953. Excel places a note by its <c>&lt;x:Anchor&gt;</c>. A save must write back the anchor it
/// loaded for a note nothing moved or resized, and compute a new one for a note that was.
/// </summary>
public class NoteAnchorRoundTripTests
{
    private static readonly XNamespace ExcelVml = "urn:schemas-microsoft-com:office:excel";
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";

    private const string EditingComments = @"Examples\Comments\EditingComments.xlsx";

    [Test]
    [Arguments(@"Examples\Comments\EditingComments.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\universal-content.xlsx")]
    [Arguments(@"TryToLoad\CommentAndButton.xlsx")]
    public async Task An_untouched_note_keeps_its_loaded_anchor(string fixture)
    {
        var loaded = ReadAnchors(Load(fixture));
        await Assert.That(loaded).IsNotEmpty();

        var saved = ReadAnchors(Save(fixture, _ => { }));

        await Assert.That(saved).IsEquivalentTo(loaded);
    }

    [Test]
    public async Task A_resized_note_gets_a_new_anchor_and_its_neighbours_keep_theirs()
    {
        var loaded = ReadAnchors(Load(EditingComments));

        var saved = ReadAnchors(Save(EditingComments, wb =>
        {
            var note = wb.Worksheet("Comments").Cell("B3").GetComment();
            note.Style.Size.SetWidth(note.Style.Size.Width * 2);
        }));

        await Assert.That(saved["Comments!B3"]).IsNotEqualTo(loaded["Comments!B3"]);
        foreach (var (cell, anchor) in loaded.Where(a => a.Key != "Comments!B3"))
            await Assert.That(saved[cell]).IsEqualTo(anchor);
    }

    [Test]
    public async Task A_note_moved_by_inserted_rows_gets_a_new_anchor()
    {
        var loaded = ReadAnchors(Load(EditingComments));

        var saved = ReadAnchors(Save(EditingComments, wb => wb.Worksheet("Comments").Row(1).InsertRowsAbove(2)));

        await Assert.That(saved).ContainsKey("Comments!B5");
        await Assert.That(saved["Comments!B5"]).IsNotEqualTo(loaded["Comments!B3"]);
    }

    private static MemoryStream Load(string fixture)
    {
        var ms = new MemoryStream();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(fixture)))
            stream.CopyTo(ms);
        return ms;
    }

    private static MemoryStream Save(string fixture, System.Action<XLWorkbook> edit)
    {
        var saved = new MemoryStream();
        using var source = Load(fixture);
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit(wb);
        wb.SaveAs(saved);
        return saved;
    }

    /// <summary>The note anchors of every sheet, by sheet and VML shape id.</summary>
    private static Dictionary<string, string> ReadAnchors(MemoryStream ms)
    {
        ms.Position = 0;
        using var ssd = SpreadsheetDocument.Open(ms, isEditable: false);
        var workbookPart = ssd.WorkbookPart!;
        var anchors = new Dictionary<string, string>();
        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
            foreach (var part in worksheetPart.GetPartsOfType<VmlDrawingPart>())
            {
                using var vml = part.GetStream();
                foreach (var shape in XDocument.Load(vml).Descendants(Vml + "shape"))
                {
                    var clientData = shape.Element(ExcelVml + "ClientData");
                    if (clientData?.Attribute("ObjectType")?.Value != "Note")
                        continue;

                    var row = int.Parse(clientData.Element(ExcelVml + "Row")!.Value) + 1;
                    var column = int.Parse(clientData.Element(ExcelVml + "Column")!.Value) + 1;
                    var key = sheet.Name!.Value + "!" + XLHelper.GetColumnLetterFromNumber(column) + row;
                    anchors.Add(key, Normalize(clientData.Element(ExcelVml + "Anchor")!.Value));
                }
            }
        }

        return anchors;
    }

    // Excel writes the anchor with a line break and spaces around it; only the numbers matter.
    private static string Normalize(string anchor) =>
        string.Join(",", anchor.Split(',').Select(n => n.Trim()));
}
