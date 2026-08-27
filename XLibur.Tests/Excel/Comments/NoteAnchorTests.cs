using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// D17. A note used to state its anchoring mode twice — <see cref="XLComment.Anchor"/> against
/// <c>Style.Properties.Positioning</c> — and the two disagreed on every note XLibur created:
/// <c>Initialize</c> said move-and-size-with-cells while the inherited <c>DefaultCommentStyle</c>
/// said absolute, and the VML writer read the style. The two are now one value.
/// </summary>
public class NoteAnchorTests
{
    [Test]
    public async Task A_new_note_states_one_anchoring_mode_not_two()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var note = (XLComment)ws.Cell("C10").CreateComment();
        note.AddText("note");

        await Assert.That(note.Anchor).IsEqualTo(XLDrawingAnchor.MoveAndSizeWithCells);
        await Assert.That(note.Style.Properties.Positioning).IsEqualTo(note.Anchor);
    }

    [Test]
    public async Task A_new_note_is_written_as_moving_and_sizing_with_its_cell()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("C10").CreateComment().AddText("note");
            wb.SaveAs(ms);
        }

        // Counterintuitive, exactly as VmlDrawingPartWriter's own comment records it: False here
        // means the note *does* move, and *does* resize, with its cells.
        var clientData = ReadNoteClientData(ms);
        await Assert.That(clientData.Element(ExcelVml + "MoveWithCells")?.Value).IsEqualTo("False");
        await Assert.That(clientData.Element(ExcelVml + "SizeWithCells")?.Value).IsEqualTo("False");
    }

    [Test]
    [Arguments(XLDrawingAnchor.MoveAndSizeWithCells)]
    [Arguments(XLDrawingAnchor.MoveWithCells)]
    [Arguments(XLDrawingAnchor.Absolute)]
    public async Task A_notes_anchoring_mode_survives_a_round_trip(XLDrawingAnchor anchor)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var note = (XLComment)ws.Cell("C10").CreateComment();
            note.AddText("note");
            note.Anchor = anchor;
            wb.SaveAs(ms);
        }

        using var loaded = new XLWorkbook(ms);
        var reloaded = (XLComment)loaded.Worksheet("Sheet1").Cell("C10").GetComment();

        await Assert.That(reloaded.Anchor).IsEqualTo(anchor);
        await Assert.That(reloaded.Style.Properties.Positioning).IsEqualTo(anchor);
    }

    private static readonly XNamespace ExcelVml = "urn:schemas-microsoft-com:office:excel";

    private static XElement ReadNoteClientData(MemoryStream ms)
    {
        ms.Position = 0;
        using var ssd = SpreadsheetDocument.Open(ms, isEditable: false);
        var wsp = ssd.GetPartsOfType<WorkbookPart>().Single().GetPartsOfType<WorksheetPart>().Single();
        using var vml = wsp.GetPartsOfType<VmlDrawingPart>().Single().GetStream();
        return XDocument.Load(vml).Descendants(ExcelVml + "ClientData").Single();
    }
}
