using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// #905. Excel writes every note free-floating, locked and with a shadow:
/// <c>&lt;x:MoveWithCells/&gt;&lt;x:SizeWithCells/&gt;</c>, no <c>x:Locked</c>, and
/// <c>&lt;v:shadow on="t"/&gt;</c>. A load and save must keep all three.
/// </summary>
public class NoteVmlRoundTripTests
{
    private static readonly XNamespace ExcelVml = "urn:schemas-microsoft-com:office:excel";
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";

    private const string Fixture = @"Examples\Comments\EditingComments.xlsx";

    [Test]
    public async Task An_Excel_note_loads_free_floating_and_locked()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        using var wb = new XLWorkbook(stream);

        var notes = Notes(wb);
        await Assert.That(notes).Count().IsEqualTo(5);
        foreach (var note in notes)
        {
            await Assert.That(note.Anchor).IsEqualTo(XLDrawingAnchor.Absolute);
            await Assert.That(note.Style.Protection.Locked).IsTrue();
            await Assert.That(note.Style.Protection.LockText).IsTrue();
        }
    }

    [Test]
    public async Task An_Excel_note_keeps_its_placement_lock_and_shadow_through_a_save()
    {
        using var saved = new MemoryStream();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture)))
        using (var wb = new XLWorkbook(stream))
            wb.SaveAs(saved);

        using (var reloaded = new XLWorkbook(saved))
        {
            var notes = Notes(reloaded);
            await Assert.That(notes).Count().IsEqualTo(5);
            foreach (var note in notes)
            {
                await Assert.That(note.Anchor).IsEqualTo(XLDrawingAnchor.Absolute);
                await Assert.That(note.Style.Protection.Locked).IsTrue();
            }
        }

        var shapes = ReadNoteShapes(saved);
        await Assert.That(shapes).Count().IsEqualTo(5);
        foreach (var shape in shapes)
        {
            await Assert.That(shape.Element(Vml + "shadow")?.Attribute("on")?.Value).IsEqualTo("true");

            var clientData = shape.Element(ExcelVml + "ClientData")!;
            await Assert.That(clientData.Element(ExcelVml + "MoveWithCells")?.Value).IsEqualTo("True");
            await Assert.That(clientData.Element(ExcelVml + "SizeWithCells")?.Value).IsEqualTo("True");
            await Assert.That(clientData.Element(ExcelVml + "Locked")?.Value).IsNotEqualTo("False");
        }
    }

    [Test]
    public async Task A_free_floating_note_still_follows_its_cell_when_rows_are_inserted_above()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet(1);
        var rowBefore = ws.Cell("B3").GetComment().Position.Row;

        ws.Row(1).InsertRowsAbove(2);

        var note = (XLComment)ws.Cell("B5").GetComment();
        await Assert.That(note.Anchor).IsEqualTo(XLDrawingAnchor.Absolute);
        await Assert.That(note.Position.Row).IsEqualTo(rowBefore + 2);
    }

    private static List<XLComment> Notes(XLWorkbook wb) =>
        wb.Worksheets.SelectMany(ws => ws.CellsUsed(XLCellsUsedOptions.Comments))
            .Where(c => c.HasComment)
            .Select(c => (XLComment)c.GetComment())
            .ToList();

    private static List<XElement> ReadNoteShapes(MemoryStream ms)
    {
        ms.Position = 0;
        using var ssd = SpreadsheetDocument.Open(ms, isEditable: false);
        return ssd.WorkbookPart!.WorksheetParts
            .SelectMany(wsp => wsp.GetPartsOfType<VmlDrawingPart>())
            .SelectMany(part =>
            {
                using var vml = part.GetStream();
                return XDocument.Load(vml).Descendants(Vml + "shape")
                    .Where(s => s.Element(ExcelVml + "ClientData")?.Attribute("ObjectType")?.Value == "Note")
                    .ToList();
            })
            .ToList();
    }
}
