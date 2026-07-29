using XLibur.Excel;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Comments;

public partial class CommentsTests
{
    [Test]
    public async Task CanConvertVmlPaletteEntriesToColors()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentsWithColorNamesAndIndexes.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var c = ws.FirstCellUsed();

        // None indicates an absence of a color
        var lineColor = c.GetComment().Style.ColorsAndLines.LineColor;
        await Assert.That(lineColor.ColorType).IsEqualTo(XLColorType.Color);
        await Assert.That(lineColor.Color.ToHex()).IsEqualTo("00000000");

        var bgColor = c.GetComment().Style.ColorsAndLines.FillColor;
        await Assert.That(bgColor.ColorType).IsEqualTo(XLColorType.Color);
        await Assert.That(bgColor.Color.ToHex()).IsEqualTo("FFFFFFE1");
    }

    [Test]
    public async Task CopyCommentStyle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var strExcelComment = "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "2) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "3) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "4) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "5) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "6) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "7) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "8) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
        strExcelComment = strExcelComment + "9) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;

        var cell = ws.Cell(2, 2).SetValue("Comment 1");

        cell.GetComment()
            .SetVisible(false)
            .AddText(strExcelComment);

        cell.GetComment()
            .Style
            .Alignment
            .SetAutomaticSize();

        cell.GetComment()
            .Style
            .ColorsAndLines
            .SetFillColor(XLColor.Red);

        ws.Row(1).InsertRowsAbove(1);

        async Task Validate(IXLCell c)
        {
            await Assert.That(c.GetComment().Style.Alignment.AutomaticSize).IsTrue();
            await Assert.That(c.GetComment().Style.ColorsAndLines.FillColor).IsEqualTo(XLColor.Red);
        }

        await Validate(ws.Cell("B3"));

        ws.Column(1).InsertColumnsBefore(2);

        await Validate(ws.Cell("D3"));

        ws.Column(1).Delete();

        await Validate(ws.Cell("C3"));

        ws.Row(1).Delete();

        await Validate(ws.Cell("C2"));
    }

    [Test]
    public async Task EnsureUnaffectedCommentAndVmlPartIdsAndUris()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx"));
        using var ms = new MemoryStream();
        string commentPartId;
        string commentPartUri;

        string vmlPartId;
        string vmlPartUri;

        using (var ssd = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            var wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
            var wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

            var wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
            commentPartId = wsp.GetIdOfPart(wscp);
            commentPartUri = wscp.Uri.ToString();

            var vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
            vmlPartId = wsp.GetIdOfPart(vmlp);
            vmlPartUri = vmlp.Uri.ToString();
        }

        stream.Position = 0;
        stream.CopyTo(ms);
        ms.Position = 0;

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.FirstCell().HasComment).IsTrue();

            wb.SaveAs(ms);
        }

        ms.Position = 0;

        using (var ssd = SpreadsheetDocument.Open(ms, isEditable: false))
        {
            var wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
            var wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

            var wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
            await Assert.That(wscp.Uri.ToString()).IsEqualTo(commentPartUri);
            await Assert.That(wsp.GetIdOfPart(wscp)).IsEqualTo(commentPartId);

            var vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
            await Assert.That(vmlp.Uri.ToString()).IsEqualTo(vmlPartUri);
            await Assert.That(wsp.GetIdOfPart(vmlp)).IsEqualTo(vmlPartId);
        }
    }

    [Test]
    public async Task SavingDoesNotCauseTwoRootElements() // See #1157
    {
        using var ms = new MemoryStream();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            wb.SaveAs(ms);
        }

        await Assert.That(() => _ = new XLWorkbook(ms)).ThrowsNothing();
    }

    [Test]
    public async Task CanLoadCommentVisibility()
    {
        using var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Drawings\Comments\inputfile.xlsx"));
        using var workbook = new XLWorkbook(inputStream);
        var ws = workbook.Worksheets.First();

        await Assert.That(ws.Cell("A1").GetComment().Visible).IsTrue();
        await Assert.That(ws.Cell("A4").GetComment().Visible).IsFalse();
    }

    [Test]
    public async Task Margins_are_converted_to_physical_length()
    {
        // Technically, it's insets on a textbox. Each comment uses a different unit, but all
        // should have same final dimension at left and top margin (easily visible in the
        // sheet). Tested units: in, cm, mm, pt, pc, emu, px, em, ex. Pixels are converted
        // through supplied DPI.
        // The last comment in vmlDrawing1 also has invalid units and number. These are
        // converted to 0, so we don't crash on load (Excel also ignores invalid values).
        var commentCells = new[] { "A1", "A7", "A16", "A22", "A28" };
        await TestHelper.LoadAndAssert(async (_, ws) =>
        {
            foreach (var commentCell in commentCells)
            {
                var cell = ws.Cell(commentCell);
                await Assert.That(cell.HasComment).IsTrue();
                var margins = cell.GetComment().Style.Margins;

                await Assert.That(margins.Left).IsEqualTo(0.5);
                await Assert.That(margins.Top).IsEqualTo(0.75);

                await Assert.That(margins.Right).IsEqualTo(0);
                await Assert.That(margins.Bottom).IsEqualTo(0);
            }
        }, @"Other\Comments\InsetsUnitConversion.xlsx", new XLibur.Excel.LoadOptions { Dpi = new Point(120, 120) });
    }

    [Test]
    public async Task Can_load_threaded_comment()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\ThreadedComment.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var c = ws.FirstCellUsed()!;

        // The thread is the user-visible annotation; the "[Threaded comment]" note Excel pairs with
        // it is a fallback for older versions and must not surface as a note.
        await Assert.That(c.HasThreadedComment).IsTrue();
        await Assert.That(c.HasComment).IsFalse();

        var thread = c.GetThreadedComment()!;
        await Assert.That(thread.Text).IsEqualTo("This is a threaded comment.");
        await Assert.That(thread.Author.DisplayName).IsEqualTo("Herzog, Bernd");
        await Assert.That(thread.Author.ProviderId).IsEqualTo("AD");
        await Assert.That(thread.Author.UserId).IsEqualTo("S-1-5-21-2931304574-606859833-1339073683-15873");
        await Assert.That(thread.Id).IsEqualTo(Guid.Parse("{9E032651-3A03-49E4-A15E-A90318F086F4}"));
        await Assert.That(thread.Parent).IsNull();
        await Assert.That(thread.Resolved).IsFalse();
        await Assert.That(thread.CreatedUtc)
            .IsEqualTo(new DateTime(2020, 7, 1, 6, 20, 6, 20, DateTimeKind.Utc));

        await Assert.That(thread.Replies.Count).IsEqualTo(1);
        var reply = thread.Replies[0];
        await Assert.That(reply.Text).IsEqualTo("This is a reply.");
        await Assert.That(reply.Parent).IsEqualTo(thread);
        await Assert.That(reply.Author).IsEqualTo(thread.Author);
        await Assert.That(reply.CreatedUtc)
            .IsEqualTo(new DateTime(2020, 7, 1, 6, 20, 22, 880, DateTimeKind.Utc));

        // The person is registered once at workbook level, not once per comment.
        await Assert.That(wb.Persons.Count).IsEqualTo(1);
    }

    [Test]
    public async Task Can_load_threaded_comment_text() // #2344
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\celltextcomment_load_2344.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var c = ws.Cell("B2");

        // A thread with no replies is still a thread, not a note.
        await Assert.That(c.HasThreadedComment).IsTrue();
        await Assert.That(c.HasComment).IsFalse();

        var thread = c.GetThreadedComment()!;
        await Assert.That(thread.Text).IsEqualTo("This is the comment in b2");
        await Assert.That(thread.Replies.Count).IsEqualTo(0);
    }

    [Test]
    public async Task AutomaticSize_fits_comment_box_to_text()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("CommentSize");
            var comment = ws.Cell("A1").CreateComment();
            comment.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            comment.Style.Size.SetAutomaticSize();
            wb.SaveAs(ms);
        }

        // Inspect the VML drawing
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, isEditable: false);
        var wsPart = doc.WorkbookPart!.WorksheetParts.First();
        var vmlPart = wsPart.VmlDrawingParts.First();
        using var vmlStream = vmlPart.GetStream();
        var vml = XDocument.Load(vmlStream);
        var vmlStr = vml.ToString();

        // Verify mso-fit-shape-to-text is present
        await Assert.That(vmlStr).Contains("mso-fit-shape-to-text:t");

        // Verify that the height was auto-sized to be larger than default 59.25pt.
        // The lorem ipsum text at Tahoma 9pt in 144pt-wide box wraps to 5 lines,
        // requiring more height than the default 59.25pt.
        var heightMatch = HeightPattern().Match(vmlStr);
        await Assert.That(heightMatch.Success).IsTrue();
        var height = double.Parse(heightMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        await Assert.That(height).IsGreaterThan(59.25);
    }

    [Test]
    public async Task Can_load_comment_with_missing_textbox_in_vml()
    {
        // Create a workbook with a comment, then strip the textbox element from VML.
        // This reproduces files where notes/comments have shapes without a textbox element.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue("Test");
            ws.Cell("A1").GetComment().AddText("Comment without textbox");
            wb.SaveAs(ms);
        }

        // Remove the textbox element from VML drawing part
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, isEditable: true))
        {
            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var vmlPart = wsPart.VmlDrawingParts.First();
            using var vmlStream = vmlPart.GetStream(FileMode.Open);
            var vml = XDocument.Load(vmlStream);

            var textboxes = vml.Descendants().Where(e => e.Name.LocalName == "textbox").ToList();
            foreach (var tb in textboxes)
                tb.Remove();

            vmlStream.SetLength(0);
            vml.Save(vmlStream);
        }

        // Loading should not throw despite missing textbox
        ms.Position = 0;
        var hasComment = false;
        await Assert.That(() =>
        {
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheets.First();
            hasComment = ws.Cell("A1").HasComment;
        }).ThrowsNothing();
        await Assert.That(hasComment).IsTrue();
    }

    [Test]
    public async Task Deleting_a_note_that_has_been_shifted_clears_the_right_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("B2").GetComment().AddText("Note.");

        ws.Row(1).InsertRowsAbove(2);         // B2 -> B4
        ws.Column(1).InsertColumnsBefore(2);  // B4 -> D4
        ws.Row(1).Delete();                   // D4 -> D3
        ws.Column(1).Delete();                // D3 -> C3
        await Assert.That(ws.Cell("C3").HasComment).IsTrue();

        ws.Cell("C3").GetComment().Delete();

        await Assert.That(ws.Cell("C3").HasComment).IsFalse();
        await Assert.That(ws.CellsUsed(XLCellsUsedOptions.Comments).Count()).IsEqualTo(0);
    }

    [Test]
    public async Task Deleting_a_note_that_is_no_longer_on_the_sheet_leaves_the_cell_alone()
    {
        // A5's note goes with the deleted row and A6's note lands on A5, so deleting the note that
        // is already gone must not take the one that replaced it.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var deletedNote = ws.Cell("A5").GetComment();
        deletedNote.AddText("Five.");
        ws.Cell("A6").GetComment().AddText("Six.");

        ws.Row(5).Delete();
        deletedNote.Delete();

        await Assert.That(ws.Cell("A5").HasComment).IsTrue();
        await Assert.That(ws.Cell("A5").GetComment().Text).IsEqualTo("Six.");
    }

    [GeneratedRegex(@"height:(\d+\.?\d*)pt")]
    private static partial Regex HeightPattern();
}
