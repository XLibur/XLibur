using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// Threads live in the misc slice next to legacy notes, so shifting, swapping, clearing and copying
/// should treat the two the same. These tests pin that down.
/// </summary>
public class ThreadedCommentEditingTests
{
    #region Deleting

    [Test]
    public async Task Deleting_the_root_deletes_the_whole_thread()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = ws.Cell("A1").CreateThreadedComment(person, "Root.");
        thread.AddReply(person, "First reply.");
        thread.AddReply(person, "Second reply.");

        thread.Delete();

        await Assert.That(ws.Cell("A1").HasThreadedComment).IsFalse();
        await Assert.That(ws.Cell("A1").GetThreadedComment()).IsNull();
    }

    [Test]
    public async Task Deleting_a_reply_leaves_the_rest_of_the_thread()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = ws.Cell("A1").CreateThreadedComment(person, "Root.");
        var first = thread.AddReply(person, "First reply.");
        thread.AddReply(person, "Second reply.");

        first.Delete();

        await Assert.That(ws.Cell("A1").HasThreadedComment).IsTrue();
        await Assert.That(thread.Replies.Count).IsEqualTo(1);
        await Assert.That(thread.Replies[0].Text).IsEqualTo("Second reply.");
    }

    [Test]
    public async Task Clearing_comments_removes_a_thread()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").CreateThreadedComment(person, "Root.");

        ws.Cell("A1").Clear(XLClearOptions.Comments);

        await Assert.That(ws.Cell("A1").HasThreadedComment).IsFalse();
    }

    [Test]
    public async Task Clearing_contents_leaves_a_thread_alone()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").Value = "text";
        ws.Cell("A1").CreateThreadedComment(person, "Root.");

        ws.Cell("A1").Clear(XLClearOptions.Contents);

        await Assert.That(ws.Cell("A1").HasThreadedComment).IsTrue();
    }

    [Test]
    public async Task Deleting_comments_over_a_range_removes_threads()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").CreateThreadedComment(person, "First.");
        ws.Cell("A2").CreateThreadedComment(person, "Second.");

        ws.Range("A1:A2").DeleteComments();

        await Assert.That(ws.Cell("A1").HasThreadedComment).IsFalse();
        await Assert.That(ws.Cell("A2").HasThreadedComment).IsFalse();
    }

    #endregion

    #region Shifting

    [Test]
    public async Task Inserting_rows_shifts_a_thread_down()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = ws.Cell("A5").CreateThreadedComment(person, "Root.");
        thread.AddReply(person, "Reply.");

        ws.Row(1).InsertRowsAbove(2);

        await Assert.That(ws.Cell("A5").HasThreadedComment).IsFalse();
        await Assert.That(ws.Cell("A7").HasThreadedComment).IsTrue();

        var moved = ws.Cell("A7").GetThreadedComment()!;
        await Assert.That(moved.Text).IsEqualTo("Root.");
        await Assert.That(moved.Replies.Count).IsEqualTo(1);
    }

    [Test]
    public async Task Deleting_rows_shifts_a_thread_up()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A5").CreateThreadedComment(person, "Root.");

        ws.Row(1).Delete();

        await Assert.That(ws.Cell("A5").HasThreadedComment).IsFalse();
        await Assert.That(ws.Cell("A4").GetThreadedComment()!.Text).IsEqualTo("Root.");
    }

    [Test]
    public async Task Inserting_columns_shifts_a_thread_right()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("C1").CreateThreadedComment(person, "Root.");

        ws.Column(1).InsertColumnsBefore(2);

        await Assert.That(ws.Cell("C1").HasThreadedComment).IsFalse();
        await Assert.That(ws.Cell("E1").GetThreadedComment()!.Text).IsEqualTo("Root.");
    }

    [Test]
    public async Task Deleting_a_row_deletes_the_threads_on_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").CreateThreadedComment(person, "Doomed.");
        ws.Cell("A2").CreateThreadedComment(person, "Survivor.");

        ws.Row(1).Delete();

        await Assert.That(ws.Cell("A1").GetThreadedComment()!.Text).IsEqualTo("Survivor.");
    }

    [Test]
    public async Task A_shifted_thread_still_saves_against_its_new_address()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Reviewer");
            ws.Cell("A5").CreateThreadedComment(person, "Root.");
            ws.Row(1).InsertRowsAbove(2);
            wb.SaveAs(ms, validate: true);
        }

        using var reloaded = new XLWorkbook(ms);
        var loadedWs = reloaded.Worksheet("Sheet1");
        await Assert.That(loadedWs.Cell("A5").HasThreadedComment).IsFalse();
        await Assert.That(loadedWs.Cell("A7").GetThreadedComment()!.Text).IsEqualTo("Root.");
    }

    #endregion

    #region Copying

    [Test]
    public async Task Copying_a_cell_copies_the_thread_with_a_new_id()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var source = ws.Cell("A1").CreateThreadedComment(person, "Root.");
        source.AddReply(person, "Reply.");

        ws.Cell("A1").CopyTo(ws.Cell("C3"));

        var copy = ws.Cell("C3").GetThreadedComment()!;
        await Assert.That(copy.Text).IsEqualTo("Root.");
        await Assert.That(copy.Replies.Count).IsEqualTo(1);
        await Assert.That(copy.Replies[0].Text).IsEqualTo("Reply.");
        await Assert.That(copy.Author).IsEqualTo(source.Author);

        // Two threads may not share an id: it is what the fallback note's uid points at.
        await Assert.That(copy.Id).IsNotEqualTo(source.Id);
        await Assert.That(copy.Replies[0].Id).IsNotEqualTo(source.Replies[0].Id);

        // The original is untouched.
        await Assert.That(ws.Cell("A1").GetThreadedComment()!.Text).IsEqualTo("Root.");
    }

    [Test]
    public async Task Copying_a_thread_over_a_note_replaces_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").CreateThreadedComment(person, "Root.");
        ws.Cell("C3").GetComment().AddText("A plain note.");

        ws.Cell("A1").CopyTo(ws.Cell("C3"));

        await Assert.That(ws.Cell("C3").HasThreadedComment).IsTrue();
        await Assert.That(ws.Cell("C3").HasComment).IsFalse();
    }

    [Test]
    public async Task Copying_a_note_over_a_thread_replaces_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").GetComment().AddText("A plain note.");
        ws.Cell("C3").CreateThreadedComment(person, "Root.");

        ws.Cell("A1").CopyTo(ws.Cell("C3"));

        await Assert.That(ws.Cell("C3").HasComment).IsTrue();
        await Assert.That(ws.Cell("C3").HasThreadedComment).IsFalse();
    }

    [Test]
    public async Task Copying_across_workbooks_maps_the_person_into_the_target()
    {
        using var source = new XLWorkbook();
        var sourceWs = source.AddWorksheet("Sheet1");
        var sourcePerson = source.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
        sourceWs.Cell("A1").CreateThreadedComment(sourcePerson, "Root.");

        using var target = new XLWorkbook();
        var targetWs = target.AddWorksheet("Sheet1");

        sourceWs.Cell("A1").CopyTo(targetWs.Cell("A1"));

        var copied = targetWs.Cell("A1").GetThreadedComment()!;
        await Assert.That(copied.Text).IsEqualTo("Root.");
        await Assert.That(copied.Author.DisplayName).IsEqualTo("Reviewer");
        await Assert.That(copied.Author.UserId).IsEqualTo("S-1-5-21-1");

        // The person is now owned by the target workbook, not borrowed from the source.
        await Assert.That(target.Persons.Count).IsEqualTo(1);
        await Assert.That(target.Persons.Get(copied.Author.Id)).IsEqualTo(copied.Author);
    }

    [Test]
    public async Task Copying_two_threads_by_the_same_person_adds_that_person_once()
    {
        using var source = new XLWorkbook();
        var sourceWs = source.AddWorksheet("Sheet1");
        var sourcePerson = source.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
        sourceWs.Cell("A1").CreateThreadedComment(sourcePerson, "First.");
        sourceWs.Cell("A2").CreateThreadedComment(sourcePerson, "Second.");

        using var target = new XLWorkbook();
        var targetWs = target.AddWorksheet("Sheet1");
        sourceWs.Range("A1:A2").CopyTo(targetWs.Cell("A1"));

        await Assert.That(target.Persons.Count).IsEqualTo(1);
        await Assert.That(targetWs.Cell("A1").GetThreadedComment()!.Author)
            .IsEqualTo(targetWs.Cell("A2").GetThreadedComment()!.Author);
    }

    [Test]
    public async Task A_thread_copied_across_workbooks_saves_and_reloads()
    {
        using var ms = new MemoryStream();
        using (var source = new XLWorkbook())
        using (var target = new XLWorkbook())
        {
            var sourceWs = source.AddWorksheet("Sheet1");
            var person = source.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
            var thread = sourceWs.Cell("A1").CreateThreadedComment(person, "Root.");
            thread.AddReply(person, "Reply.");

            var targetWs = target.AddWorksheet("Sheet1");
            sourceWs.Cell("A1").CopyTo(targetWs.Cell("B2"));
            target.SaveAs(ms, validate: true);
        }

        using var reloaded = new XLWorkbook(ms);
        var loaded = reloaded.Worksheet("Sheet1").Cell("B2").GetThreadedComment()!;
        await Assert.That(loaded.Text).IsEqualTo("Root.");
        await Assert.That(loaded.Replies.Count).IsEqualTo(1);
        await Assert.That(loaded.Author.DisplayName).IsEqualTo("Reviewer");
        await Assert.That(reloaded.Persons.Count).IsEqualTo(1);
    }

    #endregion

    #region Cells used

    [Test]
    public async Task A_cell_with_only_a_thread_counts_as_used_for_comments()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("C3").CreateThreadedComment(person, "Root.");

        var used = ws.CellsUsed(XLCellsUsedOptions.All).ToList();
        await Assert.That(used.Count).IsEqualTo(1);
        await Assert.That(used[0].Address.ToString()).IsEqualTo("C3");

        // A thread is an annotation, not content.
        await Assert.That(ws.CellsUsed(XLCellsUsedOptions.Contents).Any()).IsFalse();
    }

    #endregion

    #region Replies

    [Test]
    public async Task Replies_are_flat_so_replying_to_a_reply_appends_to_the_thread()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = ws.Cell("A1").CreateThreadedComment(person, "Root.");
        var first = thread.AddReply(person, "First.");
        var second = first.AddReply(person, "Second.");

        await Assert.That(thread.Replies.Count).IsEqualTo(2);
        await Assert.That(first.Replies.Count).IsEqualTo(0);
        await Assert.That(second.Parent).IsEqualTo(thread);
    }

    [Test]
    public async Task Reply_order_survives_a_round_trip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Reviewer");
            var thread = ws.Cell("A1").CreateThreadedComment(person, "Root.");
            for (var i = 1; i <= 5; i++)
                thread.AddReply(person, $"Reply {i}.");

            wb.SaveAs(ms, validate: true);
        }

        using var reloaded = new XLWorkbook(ms);
        var loaded = reloaded.Worksheet("Sheet1").Cell("A1").GetThreadedComment()!;

        await Assert.That(loaded.Replies.Count).IsEqualTo(5);
        for (var i = 1; i <= 5; i++)
            await Assert.That(loaded.Replies[i - 1].Text).IsEqualTo($"Reply {i}.");
    }

    [Test]
    public async Task A_reply_by_a_person_from_another_workbook_is_mapped_in()
    {
        using var other = new XLWorkbook();
        var stranger = other.Persons.Add("Stranger", "S-1-5-21-9", "AD");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = ws.Cell("A1").CreateThreadedComment(person, "Root.");

        var reply = thread.AddReply(stranger, "Reply from elsewhere.");

        await Assert.That(reply.Author.DisplayName).IsEqualTo("Stranger");
        await Assert.That(wb.Persons.Count).IsEqualTo(2);
        await Assert.That(wb.Persons.Get(reply.Author.Id)).IsNotNull();
    }

    #endregion
}
