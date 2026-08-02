using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

public class ThreadedCommentsTests
{
    private const string ThreadRootId = "{9E032651-3A03-49E4-A15E-A90318F086F4}";

    private const string PersonId = "{273A9FF0-C2D7-46D8-8991-E442B5127E22}";

    #region Round trip of an Excel authored file

    [Test]
    public async Task Round_trip_preserves_thread_structure()
    {
        using var saved = LoadResourceAndSave(@"TryToLoad\ThreadedComment.xlsx");
        using var wb = new XLWorkbook(saved);
        var thread = wb.Worksheets.First().Cell("A1").GetThreadedComment()!;

        await Assert.That(thread.Text).IsEqualTo("This is a threaded comment.");
        await Assert.That(thread.Id).IsEqualTo(Guid.Parse(ThreadRootId));
        await Assert.That(thread.Author.DisplayName).IsEqualTo("Herzog, Bernd");
        await Assert.That(thread.Author.Id).IsEqualTo(Guid.Parse(PersonId));
        await Assert.That(thread.CreatedUtc)
            .IsEqualTo(new DateTime(2020, 7, 1, 6, 20, 6, 20, DateTimeKind.Utc));

        await Assert.That(thread.Replies.Count).IsEqualTo(1);
        await Assert.That(thread.Replies[0].Text).IsEqualTo("This is a reply.");
        await Assert.That(thread.Replies[0].CreatedUtc)
            .IsEqualTo(new DateTime(2020, 7, 1, 6, 20, 22, 880, DateTimeKind.Utc));
    }

    [Test]
    public async Task Round_trip_keeps_person_and_thread_ids_stable()
    {
        using var saved = LoadResourceAndSave(@"TryToLoad\ThreadedComment.xlsx");

        var personXml = ReadPart(saved, "xl/persons/");
        await Assert.That(personXml).Contains($"id=\"{PersonId}\"");
        await Assert.That(personXml).Contains("displayName=\"Herzog, Bernd\"");
        await Assert.That(personXml).Contains("providerId=\"AD\"");
        await Assert.That(personXml).Contains("userId=\"S-1-5-21-2931304574-606859833-1339073683-15873\"");

        var threadXml = ReadPart(saved, "xl/threadedComments/");
        await Assert.That(threadXml).Contains($"id=\"{ThreadRootId}\"");
        await Assert.That(threadXml).Contains($"personId=\"{PersonId}\"");
        await Assert.That(threadXml).Contains($"parentId=\"{ThreadRootId}\"");
        await Assert.That(threadXml).Contains("dT=\"2020-07-01T06:20:06.02\"");
        await Assert.That(threadXml).Contains("<text>This is a threaded comment.</text>");
        await Assert.That(threadXml).Contains("<text>This is a reply.</text>");
    }

    [Test]
    public async Task Round_trip_writes_the_legacy_fallback_note_excel_pairs_with_a_thread()
    {
        using var saved = LoadResourceAndSave(@"TryToLoad\ThreadedComment.xlsx");

        // Older Excel shows this note; 365 hides it and shows the thread. The pairing is the
        // "tc={rootId}" author together with the xr:uid pointing at the same root.
        var commentsXml = ReadPart(saved, "xl/comments1.xml");
        await Assert.That(commentsXml).Contains($">tc={ThreadRootId}<");
        await Assert.That(commentsXml).Contains($"uid=\"{ThreadRootId}\"");
        await Assert.That(commentsXml).Contains("[Threaded comment]");
        await Assert.That(commentsXml).Contains("Comment:");
        await Assert.That(commentsXml).Contains("Reply:");
    }

    [Test]
    public async Task Round_trip_regenerates_the_fallback_note_from_the_edited_thread()
    {
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\ThreadedComment.xlsx"));
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
        {
            wb.Worksheets.First().Cell("A1").GetThreadedComment()!.Text = "Edited root.";
            wb.SaveAs(ms, validate: true);
        }

        // The fallback is derived from the thread on every save, so an edit cannot leave the two
        // showing different text to different Excel versions.
        var commentsXml = ReadPart(ms, "xl/comments1.xml");
        await Assert.That(commentsXml).Contains("Edited root.");
        await Assert.That(commentsXml).DoesNotContain("This is a threaded comment.");
    }

    #endregion

    #region Threads created through the API

    [Test]
    public async Task Can_create_a_thread_from_scratch_and_read_it_back()
    {
        using var ms = new MemoryStream();
        DateTime createdUtc;
        DateTime repliedUtc;

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var reviewer = wb.Persons.Add("Reviewer");
            var author = wb.Persons.Add("Author", "S-1-5-21-1", "AD");

            var thread = ws.Cell("B2").CreateThreadedComment(reviewer, "Please check this figure.");
            var reply = thread.AddReply(author, "Checked, it is right.");
            createdUtc = thread.CreatedUtc;
            repliedUtc = reply.CreatedUtc;

            wb.SaveAs(ms, validate: true);
        }

        using var reloaded = new XLWorkbook(ms);
        var loadedWs = reloaded.Worksheet("Sheet1");
        var loaded = loadedWs.Cell("B2").GetThreadedComment()!;

        await Assert.That(loaded.Text).IsEqualTo("Please check this figure.");
        await Assert.That(loaded.Author.DisplayName).IsEqualTo("Reviewer");
        await Assert.That(loaded.Author.UserId).IsNull();
        await Assert.That(loaded.CreatedUtc).IsEqualTo(createdUtc);
        await Assert.That(loaded.CreatedUtc.Kind).IsEqualTo(DateTimeKind.Utc);

        await Assert.That(loaded.Replies.Count).IsEqualTo(1);
        await Assert.That(loaded.Replies[0].Text).IsEqualTo("Checked, it is right.");
        await Assert.That(loaded.Replies[0].Author.DisplayName).IsEqualTo("Author");
        await Assert.That(loaded.Replies[0].Author.UserId).IsEqualTo("S-1-5-21-1");
        await Assert.That(loaded.Replies[0].Author.ProviderId).IsEqualTo("AD");
        await Assert.That(loaded.Replies[0].CreatedUtc).IsEqualTo(repliedUtc);

        await Assert.That(reloaded.Persons.Count).IsEqualTo(2);
    }

    [Test]
    public async Task Resolved_state_survives_a_round_trip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Reviewer");
            ws.Cell("A1").CreateThreadedComment(person, "Resolved thread.").Resolved = true;
            ws.Cell("A2").CreateThreadedComment(person, "Open thread.");
            wb.SaveAs(ms, validate: true);
        }

        await Assert.That(ReadPart(ms, "xl/threadedComments/")).Contains("done=\"1\"");

        using var reloaded = new XLWorkbook(ms);
        var ws2 = reloaded.Worksheet("Sheet1");
        await Assert.That(ws2.Cell("A1").GetThreadedComment()!.Resolved).IsTrue();
        await Assert.That(ws2.Cell("A2").GetThreadedComment()!.Resolved).IsFalse();
    }

    [Test]
    public async Task Resolved_reads_through_from_a_reply_but_cannot_be_set_on_one()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = ws.Cell("A1").CreateThreadedComment(person, "Root.");
        var reply = thread.AddReply(person, "Reply.");

        thread.Resolved = true;
        await Assert.That(reply.Resolved).IsTrue();

        await Assert.That(() => reply.Resolved = false).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task A_sheet_without_threads_writes_no_threaded_comment_part()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").GetComment().AddText("A plain note.");
            wb.SaveAs(ms, validate: true);
        }

        await Assert.That(PartExists(ms, "xl/threadedComments/")).IsFalse();
        await Assert.That(PartExists(ms, "xl/persons/")).IsFalse();
        await Assert.That(PartExists(ms, "xl/comments1.xml")).IsTrue();
    }

    [Test]
    public async Task Deleting_every_thread_removes_the_threaded_comment_part()
    {
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\ThreadedComment.xlsx"));
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
        {
            wb.Worksheets.First().Cell("A1").GetThreadedComment()!.Delete();
            wb.SaveAs(ms, validate: true);
        }

        await Assert.That(PartExists(ms, "xl/threadedComments/")).IsFalse();

        using var reloaded = new XLWorkbook(ms);
        var cell = reloaded.Worksheets.First().Cell("A1");
        await Assert.That(cell.HasThreadedComment).IsFalse();

        // The fallback note went with the thread rather than being left behind as an orphan.
        await Assert.That(cell.HasComment).IsFalse();
    }

    [Test]
    public async Task Mentions_survive_a_round_trip_of_an_untouched_comment()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
            var thread = (XLibur.Excel.XLThreadedComment)ws.Cell("A1")
                .CreateThreadedComment(person, "@Reviewer please look");
            thread.MentionsXml =
                "<mentions xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\">" +
                $"<mention mentionpersonId=\"{{{person.Id.ToString().ToUpperInvariant()}}}\" " +
                "mentionId=\"{11111111-1111-1111-1111-111111111111}\" startIndex=\"0\" length=\"9\"/></mentions>";

            wb.SaveAs(ms, validate: true);
        }

        var threadXml = ReadPart(ms, "xl/threadedComments/");
        await Assert.That(threadXml).Contains("<mention");
        await Assert.That(threadXml).Contains("startIndex=\"0\"");
    }

    [Test]
    public async Task Editing_the_text_drops_mentions_whose_offsets_no_longer_apply()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        var thread = (XLibur.Excel.XLThreadedComment)ws.Cell("A1").CreateThreadedComment(person, "@Reviewer hi");
        thread.MentionsXml = "<mentions/>";

        thread.Text = "Something else entirely";

        await Assert.That(thread.MentionsXml).IsNull();
    }

    #endregion

    #region A cell holds either a note or a thread

    [Test]
    public async Task Creating_a_thread_on_a_cell_with_a_note_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").GetComment().AddText("A plain note.");

        await Assert.That(() => ws.Cell("A1").CreateThreadedComment(person, "Thread"))
            .Throws<InvalidOperationException>();
    }

    [Test]
    public async Task Creating_a_note_on_a_cell_with_a_thread_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").CreateThreadedComment(person, "Thread");

        await Assert.That(() => ws.Cell("A1").GetComment()).Throws<InvalidOperationException>();
        await Assert.That(() => ws.Cell("A1").CreateComment()).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task Deleting_a_note_frees_the_cell_for_a_thread()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A1").GetComment().AddText("A plain note.");
        ws.Cell("A1").Clear(XLClearOptions.Comments);

        var thread = ws.Cell("A1").CreateThreadedComment(person, "Thread");
        await Assert.That(thread.Text).IsEqualTo("Thread");
        await Assert.That(ws.Cell("A1").HasComment).IsFalse();
    }

    #endregion

    #region Persons

    [Test]
    public async Task Adding_a_person_generates_a_unique_id()
    {
        using var wb = new XLWorkbook();
        var first = wb.Persons.Add("Reviewer");
        var second = wb.Persons.Add("Reviewer");

        await Assert.That(first.Id).IsNotEqualTo(second.Id);
        await Assert.That(wb.Persons.Count).IsEqualTo(2);
        await Assert.That(wb.Persons.Get(first.Id)).IsEqualTo(first);
        await Assert.That(wb.Persons.GetByDisplayName("Reviewer")).IsEqualTo(first);
        await Assert.That(wb.Persons.Get(Guid.NewGuid())).IsNull();
    }

    [Test]
    public async Task Removing_a_person_takes_it_out_of_the_list()
    {
        using var wb = new XLWorkbook();
        var person = wb.Persons.Add("Reviewer");

        await Assert.That(wb.Persons.Remove(person.Id)).IsTrue();
        await Assert.That(wb.Persons.Count).IsEqualTo(0);
        await Assert.That(wb.Persons.Remove(person.Id)).IsFalse();
    }

    #endregion

    #region Helpers

    private static MemoryStream LoadResourceAndSave(string resourcePath)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resourcePath));
        var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
            wb.SaveAs(ms, validate: true);

        return ms;
    }

    /// <summary>
    /// Reads the single part under <paramref name="partPathPrefix"/>. Parts are matched by prefix
    /// and case insensitively because the OpenXML SDK names a part it creates itself differently
    /// from one Excel wrote — "xl/threadedcomments/threadedcomment.xml" rather than
    /// "xl/threadedComments/". Excel resolves parts through relationships, so
    /// the name is not part of the contract.
    /// </summary>
    private static string ReadPart(MemoryStream package, string partPathPrefix)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = FindEntry(archive, partPathPrefix)
                    ?? throw new InvalidOperationException(
                        $"The package has no part under '{partPathPrefix}'. It has: " +
                        string.Join(", ", archive.Entries.Select(e => e.FullName)));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static bool PartExists(MemoryStream package, string partPathPrefix)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return FindEntry(archive, partPathPrefix) is not null;
    }

    // FirstOrDefault returns null when nothing matches, and both callers handle that. The project
    // now has a nullable context, so the annotation says so rather than being suppressed.
    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string partPathPrefix)
    {
        return archive.Entries.FirstOrDefault(e =>
            e.FullName.StartsWith(partPathPrefix, StringComparison.OrdinalIgnoreCase));
    }

    #endregion
}
