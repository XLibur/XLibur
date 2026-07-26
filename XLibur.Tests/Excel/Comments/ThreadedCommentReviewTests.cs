using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Comments;

public class ThreadedCommentReviewTests
{
    [Test]
    public async Task Deleting_a_thread_that_has_been_shifted_clears_the_right_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var person = wb.Persons.Add("Reviewer");
        ws.Cell("A5").CreateThreadedComment(person, "Root.");

        ws.Row(1).InsertRowsAbove(2);
        ws.Cell("A7").GetThreadedComment()!.Delete();

        await Assert.That(ws.Cell("A7").HasThreadedComment).IsFalse();
    }

    [Test]
    public async Task A_person_still_referenced_by_a_thread_is_written_even_if_removed_from_the_list()
    {
        using var ms = new MemoryStream();
        Guid authorId;

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
            authorId = person.Id;
            ws.Cell("A1").CreateThreadedComment(person, "Root.");

            wb.Persons.Remove(person.Id);
            wb.SaveAs(ms, validate: true);
        }

        // A personId with no matching <person> is a dangling reference Excel cannot resolve.
        var personXml = ReadPart(ms, "xl/persons/");
        await Assert.That(personXml).Contains(authorId.ToString("B").ToUpperInvariant());

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Sheet1").Cell("A1").GetThreadedComment()!.Author.DisplayName)
            .IsEqualTo("Reviewer");
    }

    [Test]
    public async Task Two_provider_less_people_sharing_a_name_are_not_merged_across_workbooks()
    {
        using var source = new XLWorkbook();
        var sourceWs = source.AddWorksheet("Sheet1");
        var sourcePerson = source.Persons.Add("Reviewer");
        sourceWs.Cell("A1").CreateThreadedComment(sourcePerson, "From the source workbook.");

        using var target = new XLWorkbook();
        var targetWs = target.AddWorksheet("Sheet1");
        var targetPerson = target.Persons.Add("Reviewer");
        targetWs.Cell("B2").CreateThreadedComment(targetPerson, "Already here.");

        sourceWs.Cell("A1").CopyTo(targetWs.Cell("A1"));

        // Without a userId or providerId there is nothing that says these are the same human, and
        // merging them would reattribute one person's comment to the other.
        var copiedAuthor = targetWs.Cell("A1").GetThreadedComment()!.Author;
        await Assert.That(copiedAuthor.Id).IsNotEqualTo(targetPerson.Id);
        await Assert.That(target.Persons.Count).IsEqualTo(2);
    }

    [Test]
    public async Task Two_people_sharing_a_provider_identity_are_still_merged()
    {
        using var source = new XLWorkbook();
        var sourceWs = source.AddWorksheet("Sheet1");
        var sourcePerson = source.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
        sourceWs.Cell("A1").CreateThreadedComment(sourcePerson, "From the source workbook.");

        using var target = new XLWorkbook();
        var targetWs = target.AddWorksheet("Sheet1");
        var targetPerson = target.Persons.Add("Reviewer", "S-1-5-21-1", "AD");
        targetWs.Cell("B2").CreateThreadedComment(targetPerson, "Already here.");

        sourceWs.Cell("A1").CopyTo(targetWs.Cell("A1"));

        await Assert.That(target.Persons.Count).IsEqualTo(1);
        await Assert.That(targetWs.Cell("A1").GetThreadedComment()!.Author.Id).IsEqualTo(targetPerson.Id);
    }

    [Test]
    public async Task A_comment_with_no_timestamp_still_reports_a_utc_created_date()
    {
        // Excel always writes dT, but a hand-edited or third-party file need not.
        var package = BuildPackageWithThreadMissingTimestamp();
        using var wb = new XLWorkbook(package);

        var thread = wb.Worksheets.First().Cell("A1").GetThreadedComment()!;
        await Assert.That(thread.CreatedUtc.Kind).IsEqualTo(DateTimeKind.Utc);
    }

    private static MemoryStream BuildPackageWithThreadMissingTimestamp()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Reviewer");
            ws.Cell("A1").CreateThreadedComment(person, "No timestamp.");
            wb.SaveAs(ms, validate: true);
        }

        // Strip the dT attribute the writer emitted.
        ms.Position = 0;
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.Entries.First(e =>
                e.FullName.StartsWith("xl/threadedcomments/", StringComparison.OrdinalIgnoreCase));

            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                xml = reader.ReadToEnd();

            xml = System.Text.RegularExpressions.Regex.Replace(xml, " dT=\"[^\"]*\"", string.Empty);

            using var stream = entry.Open();
            stream.SetLength(0);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(xml);
        }

        ms.Position = 0;
        return ms;
    }

    private static string ReadPart(MemoryStream package, string partPathPrefix)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.FirstOrDefault(e =>
                        e.FullName.StartsWith(partPathPrefix, StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException(
                        $"The package has no part under '{partPathPrefix}'. It has: " +
                        string.Join(", ", archive.Entries.Select(e => e.FullName)));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
