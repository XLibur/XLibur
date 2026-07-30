using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.Comments;

/// <summary>
/// The threaded-comment write path runs on every save of every workbook, including the
/// overwhelming majority that have no threads at all. These tests pin down that it costs nothing
/// on those workbooks — a scan that materialises an <c>XLCell</c> per used cell here is invisible
/// in correctness tests but showed up as +23 MB on the 50K-row save benchmark.
/// </summary>
public class ThreadedCommentSaveCostTests
{
    private const int UsedCells = 20_000;

    [Test]
    public async Task Collecting_persons_does_not_materialise_a_cell_per_used_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= UsedCells; row++)
            ws.Cell(row, 1).Value = row;

        // Warm up so JIT and any first-call setup are outside the measured window.
        PersonPartWriter.CollectReferencedPersons(wb);

        var before = GC.GetTotalAllocatedBytes(precise: true);
        var persons = PersonPartWriter.CollectReferencedPersons(wb);
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - before;

        await Assert.That(persons).IsEmpty();

        // One XLCell wrapper per used cell is 48 bytes, so the scan this replaced allocated at
        // least 960 KB here. Walking the misc slice allocates the empty result list and nothing
        // per cell; the bound leaves room for that without leaving room for a per-cell wrapper.
        await Assert.That(allocated).IsLessThan(64 * 1024);
    }

    [Test]
    public async Task Threads_on_many_cells_round_trip_with_the_right_references()
    {
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var person = wb.Persons.Add("Ada Lovelace");

            for (var row = 1; row <= UsedCells; row++)
                ws.Cell(row, 1).Value = row;

            // Spread across rows and columns so a row-major slice walk has to pair each thread
            // with the point it actually sits on.
            ws.Cell("A1").CreateThreadedComment(person, "first");
            ws.Cell("C5").CreateThreadedComment(person, "middle");
            ws.Cell("B20000").CreateThreadedComment(person, "last");

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var sheet = reloaded.Worksheets.First();

        await Assert.That(sheet.Cell("A1").GetThreadedComment()!.Text).IsEqualTo("first");
        await Assert.That(sheet.Cell("C5").GetThreadedComment()!.Text).IsEqualTo("middle");
        await Assert.That(sheet.Cell("B20000").GetThreadedComment()!.Text).IsEqualTo("last");
        await Assert.That(reloaded.Persons.Count).IsEqualTo(1);
        await Assert.That(reloaded.Persons.First().DisplayName).IsEqualTo("Ada Lovelace");
    }
}
