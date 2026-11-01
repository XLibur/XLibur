using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Covers the batched delete's cell compaction, which moves whole rows in one pass instead of moving
/// cells a block at a time.
/// <para>
/// The compaction bypasses the per-cell write path, so everything that path maintained as a side
/// effect has to be maintained here too: per-column usage counts (which back <c>LastColumnUsed</c>),
/// the max row and column, and the shared-string reference counts held by text values. None of that
/// is visible in the cell values themselves, which is why these go looking for it specifically.
/// </para>
/// </summary>
public class BatchRowDeleteCompactionTests
{
    [Test]
    public async Task EveryValueTypeSurvivesTheMove()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A5").Value = "text";
        ws.Cell("B5").Value = 42.5;
        ws.Cell("C5").Value = true;
        ws.Cell("D5").Value = new DateTime(2026, 8, 1);
        ws.Cell("E5").Value = TimeSpan.FromHours(3);
        ws.Cell("F5").FormulaA1 = "1+1";

        ws.Rows("1:1,2:2").Delete();

        await Assert.That(ws.Cell("A3").GetString()).IsEqualTo("text");
        await Assert.That(ws.Cell("B3").GetDouble()).IsEqualTo(42.5);
        await Assert.That(ws.Cell("C3").GetBoolean()).IsTrue();
        await Assert.That(ws.Cell("D3").GetDateTime()).IsEqualTo(new DateTime(2026, 8, 1));
        await Assert.That(ws.Cell("E3").GetTimeSpan()).IsEqualTo(TimeSpan.FromHours(3));
        await Assert.That(ws.Cell("F3").FormulaA1).IsEqualTo("1+1");
    }

    [Test]
    public async Task StylesMoveWithTheirRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A5").Value = "styled";
        ws.Cell("A5").Style.Font.Bold = true;
        ws.Cell("A5").Style.Fill.BackgroundColor = XLColor.Red;

        ws.Rows("1:1,2:2").Delete();

        await Assert.That(ws.Cell("A3").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("A3").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("A5").Style.Font.Bold).IsFalse();
    }

    [Test]
    public async Task CommentsMoveWithTheirRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A5").GetComment().AddText("note");

        ws.Rows("1:1,2:2").Delete();

        await Assert.That(ws.Cell("A3").HasComment).IsTrue();
        await Assert.That(ws.Cell("A3").GetComment().Text).IsEqualTo("note");
        await Assert.That(ws.Cell("A5").HasComment).IsFalse();
    }

    /// <summary>
    /// Per-column usage counts back <c>LastColumnUsed</c>, and the compaction has to decrement them for
    /// the departing rows itself. Deleting the only row that used a far-right column must retire it.
    /// </summary>
    [Test]
    public async Task LastColumnUsedDropsWhenTheOnlyRowUsingItGoes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("H3").Value = "far right";
        ws.Cell("A4").Value = 4;

        await Assert.That(ws.LastColumnUsed()!.ColumnNumber()).IsEqualTo(8);

        ws.Rows("3:3").Delete();

        await Assert.That(ws.LastColumnUsed()!.ColumnNumber()).IsEqualTo(1);
    }

    [Test]
    public async Task LastColumnUsedSurvivesWhenAnotherRowStillUsesIt()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("H3").Value = "far right";
        ws.Cell("H7").Value = "also far right";

        ws.Rows("3:3").Delete();

        await Assert.That(ws.LastColumnUsed()!.ColumnNumber()).IsEqualTo(8);
    }

    [Test]
    public async Task LastRowUsedReflectsTheCompactedSheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 10; row++)
            ws.Cell(row, 1).Value = row;

        ws.Rows("2:2,5:5,9:9").Delete();

        await Assert.That(ws.LastRowUsed()!.RowNumber()).IsEqualTo(7);
        await Assert.That(ws.Cell(8, 1).IsEmpty()).IsTrue();
    }

    [Test]
    public async Task DeletingEveryUsedRowLeavesAnEmptySheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 5; row++)
            ws.Cell(row, 1).Value = row;

        ws.Rows("1:5").Delete();

        await Assert.That(ws.LastRowUsed()).IsNull();
        await Assert.That(ws.LastColumnUsed()).IsNull();
    }

    /// <summary>
    /// Text values hold shared-string references. The rows that go must release theirs and the rows
    /// that move must keep theirs; getting either wrong corrupts the table, which only shows up once
    /// the workbook is written and read back.
    /// </summary>
    [Test]
    public async Task TextValuesRoundTripThroughSaveAfterABatchedDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        for (var row = 1; row <= 12; row++)
            ws.Cell(row, 1).Value = $"value-{row}";

        // "value-3" is shared by a row that goes and one that stays.
        ws.Cell(12, 2).Value = "value-3";

        ws.Rows("3:3,6:6,9:9").Delete();

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        stream.Position = 0;

        using var reloaded = new XLWorkbook(stream);
        var sheet = reloaded.Worksheet("Data");
        var expected = new[] { 1, 2, 4, 5, 7, 8, 10, 11, 12 };
        for (var i = 0; i < expected.Length; i++)
            await Assert.That(sheet.Cell(i + 1, 1).GetString()).IsEqualTo($"value-{expected[i]}");

        await Assert.That(sheet.Cell(9, 2).GetString()).IsEqualTo("value-3");
    }

    /// <summary>
    /// A differential sweep over a sheet carrying values, formulas, styles, merges and a live range.
    /// Each case deletes a pseudo-random set of rows both ways and compares the whole sheet. The seed
    /// is fixed so a failure is reproducible.
    /// </summary>
    [Test]
    [Arguments(1)]
    [Arguments(7)]
    [Arguments(42)]
    [Arguments(1234)]
    public async Task BatchedAndRowAtATimeAgreeOnARichSheet(int seed)
    {
        var random = new Random(seed);
        var targets = new SortedSet<int>();
        while (targets.Count < 12)
            targets.Add(random.Next(2, 40));

        using var perRowBook = BuildRichSheet(out var perRow);
        using var batchedBook = BuildRichSheet(out var batched);

        foreach (var row in targets.Reverse())
            perRow.Row(row).Delete();

        batched.Rows(string.Join(",", targets.Select(r => $"{r}:{r}"))).Delete();

        for (var row = 1; row <= 45; row++)
        {
            for (var column = 1; column <= 4; column++)
            {
                var a = perRow.Cell(row, column);
                var b = batched.Cell(row, column);

                await Assert.That(b.FormulaA1).IsEqualTo(a.FormulaA1);
                await Assert.That(b.GetString()).IsEqualTo(a.GetString());
                await Assert.That(b.Style.Font.Bold).IsEqualTo(a.Style.Font.Bold);
            }
        }

        await Assert.That(batched.LastRowUsed()?.RowNumber()).IsEqualTo(perRow.LastRowUsed()?.RowNumber());
        await Assert.That(batched.LastColumnUsed()?.ColumnNumber()).IsEqualTo(perRow.LastColumnUsed()?.ColumnNumber());
        await Assert.That(batched.MergedRanges.Select(r => r.RangeAddress.ToString()))
            .IsEquivalentTo(perRow.MergedRanges.Select(r => r.RangeAddress.ToString()));
    }

    private static XLWorkbook BuildRichSheet(out IXLWorksheet ws)
    {
        var wb = new XLWorkbook();
        ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= 40; row++)
        {
            ws.Cell(row, 1).Value = $"label-{row}";
            ws.Cell(row, 2).Value = row * 1.5;
            ws.Cell(row, 3).FormulaA1 = $"B{row}*2";

            if (row % 4 == 0)
                ws.Cell(row, 1).Style.Font.Bold = true;
        }

        ws.Cell(1, 4).FormulaA1 = "SUM(B5:B25)";
        ws.Range("A30:B31").Merge();
        return wb;
    }
}
