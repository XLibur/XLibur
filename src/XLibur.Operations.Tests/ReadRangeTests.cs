using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class ReadRangeTests
{
    private static readonly ReadRangeOperation Operation = new();

    [Test]
    public async Task The_text_form_is_TSV_with_column_letters_and_row_numbers()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root), Operation, new ReadRangeRequest { Path = "sales.xlsx", Mode = ReadMode.Both });

        await Assert.That(outcome.Error).IsNull();
        await Assert.That(text).IsEqualTo(
            "Sales!A1:D5\n" +
            "\tA\tB\tC\tD\n" +
            "1\tProduct\tQty\tPrice\tTotal\n" +
            "2\tApple\t3\t1.5\t=B2*C2 -> 4.5\n" +
            "3\tBanana\t12\t0.25\t=B3*C3 -> 3\n" +
            "4\tCherry\t40\t0.1\t=B4*C4 -> 4\n" +
            "5\tTotal\t=SUM(B2:B4) -> 55\t\t=SUM(D2:D4) -> 11.5\n");
    }

    [Test]
    public async Task A_defined_name_is_a_range()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var result = (ReadRangeResult)Fixtures.Run(Fixtures.Context(root), Operation,
            new ReadRangeRequest { Path = "sales.xlsx", Range = "GrandTotal" }).Outcome.Result!;

        await Assert.That(result.Range).IsEqualTo("Sales!D5");
        await Assert.That(result.Values[0][0]).IsEqualTo((XLCellValue)11.5);
    }

    [Test]
    public async Task A_quoted_sheet_name_a_sheet_scoped_name_and_an_empty_sheet_all_read()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "names.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("It's Q1");
            ws.Cell("A1").Value = "a";
            ws.Cell("B2").Value = 2;
            ws.DefinedNames.Add("Corner", "'It''s Q1'!$B$2");
            wb.AddWorksheet("Blank");
        });
        var context = Fixtures.Context(root);

        var quoted = Fixtures.Run(context, Operation, new ReadRangeRequest { Path = "names.xlsx", Range = "'It''s Q1'!A1:B2" });
        var local = (ReadRangeResult)Fixtures.Run(context, Operation,
            new ReadRangeRequest { Path = "names.xlsx", Sheet = "It's Q1", Range = "Corner" }).Outcome.Result!;
        var empty = Fixtures.Run(context, Operation, new ReadRangeRequest { Path = "names.xlsx", Sheet = "Blank" });

        await Assert.That(quoted.Text).IsEqualTo("'It''s Q1'!A1:B2\n\tA\tB\n1\ta\n2\t\t2\n");
        await Assert.That(local.Values[0][0]).IsEqualTo((XLCellValue)2d);
        await Assert.That(empty.Text).IsEqualTo("Blank: empty\n");
    }

    [Test]
    public async Task A_large_range_is_paged_by_whole_rows_and_next_reads_the_rest()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "big.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            for (var r = 1; r <= 100; r++)
            for (var c = 1; c <= 3; c++)
                ws.Cell(r, c).Value = r * 10 + c;
        });
        var context = Fixtures.Context(root);

        var rows = new List<IReadOnlyList<XLCellValue>>();
        string? range = null;
        var pages = 0;
        do
        {
            var page = (ReadRangeResult)Fixtures.Run(context, Operation,
                new ReadRangeRequest { Path = "big.xlsx", Range = range, MaxCells = 30 }).Outcome.Result!;
            rows.AddRange(page.Values);
            range = page.Next;
            pages++;
            if (pages == 1)
                await Assert.That(page.Next).IsEqualTo("Data!A11:C100");
        } while (range is not null);

        await Assert.That(pages).IsEqualTo(10);
        await Assert.That(rows.Count).IsEqualTo(100);
        await Assert.That(rows[99][2]).IsEqualTo((XLCellValue)1003);
    }

    [Test]
    public async Task The_host_cell_cap_wins_over_the_request()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "big.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            for (var r = 1; r <= 50; r++)
                ws.Cell(r, 1).Value = r;
        });
        var context = Fixtures.Context(root, limits: Limits.Default with { MaxCells = 10 });

        var page = (ReadRangeResult)Fixtures.Run(context, Operation,
            new ReadRangeRequest { Path = "big.xlsx", MaxCells = 1000 }).Outcome.Result!;

        await Assert.That(page.Values.Count).IsEqualTo(10);
        await Assert.That(page.Next).IsEqualTo("Data!A11:A50");
    }

    [Test]
    public async Task A_whole_column_is_cut_to_the_used_range()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var page = (ReadRangeResult)Fixtures.Run(Fixtures.Context(root), Operation,
            new ReadRangeRequest { Path = "sales.xlsx", Range = "B:B" }).Outcome.Result!;

        await Assert.That(page.Range).IsEqualTo("Sales!B1:B5");
        await Assert.That(page.Next).IsNull();
    }

    [Test]
    public async Task A_circular_reference_is_marked_not_thrown()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "cycle.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").FormulaA1 = "A1+1";
        });
        using (var wb = new XLWorkbook(root.Combine("cycle.xlsx")))
        {
            wb.Worksheet("S").Cell("A1").FormulaA1 = "B1";
            wb.Save();
        }

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root), Operation,
            new ReadRangeRequest { Path = "cycle.xlsx", Recalculate = true });

        await Assert.That(outcome.Error).IsNull();
        await Assert.That(text).Contains("1\t!circular\t!circular");
    }

    [Test]
    public async Task Text_with_tabs_and_newlines_is_escaped_so_a_row_stays_one_line()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "text.xlsx", wb => wb.AddWorksheet("S").Cell("A1").Value = "a\tb\nc\\d");

        var (_, text) = Fixtures.Run(Fixtures.Context(root), Operation, new ReadRangeRequest { Path = "text.xlsx" });

        await Assert.That(text).Contains("1\ta\\tb\\nc\\\\d\n");
    }

    [Test]
    public async Task The_JSON_form_tags_error_values()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "err.xlsx", wb => wb.AddWorksheet("S").Cell("A1").FormulaA1 = "1/0");

        var (_, json) = Fixtures.Run(Fixtures.Context(root), Operation, new ReadRangeRequest { Path = "err.xlsx" }, ResultFormat.Json);

        await Assert.That(json).Contains("\"values\":[[{\"error\":\"#DIV/0!\"}]]");
    }

    [Test]
    public async Task A_missing_sheet_names_the_sheets_there_are()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var (outcome, _) = Fixtures.Run(Fixtures.Context(root), Operation, new ReadRangeRequest { Path = "sales.xlsx", Range = "Nope!A1" });

        await Assert.That(outcome.Error!.Code).IsEqualTo("sheet_not_found");
        await Assert.That(outcome.Error.Message).Contains("Sales");
    }
}
