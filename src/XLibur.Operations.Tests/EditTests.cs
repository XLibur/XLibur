using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class EditTests
{
    private static readonly EditOperation Operation = new();

    [Test]
    public async Task An_edit_returns_computed_values_new_errors_and_changed_dependents()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var (outcome, _) = Fixtures.Run(Fixtures.Context(root), Operation, new EditRequest
        {
            Path = "sales.xlsx",
            Edits =
            [
                new CellEdit { Cell = "Sales!B2", Value = 10 },
                new CellEdit { Cell = "E2", Formula = "=D2*2" },
                new CellEdit { Cell = "C4", Formula = "1/0" },
            ],
        });

        var result = (EditResult)outcome.Result!;
        await Assert.That(result.Edited.Select(e => (e.Cell, e.Value))).IsEquivalentTo(
            [("Sales!B2", "10"), ("Sales!E2", "30"), ("Sales!C4", "#DIV/0!")]);
        await Assert.That(result.NewIssues.Select(i => i.Cell)).IsEquivalentTo(["Sales!C4", "Sales!D4", "Sales!D5"]);
        await Assert.That(result.Changed.Select(c => (c.Cell, c.Before, c.After))).IsEquivalentTo(
            [("Sales!D2", "4.5", "15"), ("Sales!B5", "55", "62")]);

        using var saved = new XLWorkbook(root.Combine("sales.xlsx"));
        await Assert.That(saved.Worksheet("Sales").Cell("E2").CachedValue).IsEqualTo((XLCellValue)30);
    }

    [Test]
    public async Task Max_changes_caps_the_listed_dependents_but_not_their_count()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root), Operation, new EditRequest
        {
            Path = "sales.xlsx",
            Edits = [new CellEdit { Cell = "Sales!B2", Value = 10 }],
            MaxChanges = 0,
        });

        var result = (EditResult)outcome.Result!;
        await Assert.That(result.Changed.Count).IsEqualTo(0);
        await Assert.That(result.ChangedTotal).IsEqualTo(3);
        await Assert.That(text).EndsWith("\n3 more changed\n");
    }

    [Test]
    public async Task A_shorthand_edit_without_an_equals_sign_is_bad_input()
    {
        await Expect.Refused(() => CellEdit.Parse("B2"), ErrorKind.BadInput, "bad_edit");
    }

    [Test]
    public async Task One_bad_edit_applies_none_and_writes_nothing()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        var before = await File.ReadAllBytesAsync(path);

        var (outcome, _) = Fixtures.Run(Fixtures.Context(root), Operation, new EditRequest
        {
            Path = "sales.xlsx",
            Edits = [new CellEdit { Cell = "B2", Value = 10 }, new CellEdit { Cell = "Nope!B2", Value = 1 }],
        });

        await Assert.That(outcome.Error!.Code).IsEqualTo("sheet_not_found");
        await Assert.That(await File.ReadAllBytesAsync(path)).IsEquivalentTo(before);
    }

    [Test]
    public async Task A_formula_XLibur_refuses_is_rejected_before_anything_changes()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        var before = await File.ReadAllBytesAsync(path);

        var (outcome, _) = Fixtures.Run(Fixtures.Context(root), Operation, new EditRequest
        {
            Path = "sales.xlsx",
            Edits = [new CellEdit { Cell = "B2", Formula = "SUM(B3:B4" }],
        });

        await Assert.That(outcome.Error!.Code).IsEqualTo("formula_refused");
        await Assert.That(await File.ReadAllBytesAsync(path)).IsEquivalentTo(before);
    }

    [Test]
    public async Task An_output_path_leaves_the_source_untouched_and_needs_overwrite_to_replace()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);
        var before = await File.ReadAllBytesAsync(path);
        var context = Fixtures.Context(root);
        var request = new EditRequest { Path = "sales.xlsx", OutputPath = "copy.xlsx", Edits = [new CellEdit { Cell = "B2", Clear = true }] };

        var first = Fixtures.Run(context, Operation, request).Outcome;
        var second = Fixtures.Run(context, Operation, request).Outcome;
        var third = Fixtures.Run(context, Operation, request with { Overwrite = true }).Outcome;

        await Assert.That(first.Error).IsNull();
        await Assert.That(second.Error!.Code).IsEqualTo("file_exists");
        await Assert.That(third.Error).IsNull();
        await Assert.That(await File.ReadAllBytesAsync(path)).IsEquivalentTo(before);
        using var copy = new XLWorkbook(root.Combine("copy.xlsx"));
        await Assert.That(copy.Worksheet("Sales").Cell("B2").IsEmpty()).IsTrue();
    }

    [Test]
    public async Task With_writes_off_an_edit_is_forbidden()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);

        var (outcome, _) = Fixtures.Run(Fixtures.Context(root, allowWrite: false), Operation,
            new EditRequest { Path = "sales.xlsx", Edits = [new CellEdit { Cell = "B2", Value = 1 }] });

        await Assert.That(outcome.Error!.Kind).IsEqualTo(ErrorKind.Forbidden);
        await Assert.That(outcome.Error.Code).IsEqualTo("writes_disabled");
    }

    [Test]
    [Arguments("B2=42", "B2", 42d, null, false)]
    [Arguments("B2==SUM(A1:A9)", "B2", null, "SUM(A1:A9)", false)]
    [Arguments("Sales!B2=", "Sales!B2", null, null, true)]
    public async Task The_CLI_shorthand_reads_value_formula_and_clear(string text, string cell, double? number, string? formula, bool clear)
    {
        var edit = CellEdit.Parse(text);

        await Assert.That(edit.Cell).IsEqualTo(cell);
        await Assert.That(edit.Value).IsEqualTo(number is { } n ? (XLCellValue?)n : null);
        await Assert.That(edit.Formula).IsEqualTo(formula);
        await Assert.That(edit.Clear).IsEqualTo(clear);
    }

    [Test]
    public async Task The_CLI_shorthand_keeps_text_as_text()
    {
        await Assert.That(CellEdit.Parse("A1=TRUE").Value!.Value.GetBoolean()).IsTrue();
        await Assert.That(CellEdit.Parse("A1=hello world").Value).IsEqualTo((XLCellValue)"hello world");
    }
}
