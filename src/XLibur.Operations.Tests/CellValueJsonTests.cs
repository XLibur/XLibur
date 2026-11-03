namespace XLibur.Operations.Tests;

public class CellValueJsonTests
{
    private static readonly ConvertOperation Convert = new();

    [Test]
    public async Task A_tagged_value_of_the_wrong_JSON_type_is_unreadable_input_not_a_defect()
    {
        using var root = new TempDirectory();
        root.File("in.json", "[[{\"date\":123}]]");

        var outcome = Fixtures.Run(Fixtures.Context(root), Convert, new ConvertRequest { Input = "in.json", Output = "out.csv" }).Outcome;

        await Assert.That(outcome.Error!.Code).IsEqualTo("unreadable_table");
    }

    [Test]
    public async Task A_JSON_date_keeps_the_cells_millisecond_precision_through_export_and_import()
    {
        using var root = new TempDirectory();
        root.File("in.json", "[[{\"date\":\"2026-01-31T12:00:00.1234567\"}]]");
        var context = Fixtures.Context(root);

        Fixtures.Run(context, Convert, new ConvertRequest { Input = "in.json", Output = "out.json" });
        Fixtures.Run(context, Convert, new ConvertRequest { Input = "out.json", Output = "again.json" });

        const string expected = "[[{\"date\":\"2026-01-31T12:00:00.123\"}]]";
        await Assert.That(await File.ReadAllTextAsync(root.Combine("out.json"))).IsEqualTo(expected);
        await Assert.That(await File.ReadAllTextAsync(root.Combine("again.json"))).IsEqualTo(expected);
    }
}
