using XLibur.Cli;

namespace XLibur.Operations.Tests;

public class CliTests
{
    private static (int Exit, string Output, string Error) Run(params string[] args)
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = Program.Run(args, output, error);
        return (exit, output.ToString(), error.ToString());
    }

    [Test]
    public async Task Diff_exits_0_when_identical_1_when_different_and_2_on_error()
    {
        using var root = new TempDirectory();
        var left = Fixtures.Sales(root);
        var right = Fixtures.Sales(root, "other.xlsx");
        Run("edit", right, "--set", "Sales!B2=99");

        await Assert.That(Run("diff", left, left).Exit).IsEqualTo(0);
        await Assert.That(Run("diff", left, right).Exit).IsEqualTo(1);
        await Assert.That(Run("diff", left, root.Combine("missing.xlsx")).Exit).IsEqualTo(2);
    }

    [Test]
    public async Task Options_are_the_request_properties_in_kebab_case_and_positionals_come_first()
    {
        using var root = new TempDirectory();
        var path = Fixtures.Sales(root);

        var (exit, output, _) = Run("read-range", path, "Sales!A1:B2", "--mode", "formulas", "--max-cells", "10");

        await Assert.That(exit).IsEqualTo(0);
        await Assert.That(output).IsEqualTo("Sales!A1:B2\n\tA\tB\n1\tProduct\tQty\n2\tApple\t3\n");
    }

    [Test]
    public async Task A_flag_can_be_turned_off_and_lists_repeat()
    {
        using var root = new TempDirectory();
        var left = Fixtures.Sales(root);
        var right = Fixtures.Sales(root, "other.xlsx");
        Run("edit", right, "--set", "Sales!B2=99", "--set", "Sales!C2=9");

        var ignored = Run("diff", left, right, "--ignore-ranges", "Sales!B2:C2", "--ignore-ranges", "Sales!B5:D5", "--ignore-ranges", "Sales!D2", "--no-styles");
        var json = Run("diff", left, right, "--json", "--styles", "false");

        await Assert.That(ignored.Exit).IsEqualTo(0);
        await Assert.That(json.Output).StartsWith("{\"left\":");
    }

    [Test]
    public async Task A_usage_error_names_the_command_help()
    {
        var unknown = Run("read-range", "x.xlsx", "--colour", "red");
        var missing = Run("diff", "only-one.xlsx");
        var command = Run("frobnicate");

        await Assert.That(unknown.Exit).IsEqualTo(2);
        await Assert.That(unknown.Error).Contains("Unknown option '--colour'");
        await Assert.That(missing.Error).Contains("diff needs <right>");
        await Assert.That(command.Error).Contains("Unknown command 'frobnicate'");
    }

    [Test]
    public async Task Help_lists_every_operation_and_a_command_lists_its_options()
    {
        var overview = Run("--help");
        var edit = Run("edit", "--help");

        foreach (var operation in OperationRegistry.All)
            await Assert.That(overview.Output).Contains(operation.Name.Replace('_', '-'));
        await Assert.That(edit.Output).Contains("--output-path");
        await Assert.That(edit.Output).Contains("--set");
    }

    [Test]
    public async Task A_root_confines_the_CLI_as_it_does_the_server()
    {
        using var root = new TempDirectory();
        using var outside = new TempDirectory();
        var path = Fixtures.Sales(outside);

        var (exit, _, error) = Run("--root", root.Path, "inspect", path);

        await Assert.That(exit).IsEqualTo(2);
        await Assert.That(error).Contains("path_outside_roots");
    }
}
