using System.Text.Json;
using ModelContextProtocol.Protocol;
using XLibur.Mcp;
using XLibur.Operations.Paths;

namespace XLibur.Operations.Tests;

public class McpServerTests
{
    private static Dictionary<string, JsonElement> Arguments(string json)
        => JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json)!;

    private static OperationTool Tool(OperationContext context, string name)
        => Program.Tools(context).Single(t => t.ProtocolTool.Name == name);

    private static string Text(CallToolResult result) => ((TextContentBlock)result.Content.Single()).Text;

    [Test]
    public async Task A_server_without_write_access_offers_no_tool_that_writes()
    {
        using var root = new TempDirectory();

        var readOnly = Program.Tools(Fixtures.Context(root, allowWrite: false)).Select(t => t.ProtocolTool.Name).ToList();
        var writable = Program.Tools(Fixtures.Context(root)).Select(t => t.ProtocolTool.Name).ToList();

        await Assert.That(readOnly).DoesNotContain("edit");
        await Assert.That(readOnly).DoesNotContain("convert");
        await Assert.That(readOnly).DoesNotContain("render_report");
        await Assert.That(readOnly).Contains("inspect");
        await Assert.That(writable).IsEquivalentTo(OperationRegistry.All.Select(o => o.Name));
    }

    [Test]
    public async Task A_tool_is_annotated_by_its_effect_and_its_schema_takes_a_format()
    {
        using var root = new TempDirectory();
        var context = Fixtures.Context(root);

        var inspect = Tool(context, "inspect").ProtocolTool;
        var edit = Tool(context, "edit").ProtocolTool;
        var properties = inspect.InputSchema.GetProperty("properties");

        await Assert.That(inspect.Annotations!.ReadOnlyHint).IsTrue();
        await Assert.That(inspect.Annotations.DestructiveHint).IsFalse();
        await Assert.That(edit.Annotations!.ReadOnlyHint).IsFalse();
        await Assert.That(edit.Annotations.DestructiveHint).IsTrue();
        await Assert.That(properties.TryGetProperty("path", out _)).IsTrue();
        await Assert.That(properties.GetProperty("format").GetProperty("enum").GetArrayLength()).IsEqualTo(2);
    }

    [Test]
    public async Task A_call_runs_the_operation_and_renders_text_or_JSON()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        var tool = Tool(Fixtures.Context(root), "read_range");

        var text = await tool.CallAsync(Arguments("""{"path":"sales.xlsx","range":"Sales!A1:B2"}"""), CancellationToken.None);
        var json = await tool.CallAsync(Arguments("""{"path":"sales.xlsx","range":"Sales!A1:B2","format":"json"}"""), CancellationToken.None);

        await Assert.That(text.IsError).IsFalse();
        await Assert.That(Text(text)).IsEqualTo("Sales!A1:B2\n\tA\tB\n1\tProduct\tQty\n2\tApple\t3\n");
        await Assert.That(Text(json)).StartsWith("{\"sheet\":\"Sales\",");
    }

    [Test]
    public async Task Bad_arguments_and_refused_paths_come_back_as_tool_errors()
    {
        using var root = new TempDirectory();
        using var outside = new TempDirectory();
        Fixtures.Sales(outside);
        var tool = Tool(Fixtures.Context(root), "inspect");

        var unknown = await tool.CallAsync(Arguments("""{"path":"a.xlsx","colour":"red"}"""), CancellationToken.None);
        var format = await tool.CallAsync(Arguments("""{"path":"a.xlsx","format":"xml"}"""), CancellationToken.None);
        var missing = await tool.CallAsync(null, CancellationToken.None);
        var refused = await tool.CallAsync(Arguments(JsonSerializer.Serialize(new { path = outside.Combine("sales.xlsx") })), CancellationToken.None);

        foreach (var result in new[] { unknown, format, missing, refused })
            await Assert.That(result.IsError).IsTrue();
        await Assert.That(Text(unknown)).StartsWith("error[bad_arguments]");
        await Assert.That(Text(format)).IsEqualTo("error[bad_arguments] BadInput: format is \"text\" or \"json\".");
        await Assert.That(Text(missing)).StartsWith("error[bad_arguments]");
        await Assert.That(Text(refused)).StartsWith("error[path_outside_roots]");
    }

    [Test]
    public async Task Server_options_take_roots_write_access_and_an_input_limit()
    {
        var parsed = ServerOptions.TryParse(["--root", "a", "--allow-write", "--root", "b", "--max-input-mb", "5"], out var options, out _);

        await Assert.That(parsed).IsTrue();
        await Assert.That(options.Roots).IsEquivalentTo(["a", "b"]);
        await Assert.That(options.AllowWrite).IsTrue();
        await Assert.That(options.Limits.MaxInputBytes).IsEqualTo(5L * 1024 * 1024);
        await Assert.That(options.Limits.MaxResultChars).IsEqualTo(200_000);
    }

    [Test]
    [Arguments("--bogus", "--bogus")]
    [Arguments("--max-input-mb 0", "--max-input-mb")]
    [Arguments("--max-input-mb lots", "--max-input-mb")]
    [Arguments("--root", "--root")]
    public async Task An_unknown_or_incomplete_server_argument_is_named(string args, string expected)
    {
        var parsed = ServerOptions.TryParse(args.Split(' '), out _, out var unknown);

        await Assert.That(parsed).IsFalse();
        await Assert.That(unknown).IsEqualTo(expected);
    }

    [Test]
    public async Task The_instructions_name_the_roots_and_whether_writes_are_on()
    {
        using var root = new TempDirectory();

        var readOnly = Program.Instructions(PathPolicy.Restricted([root.Path], allowWrite: false));
        var writable = Program.Instructions(PathPolicy.Restricted([root.Path], allowWrite: true));

        await Assert.That(readOnly).Contains("Writes are off");
        await Assert.That(writable).Contains("Writes are on");
        await Assert.That(writable).Contains(PathPolicy.Restricted([root.Path], allowWrite: true).Roots![0]);
    }
}
