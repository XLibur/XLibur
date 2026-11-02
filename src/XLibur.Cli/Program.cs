using System.Reflection;
using System.Text;
using XLibur.Operations;
using XLibur.Operations.Paths;

namespace XLibur.Cli;

/// <summary>
/// Exit codes: 0 done; 1 done with a finding (diff found differences, validate failed a check,
/// check-formula refused or could not evaluate, render-report listed template errors); 2 error.
/// </summary>
public static class Program
{
    public static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        return Run(args, Console.Out, Console.Error);
    }

    public static int Run(IReadOnlyList<string> args, TextWriter output, TextWriter error)
    {
        var rest = new List<string>();
        var roots = new List<string>();
        var json = false;
        var help = false;
        for (var i = 0; i < args.Count; i++)
        {
            switch (args[i])
            {
                case "--root" when i + 1 < args.Count:
                    roots.Add(args[++i]);
                    break;
                case "--json":
                    json = true;
                    break;
                case "--help" or "-h" or "help":
                    help = true;
                    break;
                case "--version":
                    output.WriteLine(typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion);
                    return 0;
                default:
                    rest.Add(args[i]);
                    break;
            }
        }

        if (rest.Count == 0)
        {
            (help ? output : error).Write(Overview());
            return help ? 0 : 2;
        }

        var operation = OperationRegistry.Find(rest[0]);
        if (operation is null)
        {
            error.WriteLine($"Unknown command '{rest[0]}'.");
            error.Write(Overview());
            return 2;
        }

        if (help)
        {
            output.Write(RequestBinder.Help(operation));
            return 0;
        }

        PathPolicy policy;
        try
        {
            policy = roots.Count == 0 ? PathPolicy.Unrestricted() : PathPolicy.Restricted(roots, allowWrite: true);
        }
        catch (DirectoryNotFoundException e)
        {
            error.WriteLine(e.Message);
            return 2;
        }

        var format = json ? ResultFormat.Json : ResultFormat.Text;
        object request;
        try
        {
            request = RequestBinder.Bind(operation, rest.Skip(1).ToList());
        }
        catch (OperationException e)
        {
            error.WriteLine(e.Error.ToText());
            error.WriteLine($"See: xlibur {RequestBinder.Command(operation)} --help");
            return 2;
        }

        var outcome = OperationRunner.Run(operation, request, new OperationContext(policy), format, out var rendered);
        var writer = outcome.Error is null ? output : error;
        writer.Write(rendered);
        if (!rendered.EndsWith('\n'))
            writer.WriteLine();
        return outcome.ExitCode;
    }

    private static string Overview()
    {
        var text = new StringBuilder();
        text.Append("xlibur: read, query, edit, recalculate and compare Excel workbooks with XLibur.\n\n");
        text.Append("Usage: xlibur [--root <dir>]... <command> [arguments] [--json]\n\nCommands:\n");
        foreach (var operation in OperationRegistry.All)
        {
            var description = operation.Description;
            var firstSentence = description.IndexOf(". ", StringComparison.Ordinal);
            text.Append("  ").Append(RequestBinder.Command(operation).PadRight(15))
                .Append(firstSentence < 0 ? description : description[..(firstSentence + 1)]).Append('\n');
        }

        text.Append("\n--root confines every path to the given directories, as the MCP server does.\n");
        text.Append("Run 'xlibur <command> --help' for a command's options.\n");
        return text.ToString();
    }
}
