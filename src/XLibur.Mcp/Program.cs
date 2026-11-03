using System.Globalization;
using System.Reflection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using XLibur.Operations;
using XLibur.Operations.Paths;

namespace XLibur.Mcp;

public static class Program
{
    private const string Usage =
        "Usage: xlibur-mcp [--root <dir>]... [--allow-write] [--max-input-mb <n>]\n" +
        "Serves MCP over stdio. Paths are confined to the roots (default: the current directory).";

    public static async Task<int> Main(string[] args)
    {
        if (!ServerOptions.TryParse(args, out var options, out var unknown))
        {
            await Console.Error.WriteLineAsync($"Unknown argument '{unknown}'.\n{Usage}");
            return 2;
        }

        PathPolicy policy;
        try
        {
            policy = PathPolicy.Restricted(options.Roots.Count == 0 ? [Directory.GetCurrentDirectory()] : options.Roots, options.AllowWrite);
        }
        catch (DirectoryNotFoundException e)
        {
            await Console.Error.WriteLineAsync(e.Message);
            return 2;
        }

        var tools = Tools(new OperationContext(policy, options.Limits));

        var builder = Host.CreateApplicationBuilder([]);
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);
        builder.Logging.SetMinimumLevel(LogLevel.Warning);
        builder.Services
            .AddMcpServer(o =>
            {
                o.ServerInfo = new Implementation
                {
                    Name = "xlibur",
                    Version = typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0",
                };
                o.ServerInstructions = Instructions(policy);
            })
            .WithStdioServerTransport()
            .WithTools(tools);

        await builder.Build().RunAsync();
        return 0;
    }

    /// <summary>One tool per operation, sharing one gate so calls run one at a time; a write-off policy gets no writing tool.</summary>
    internal static List<OperationTool> Tools(OperationContext context)
    {
        var gate = new SemaphoreSlim(1, 1);
        return OperationRegistry.All
            .Where(o => context.Paths.AllowWrite || o.Effect == OperationEffect.ReadOnly)
            .Select(o => new OperationTool(o, context, gate))
            .ToList();
    }

    internal static string Instructions(PathPolicy policy)
        => "Excel workbooks (.xlsx, .xlsm) read, edited and recalculated with XLibur's formula engine. " +
           "Call inspect first, then read_range for cells; its 'next' pages a large range. " +
           $"Paths must be inside: {string.Join("; ", policy.Roots!)}. A relative path is relative to the first. " +
           (policy.AllowWrite
               ? "Writes are on: edit saves over the source unless outputPath is given."
               : "Writes are off: edit, convert and render_report are not available. Restart with --allow-write for them.");
}

internal sealed record ServerOptions(List<string> Roots, bool AllowWrite, Limits Limits)
{
    /// <summary>False, with the argument it stopped at, for an unknown argument or a --max-input-mb that is not a positive whole number.</summary>
    internal static bool TryParse(IReadOnlyList<string> args, out ServerOptions options, out string? unknown)
    {
        options = new ServerOptions([], false, Limits.Default with { MaxResultChars = 200_000 });
        var queue = new Queue<string>(args);
        while (queue.TryDequeue(out var arg))
        {
            switch (arg)
            {
                case "--root" when queue.Count > 0:
                    options.Roots.Add(queue.Dequeue());
                    break;
                case "--allow-write":
                    options = options with { AllowWrite = true };
                    break;
                case "--max-input-mb" when queue.TryPeek(out var next) && int.TryParse(next, CultureInfo.InvariantCulture, out var mb) && mb > 0:
                    options = options with { Limits = options.Limits with { MaxInputBytes = mb * 1024L * 1024 } };
                    queue.Dequeue();
                    break;
                default:
                    unknown = arg;
                    return false;
            }
        }

        unknown = null;
        return true;
    }
}
