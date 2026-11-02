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
        var roots = new List<string>();
        var allowWrite = false;
        var limits = Limits.Default with { MaxResultChars = 200_000 };
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--root" when i + 1 < args.Length:
                    roots.Add(args[++i]);
                    break;
                case "--allow-write":
                    allowWrite = true;
                    break;
                case "--max-input-mb" when i + 1 < args.Length && int.TryParse(args[i + 1], CultureInfo.InvariantCulture, out var mb) && mb > 0:
                    limits = limits with { MaxInputBytes = mb * 1024L * 1024 };
                    i++;
                    break;
                default:
                    await Console.Error.WriteLineAsync($"Unknown argument '{args[i]}'.\n{Usage}");
                    return 2;
            }
        }

        PathPolicy policy;
        try
        {
            policy = PathPolicy.Restricted(roots.Count == 0 ? [Directory.GetCurrentDirectory()] : roots, allowWrite);
        }
        catch (DirectoryNotFoundException e)
        {
            await Console.Error.WriteLineAsync(e.Message);
            return 2;
        }

        var context = new OperationContext(policy, limits);
        var gate = new SemaphoreSlim(1, 1);
        var tools = OperationRegistry.All
            .Where(o => allowWrite || o.Effect == OperationEffect.ReadOnly)
            .Select(o => new OperationTool(o, context, gate))
            .ToList();

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

    private static string Instructions(PathPolicy policy)
        => "Excel workbooks (.xlsx, .xlsm) read, edited and recalculated with XLibur's formula engine. " +
           "Call inspect first, then read_range for cells; its 'next' pages a large range. " +
           $"Paths must be inside: {string.Join("; ", policy.Roots!)}. A relative path is relative to the first. " +
           (policy.AllowWrite
               ? "Writes are on: edit saves over the source unless outputPath is given."
               : "Writes are off: edit, convert and render_report are not available. Restart with --allow-write for them.");
}
