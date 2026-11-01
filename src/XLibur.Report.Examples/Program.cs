using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace XLibur.Report.Examples;

/// <summary>
/// Runs the worked examples, writing every template and the report generated from it.
/// </summary>
/// <remarks>
/// <code>
/// dotnet run --project XLibur.Report.Examples                      # all of them
/// dotnet run --project XLibur.Report.Examples -- AnnualSalesReport # one of them
/// dotnet run --project XLibur.Report.Examples -- --list
/// dotnet run --project XLibur.Report.Examples -- --out c:\temp\reports
/// </code>
/// </remarks>
public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Contains("--list"))
        {
            List();
            return 0;
        }

        var directory = Argument(args, "--out")
            ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports");

        var names = ExampleNames(args);
        var chosen = names.Count == 0
            ? AllExamples.Ordered
            : names.Select(AllExamples.ByName).OfType<ReportExample>().ToList();

        var unknown = names.Where(name => AllExamples.ByName(name) is null).ToList();
        if (unknown.Count > 0)
        {
            Console.Error.WriteLine("No such example: " + string.Join(", ", unknown));
            List();
            return 1;
        }

        Console.WriteLine("Writing to " + directory);
        Console.WriteLine();

        var failed = 0;

        foreach (var example in chosen)
        {
            var run = example.Run(directory, Console.Out);

            // Every example here should generate cleanly, with the one exception that exists to show
            // what an error looks like.
            if (run.Result.HasErrors && example is not ErrorsAreReportedNotThrown)
            {
                failed++;
            }
        }

        Console.WriteLine("Open each template beside the report generated from it: the template is where");
        Console.WriteLine("the interesting part of every one of these lives.");

        return failed == 0 ? 0 : 1;
    }

    private static void List()
    {
        Console.WriteLine("Examples, simplest first:");
        Console.WriteLine();

        var width = AllExamples.Ordered.Max(example => example.Name.Length);

        foreach (var example in AllExamples.Ordered)
        {
            Console.WriteLine($"  {example.Name.PadRight(width)}  {example.Summary}");
        }
    }

    private static string? Argument(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    /// <summary>
    /// The example names among the arguments: everything that is neither a flag nor the value of one.
    /// </summary>
    private static List<string> ExampleNames(string[] args)
    {
        var names = new List<string>();

#pragma warning disable S127 // --out consumes the following argument, so the index advances twice
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i].StartsWith("--", StringComparison.Ordinal))
            {
                // --out takes a value; skip it so a directory is not mistaken for an example.
                if (args[i] == "--out")
                {
                    i++;
                }

                continue;
            }

            names.Add(args[i]);
        }
#pragma warning restore S127

        return names;
    }
}
