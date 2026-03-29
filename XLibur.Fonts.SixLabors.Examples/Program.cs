using System;
using System.IO;
using XLibur.Fonts.SixLabors.Examples.FontEngine;

namespace XLibur.Fonts.SixLabors.Examples;

public static class Program
{
    public static string OutputDirectory
    {
        get
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }
    }

    private static void Main()
    {
        Console.WriteLine("Running SixLabors.Fonts v2 examples...");

        Console.WriteLine("  UsingSixLaborsFontsV2...");
        UsingSixLaborsFontsV2.Create(Path.Combine(OutputDirectory, "UsingSixLaborsFontsV2.xlsx"));

        Console.WriteLine("Done. Output in: " + OutputDirectory);
    }
}
