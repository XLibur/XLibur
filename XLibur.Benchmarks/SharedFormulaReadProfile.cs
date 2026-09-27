using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// What a column of formulas costs when Excel saved it as one shared formula, against the same column
/// saved with a formula in each cell (#686, lever 4).
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile sharedread</c>
/// </summary>
/// <remarks>
/// <para>
/// A cell of a shared formula evaluates the one R1C1 tree of its group instead of parsing its own A1
/// text. That saves the parse on the first read and the tree each cell would keep afterwards. A
/// relative R1C1 reference resolves to a different area in each cell, so it allocates a reference on
/// every evaluation where an A1 tree reuses the one it made first. The recalculation row shows that
/// cost.
/// </para>
/// <para>
/// XLibur writes a formula in each cell, so the shared file is made by rewriting the sheet part, as
/// Excel would have written it. No formula keeps a cached value, so every read evaluates. Bytes are
/// exact.
/// </para>
/// </remarks>
public static class SharedFormulaReadProfile
{
    private const int Rows = 50_000;

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        var perCell = Build(shared: false);
        var shared = Build(shared: true);

        Console.WriteLine();
        Console.WriteLine($"{Rows:N0} rows of SUM(An:En) in column F, no cached values. Bytes per formula.");
        Console.WriteLine();
        Console.WriteLine("  | formulas | load      | first read | retained after read | recalculation |");
        Console.WriteLine("  |----------|-----------|------------|---------------------|---------------|");

        foreach (var (name, file) in new[] { ("per cell", perCell), ("shared", shared) })
        {
            // Warm the paths once so the measured pass does not absorb JIT.
            Measure(file);
            var (load, read, retained, recalc) = Measure(file);
            Console.WriteLine($"  | {name,-8} | {load,9:N0} | {read,10:N0} | {retained,19:N0} | {recalc,13:N0} |");
        }
    }

    private static (long Load, long Read, long Retained, long Recalc) Measure(byte[] file)
    {
#pragma warning disable S1215 // Intentionally forcing GC so each measurement starts from a clean heap
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
#pragma warning restore S1215

        var start = GC.GetTotalAllocatedBytes(precise: true);
        using var workbook = new XLWorkbook(new MemoryStream(file, writable: false));
        var load = GC.GetTotalAllocatedBytes(precise: true) - start;

        var ws = workbook.Worksheet(1);
        if (!ws.Cell(Rows, 6).NeedsRecalculation)
            throw new InvalidOperationException("The formulas kept cached values, so a read would not evaluate them.");

        var heapBefore = GC.GetTotalMemory(forceFullCollection: true);
        var readStart = GC.GetTotalAllocatedBytes(precise: true);
        double checksum = 0;
        for (var row = 1; row <= Rows; row++)
            checksum += ws.Cell(row, 6).GetDouble();

        var read = GC.GetTotalAllocatedBytes(precise: true) - readStart;
        var retained = GC.GetTotalMemory(forceFullCollection: true) - heapBefore;

        // The first pass builds the calculation chain and the dependency tree. The second only
        // evaluates, which is the cost that differs.
        workbook.RecalculateAllFormulas();
        var recalcStart = GC.GetTotalAllocatedBytes(precise: true);
        workbook.RecalculateAllFormulas();
        var recalc = GC.GetTotalAllocatedBytes(precise: true) - recalcStart;

        checksum += ws.Cell(Rows, 6).GetDouble();
        GC.KeepAlive(checksum);
        return (load / Rows, read / Rows, retained / Rows, recalc / Rows);
    }

    private static byte[] Build(bool shared)
    {
        var package = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet("Data");
            for (var row = 1; row <= Rows; row++)
            {
                for (var column = 1; column <= 5; column++)
                    ws.Cell(row, column).Value = row * column;

                ws.Cell(row, 6).FormulaA1 = $"SUM(A{row}:E{row})";
            }

            workbook.SaveAs(package);
        }

        RewriteSheet(package, xml =>
        {
            // No cached values, so every read evaluates.
            xml = Regex.Replace(xml, "(<x:f>[^<]*</x:f>)<x:v>[^<]*</x:v>", "$1");
            if (!shared)
                return xml;

            return Regex.Replace(xml, @"<x:f>SUM\(A(\d+):E\1\)</x:f>", match => match.Groups[1].Value == "1"
                ? $"<x:f t=\"shared\" ref=\"F1:F{Rows}\" si=\"0\">SUM(A1:E1)</x:f>"
                : "<x:f t=\"shared\" si=\"0\" />");
        });

        return package.ToArray();
    }

    private static void RewriteSheet(MemoryStream package, Func<string, string> rewrite)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true);
        var entry = archive.GetEntry("xl/worksheets/sheet1.xml")
                    ?? throw new InvalidOperationException("The package has no xl/worksheets/sheet1.xml.");

        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();

        var rewritten = Encoding.UTF8.GetBytes(rewrite(xml));
        using var write = entry.Open();
        write.SetLength(0);
        write.Write(rewritten, 0, rewritten.Length);
    }
}
