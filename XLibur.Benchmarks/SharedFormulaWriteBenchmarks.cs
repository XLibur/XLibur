using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// What writing back loaded shared formulas costs and saves: a round trip of a file whose formula
/// columns are shared as Excel shares them, and a load of the file XLibur writes, each with and without
/// shared formulas (ADR 0003, #692).
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter *SharedFormulaWrite*</c>
/// </summary>
/// <remarks>
/// <para>
/// The input is built by XLibur and rewritten into the shape Excel saves: each column of filled-down
/// formulas becomes shared formulas of 64 rows, which is the most Excel puts in one. A formula with a
/// reference to another sheet or a mixed range such as <c>$H$2:H2</c> stays a formula of its own,
/// because Excel never shares either. Every formula keeps a cached value, as in a file from Excel.
/// </para>
/// <para>
/// <see cref="SaveOptions.RewriteUnchangedSheets"/> is set so the sheet is always written from the
/// model, whatever a later fast path does with an unchanged sheet.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class SharedFormulaWriteBenchmarks
{
    private const int GroupRows = 64;

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(30);

    private byte[] _input = null!;
    private byte[] _output = null!;

    [Params("FillDown50k", "Orders10k")]
    public string Fixture { get; set; } = null!;

    [Params(true, false)]
    public bool Shared { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _input = Fixture switch
        {
            "FillDown50k" => BuildFillDown(50_000),
            "Orders10k" => BuildOrders(10_000),
            _ => throw new ArgumentOutOfRangeException(nameof(Fixture), Fixture, null)
        };

        _output = Save(_input, Shared);
    }

    /// <summary>Load the file with shared formulas and save it, in the mode under test.</summary>
    [Benchmark]
    public long RoundTrip() => Save(_input, Shared).LongLength;

    /// <summary>Load the file that the round trip writes in the mode under test.</summary>
    [Benchmark]
    public int LoadOutput()
    {
        using var workbook = new XLWorkbook(new MemoryStream(_output, writable: false));
        return workbook.Worksheets.Count;
    }

    /// <summary>
    /// Exact sizes and allocation for each fixture and mode: the save alone, the load of what it
    /// wrote, and the first recalculation after that load.
    /// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile sharedwrite</c>
    /// </summary>
    public static void Profile()
    {
        SixLaborsV1FontBootstrap.Register();
        Console.WriteLine("| fixture | mode | file bytes | sheet XML bytes | save alloc | load alloc | first recalc alloc |");
        Console.WriteLine("|---|---|---:|---:|---:|---:|---:|");
        foreach (var (name, input) in new[] { ("FillDown50k", BuildFillDown(50_000)), ("Orders10k", BuildOrders(10_000)) })
        {
            foreach (var shared in new[] { true, false })
            {
                // The second pass is the one reported, so JIT and first-use costs are left out.
                long save = 0, load = 0, recalc = 0;
                byte[] output = [];
                for (var pass = 0; pass < 2; pass++)
                {
                    using (var workbook = new XLWorkbook(new MemoryStream(input, writable: false)))
                    {
                        var stream = new MemoryStream(capacity: 16 << 20);
                        var start = GC.GetAllocatedBytesForCurrentThread();
                        workbook.SaveAs(stream, new SaveOptions { RewriteUnchangedSheets = true, WriteSharedFormulas = shared });
                        save = GC.GetAllocatedBytesForCurrentThread() - start;
                        output = stream.ToArray();
                    }

                    var loadStart = GC.GetAllocatedBytesForCurrentThread();
                    using var reloaded = new XLWorkbook(new MemoryStream(output, writable: false));
                    load = GC.GetAllocatedBytesForCurrentThread() - loadStart;
                    var recalcStart = GC.GetAllocatedBytesForCurrentThread();
                    reloaded.RecalculateAllFormulas();
                    recalc = GC.GetAllocatedBytesForCurrentThread() - recalcStart;
                }

                long sheetBytes = 0;
                using (var archive = new ZipArchive(new MemoryStream(output), ZipArchiveMode.Read))
                {
                    foreach (var entry in archive.Entries)
                    {
                        if (entry.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal))
                            sheetBytes += entry.Length;
                    }
                }

                Console.WriteLine($"| {name} | {(shared ? "shared" : "expanded")} | {output.Length:N0} | {sheetBytes:N0} | {save:N0} | {load:N0} | {recalc:N0} |");
            }
        }
    }

    /// <summary>
    /// Load the file that the round trip writes and recalculate every formula, which reads each cell of
    /// a shared formula through the one tree of its group (#686).
    /// </summary>
    [Benchmark]
    public int LoadOutputAndRecalculate()
    {
        using var workbook = new XLWorkbook(new MemoryStream(_output, writable: false));
        workbook.RecalculateAllFormulas();
        return workbook.Worksheets.Count;
    }

    internal static byte[] Save(byte[] input, bool shared)
    {
        using var workbook = new XLWorkbook(new MemoryStream(input, writable: false));
        using var output = new MemoryStream();
        workbook.SaveAs(output, new SaveOptions { RewriteUnchangedSheets = true, WriteSharedFormulas = shared });
        return output.ToArray();
    }

    /// <summary>A sheet of numbers in A:E and <c>SUM(An:En)</c> filled down column F.</summary>
    internal static byte[] BuildFillDown(int rows)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");
        for (var row = 1; row <= rows; row++)
        {
            for (var column = 1; column <= 5; column++)
                ws.Cell(row, column).Value = row * column;

            ws.Cell(row, 6).FormulaA1 = $"SUM(A{row}:E{row})";
        }

        return ShareAsExcelDoes(SaveWithValues(workbook), "xl/worksheets/sheet1.xml", ["F"], firstRow: 1);
    }

    /// <summary>
    /// An order list with formula columns filled down, two of them lookups into a product sheet, and a
    /// running total over a mixed range.
    /// </summary>
    internal static byte[] BuildOrders(int rows)
    {
        using var workbook = new XLWorkbook();
        var orders = workbook.AddWorksheet("Orders");
        var products = workbook.AddWorksheet("Products");
        products.Cell(1, 1).Value = "Code";
        products.Cell(1, 2).Value = "Name";
        products.Cell(1, 3).Value = "Price";
        for (var i = 1; i <= 50; i++)
        {
            products.Cell(i + 1, 1).Value = $"P{i:000}";
            products.Cell(i + 1, 2).Value = $"Product {i}";
            products.Cell(i + 1, 3).Value = Math.Round(5 + (i * 37 % 200) / 3.0, 2);
        }

        string[] headers = ["Date", "Code", "Qty", "Name", "Price", "Net", "Tax", "Total", "Band", "Running"];
        for (var c = 0; c < headers.Length; c++)
            orders.Cell(1, c + 1).Value = headers[c];

        for (var row = 2; row <= rows + 1; row++)
        {
            orders.Cell(row, 1).Value = new DateTime(2026, 1, 1).AddDays(row * 7 % 365);
            orders.Cell(row, 2).Value = $"P{row * 13 % 50 + 1:000}";
            orders.Cell(row, 3).Value = row * 17 % 40 + 1;
            orders.Cell(row, 4).FormulaA1 = $"VLOOKUP(B{row},Products!$A$2:$C$51,2,FALSE)";
            orders.Cell(row, 5).FormulaA1 = $"VLOOKUP(B{row},Products!$A$2:$C$51,3,FALSE)";
            orders.Cell(row, 6).FormulaA1 = $"C{row}*E{row}";
            orders.Cell(row, 7).FormulaA1 = $"ROUND(F{row}*0.1,2)";
            orders.Cell(row, 8).FormulaA1 = $"F{row}+G{row}";
            orders.Cell(row, 9).FormulaA1 = $"IF(H{row}>=1000,\"Large\",IF(H{row}>=250,\"Medium\",\"Small\"))";
            orders.Cell(row, 10).FormulaA1 = $"SUM($H$2:H{row})";
        }

        return ShareAsExcelDoes(SaveWithValues(workbook), "xl/worksheets/sheet1.xml", ["F", "G", "H", "I"], firstRow: 2);
    }

    private static byte[] SaveWithValues(XLWorkbook workbook)
    {
        using var stream = new MemoryStream();
        workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        return stream.ToArray();
    }

    /// <summary>
    /// Rewrite the formulas of <paramref name="columns"/> into shared formulas of 64 rows, numbered in
    /// the order their masters appear, as Excel writes a filled-down column.
    /// </summary>
    private static byte[] ShareAsExcelDoes(byte[] package, string sheetPart, string[] columns, int firstRow)
    {
        var shared = new HashSet<string>(columns, StringComparer.Ordinal);
        var lastRow = new Dictionary<string, int>(StringComparer.Ordinal);
        var cellPattern = new Regex("<x:c r=\"([A-Z]+)(\\d+)\"([^>]*)><x:f>([^<]*)</x:f>", RegexOptions.None, RegexTimeout);

        return RewritePart(package, sheetPart, xml =>
        {
            foreach (Match match in cellPattern.Matches(xml))
            {
                var column = match.Groups[1].Value;
                if (shared.Contains(column))
                    lastRow[column] = Math.Max(lastRow.GetValueOrDefault(column), int.Parse(match.Groups[2].Value));
            }

            var indexes = new Dictionary<(string Column, int Block), int>();
            return cellPattern.Replace(xml, match =>
            {
                var column = match.Groups[1].Value;
                if (!shared.Contains(column))
                    return match.Value;

                var row = int.Parse(match.Groups[2].Value);
                var block = (row - firstRow) / GroupRows;
                var head = $"<x:c r=\"{column}{row}\"{match.Groups[3].Value}>";
                if (!indexes.TryGetValue((column, block), out var index))
                {
                    index = indexes.Count;
                    indexes.Add((column, block), index);
                    var end = Math.Min(firstRow + (block + 1) * GroupRows - 1, lastRow[column]);
                    return $"{head}<x:f t=\"shared\" ref=\"{column}{row}:{column}{end}\" si=\"{index}\">{match.Groups[4].Value}</x:f>";
                }

                // Spelled as Excel spells it, so the file is not byte for byte what XLibur writes.
                // Otherwise a save that writes the same shared formulas back keeps the sheet as the
                // file had it, and measures that shortcut instead of the writer.
                return $"{head}<x:f t=\"shared\" si=\"{index}\"/>";
            });
        });
    }

    private static byte[] RewritePart(byte[] package, string partName, Func<string, string> rewrite)
    {
        var stream = new MemoryStream();
        stream.Write(package);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(partName)
                        ?? throw new InvalidOperationException($"The package has no {partName}.");

            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();

            var rewritten = Encoding.UTF8.GetBytes(rewrite(xml));
            using var write = entry.Open();
            write.SetLength(0);
            write.Write(rewritten, 0, rewritten.Length);
        }

        return stream.ToArray();
    }
}
