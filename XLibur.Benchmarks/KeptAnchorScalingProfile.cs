using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// Whether loading and saving pictures whose two-cell anchor is kept (<c>editAs="oneCell"</c> or
/// <c>"absolute"</c>) scales with the number of pictures plus the number of rows, or with their product.
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile keptanchors</c>
///
/// <c>profile keptanchors 2000 12500</c> measures one sheet of 2,000 pictures on 12,500 rows instead,
/// so the cost can be compared against the number of pictures (#797).
/// </summary>
/// <remarks>
/// <para>
/// Issue #785. A kept anchor records the sizes of the rows and columns under it at load and compares
/// them at save, so that a resized row makes the writer regenerate the anchor. Each picture used to
/// walk the sheet's whole row and column collections to do it, at load and again at save, so 1,000
/// pictures on a sheet of 50,000 sized rows walked 50 million rows each way.
/// </para>
/// <para>
/// Every row of the sheet has a set height, so every row is one the comparison has to account for:
/// the worst case for the walk. The 1,000 pictures are spread down the whole sheet, so an absolute
/// picture's rows, which run from the top of the sheet to its bottom-right corner, grow with the sheet.
/// <c>twoCell</c> keeps no anchor and is the baseline: what the load and save of the same sheet cost
/// without the comparison. Doubling the rows doubles a cost that is linear in them; with the walk,
/// the gap between <c>twoCell</c> and the other two doubles too, and is far larger.
/// </para>
/// </remarks>
public static class KeptAnchorScalingProfile
{
    private const int Pictures = 1_000;
    private static readonly int[] RowCounts = [12_500, 25_000, 50_000];
    private static readonly string[] EditAs = ["twoCell", "oneCell", "absolute"];

    // The drawing of 1,000 anchors is counted in well under a second.
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(10);

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    public static void Run(string[] args)
    {
        PinataLabsFontBootstrap.Register();

        // args[0] is "profile" and args[1] the mode.
        var pictures = args.Length > 2 ? int.Parse(args[2], CultureInfo.InvariantCulture) : Pictures;
        int[] rowCounts = args.Length > 3 ? [int.Parse(args[3], CultureInfo.InvariantCulture)] : RowCounts;

        // Warm the load and save paths so the first measured case does not absorb JIT.
        var warmup = Fixture(500, 50);
        foreach (var editAs in EditAs)
            LoadAndSave(WithEditAs(warmup, editAs), 50, editAs);

        Console.WriteLine();
        Console.WriteLine($"Load and save of {pictures:N0} pictures on a sheet whose every row has a set height.");
        Console.WriteLine("Best of three single passes. twoCell keeps no anchor and is the baseline.");
        Console.WriteLine();
        Console.WriteLine("  |   rows | editAs   | load ms | save ms |");
        Console.WriteLine("  |--------|----------|---------|---------|");

        foreach (var rows in rowCounts)
        {
            var fixture = Fixture(rows, pictures);
            foreach (var editAs in EditAs)
            {
                var source = WithEditAs(fixture, editAs);
                double bestLoad = double.MaxValue, bestSave = double.MaxValue;
                for (var run = 0; run < 3; run++)
                {
                    var (load, save) = LoadAndSave(source, pictures, editAs);
                    bestLoad = Math.Min(bestLoad, load);
                    bestSave = Math.Min(bestSave, save);
                }

                Console.WriteLine($"  | {rows,6:N0} | {editAs,-8} | {bestLoad,7:F0} | {bestSave,7:F0} |");
            }
        }
    }

    /// <summary>
    /// A sheet of <paramref name="rows"/> rows, each 20 pt tall, with <paramref name="pictures"/>
    /// pictures spread evenly down it, each anchored over three columns and four rows.
    /// </summary>
    internal static byte[] Fixture(int rows, int pictures)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= rows; row++)
            ws.Row(row).Height = 20;

        var step = rows / pictures;
        for (var i = 0; i < pictures; i++)
        {
            var top = 1 + i * step;
            ws.AddPicture(new MemoryStream(Png), $"Picture{i}")
                .MoveTo(ws.Cell(top, 2), ws.Cell(Math.Min(top + 3, rows), 4));
        }

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>The fixture with every anchor's <c>editAs</c> set, as Excel writes an inserted picture.</summary>
    internal static byte[] WithEditAs(byte[] fixture, string editAs)
    {
        using var stream = new MemoryStream();
        stream.Write(fixture);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("xl/drawings/drawing1.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();

            entry.Delete();
            var rewritten = zip.CreateEntry("xl/drawings/drawing1.xml");
            using var writer = new StreamWriter(rewritten.Open());
            writer.Write(xml.Replace("<xdr:twoCellAnchor>", $"<xdr:twoCellAnchor editAs=\"{editAs}\">", StringComparison.Ordinal));
        }

        return stream.ToArray();
    }

    private static (double LoadMs, double SaveMs) LoadAndSave(byte[] source, int pictures, string editAs)
    {
        var sw = Stopwatch.StartNew();
        using var wb = new XLWorkbook(new MemoryStream(source));
        var loadMs = sw.Elapsed.TotalMilliseconds;

        using var saved = new MemoryStream();
        sw.Restart();
        wb.SaveAs(saved);
        var saveMs = sw.Elapsed.TotalMilliseconds;

        // The comparison must still find every anchor unchanged, or the profile measures the
        // regenerating writer instead.
        if (editAs != "twoCell" && CountKept(saved, editAs) != pictures)
            throw new InvalidOperationException($"Expected {pictures} kept {editAs} anchors after the save.");

        return (loadMs, saveMs);
    }

    private static int CountKept(MemoryStream saved, string editAs)
    {
        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(zip.GetEntry("xl/drawings/drawing1.xml")!.Open());
        return Regex.Count(reader.ReadToEnd(), $"<xdr:twoCellAnchor editAs=\"{editAs}\">", RegexOptions.None, RegexTimeout);
    }
}
