using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Parser;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Functions;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The regression gate for formula text: each syntax form in <c>FormulaTextCorpus.tsv</c>, sent down
/// each path that hands formula text to the parser, with the answer that path gives.
/// <para>
/// A row is a syntax form. A column is a path: evaluation, the references a defined name collects,
/// the extent a shift filters on, a row shift, a sheet rename, conversion to R1C1, future-function
/// prefixing, and the check SUBTOTAL and AGGREGATE use to skip a nested call. Each cell is the text
/// or value the path produced, <c>REFUSED</c> where the path reported that the parser refused the
/// text, or <c>THROWS</c> and the exception type.
/// </para>
/// <para>
/// Encoding: <c>&lt;empty&gt;</c> is the empty string, and <c>␠</c> is a space at the start or end of
/// a text, where a TSV would lose it. A text in the evaluate column is quoted.
/// </para>
/// </summary>
public class FormulaTextCorpusTests
{
    private const string EmptyToken = "<empty>";
    private const char VisibleSpace = '␠';

    [Test]
    [MethodDataSource(nameof(Corpus))]
    public async Task Path_matches_the_corpus(CorpusCell cell)
    {
        var actual = Run(cell.Path, cell.Text);

        await Assert.That(actual).IsEqualTo(cell.Expected);
    }

    /// <summary>
    /// The corpus must not quietly lose a path or a row: every row has a cell for every path.
    /// </summary>
    [Test]
    public async Task Every_row_has_a_cell_for_every_path()
    {
        var (paths, rows) = ReadCorpus();

        await Assert.That(paths).IsEquivalentTo(Paths);
        foreach (var row in rows)
            await Assert.That(row.Length).IsEqualTo(Paths.Length + 2);
    }

    internal static readonly string[] Paths =
        ["evaluate", "references", "extent", "shift", "rename", "to_r1c1", "add_prefix", "calls_subtotal"];

    internal static string Run(string path, string text)
    {
        try
        {
            return Encode(RunUnguarded(path, text));
        }
        catch (Exception ex)
        {
            return "THROWS " + ex.GetType().Name;
        }
    }

    private static string RunUnguarded(string path, string text)
    {
        switch (path)
        {
            case "evaluate":
            {
                using var wb = Fixture();
                return Render(wb.Worksheet("Sheet1").Evaluate(text, "H2"));
            }
            case "references":
            {
                using var wb = Fixture();
                return References(wb, text);
            }
            case "extent":
            {
                var extent = FormulaExtent.Of(text);
                return string.Create(CultureInfo.InvariantCulture, $"{extent.MaxRow},{extent.MaxColumn}");
            }
            case "shift":
            {
                // Two rows inserted above row 2 of Sheet1, by a formula on Sheet1.
                using var wb = Fixture();
                var sheet = (XLWorksheet)wb.Worksheet("Sheet1");
                var inserted = (XLRange)sheet.Range(2, 1, 3, XLHelper.MaxColumnNumber);
                return XLCellFormulaShifter.ShiftFormulaRows(text, sheet, inserted, 2);
            }
            case "rename":
            {
                var formula = XLCellFormula.NormalA1(text);
                formula.RenameSheet(new Point(2, 8), "Sheet1", "Data");
                return formula.A1;
            }
            case "to_r1c1":
                return XLCellFormula.GetFormula(text, FormulaConversionType.A1ToR1C1, new Point(3, 3));
            case "add_prefix":
                return FormulaText.AddFuturePrefixes(text, "Sheet1", new Point(3, 3));
            case "calls_subtotal":
                return CalcContext.IsSkippedByNestingCheck(text, TallyNumbers.SubtotalAndAggregate) ? "true" : "false";
            default:
                throw new ArgumentOutOfRangeException(nameof(path), path, "Not a corpus path.");
        }
    }

    /// <summary>
    /// <c>Sheet1</c> with 1 to 9 in <c>A1:C3</c>, a table <c>Table1</c> in <c>E1:F3</c> whose second
    /// column is named <c>Start: Date</c> and holds 3 and 4, a workbook-scoped name <c>WbName</c>
    /// (<c>A1:C1</c>, which sums to 6) and a sheet-scoped name <c>Local</c> (<c>B2</c>, which is 5).
    /// </summary>
    private static XLWorkbook Fixture()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= 3; row++)
        {
            for (var column = 1; column <= 3; column++)
                ws.Cell(row, column).Value = (row - 1) * 3 + column;
        }

        ws.Cell("E1").Value = "Name";
        ws.Cell("F1").Value = "Start: Date";
        ws.Cell("E2").Value = "a";
        ws.Cell("F2").Value = 3;
        ws.Cell("E3").Value = "b";
        ws.Cell("F3").Value = 4;
        ws.Range("E1:F3").CreateTable("Table1");

        wb.DefinedNames.Add("WbName", "Sheet1!$A$1:$C$1");
        ws.DefinedNames.Add("Local", "Sheet1!$B$2");
        return wb;
    }

    private static string Render(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => "(blank)",
        XLDataType.Text => "\"" + value.GetText() + "\"",
        XLDataType.Number => value.GetNumber().ToString("R", CultureInfo.InvariantCulture),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    private static string References(XLWorkbook wb, string text)
    {
        if (!FormulaReferences.TryForFormula(text, out var references, out var failure))
        {
            return failure is ParsingException || failure.InnerException is ParsingException
                ? "REFUSED"
                : "FAILS " + failure.GetType().Name;
        }

        var parts = new List<string>();
        if (references.References.Count > 0)
            parts.Add("local " + Join(references.References.Select(x => x.GetA1())));

        if (references.SheetReferences.Count > 0)
            parts.Add("sheet " + Join(references.SheetReferences.Select(x => x.GetA1())));

        var ranges = references.GetExternalRanges(wb, new Point(1, 1));
        if (ranges.Count > 0)
            parts.Add("ranges " + Join(ranges.Select(x => x.RangeAddress.ToString(XLReferenceStyle.A1, true))));

        if (references.ContainsRefError)
            parts.Add("#REF!");

        return parts.Count == 0 ? "none" : string.Join("; ", parts);

        static string Join(IEnumerable<string> items) => string.Join(",", items.Order(StringComparer.Ordinal));
    }

    private static string Encode(string text)
    {
        if (text.Length == 0)
            return EmptyToken;

        var start = text.Length - text.AsSpan().TrimStart(' ').Length;
        var end = text.AsSpan().TrimEnd(' ').Length;
        if (start == text.Length)
            return new string(VisibleSpace, text.Length);

        return new string(VisibleSpace, start) + text[start..end] + new string(VisibleSpace, text.Length - end);
    }

    private static string Decode(string cell)
        => cell == EmptyToken ? string.Empty : cell.Replace(VisibleSpace, ' ');

    internal static (string[] Paths, List<string[]> Rows) ReadCorpus()
    {
        // The extractor prefixes "XLibur.Tests.Resource." itself.
        using var stream = TestHelper.GetStreamFromResource("Other.FormulaTextCorpus.tsv");
        using var reader = new StreamReader(stream);

        string[]? paths = null;
        var rows = new List<string[]>();
        while (reader.ReadLine() is { } line)
        {
            if (line.Length == 0)
                continue;

            if (line[0] == '#')
            {
                // The first comment line is the header: form, text, then one column per path.
                paths ??= line.TrimStart('#', ' ').Split('\t').Skip(2).ToArray();
                continue;
            }

            rows.Add(line.Split('\t'));
        }

        return (paths ?? [], rows);
    }

    public static IEnumerable<Func<CorpusCell>> Corpus()
    {
        var (paths, rows) = ReadCorpus();
        foreach (var row in rows)
        {
            for (var i = 0; i < paths.Length; i++)
            {
                var cell = new CorpusCell(row[0], Decode(row[1]), paths[i], row[i + 2]);
                yield return () => cell;
            }
        }
    }

    public sealed record CorpusCell(string Form, string Text, string Path, string Expected)
    {
        // Keeps the test-name column readable instead of showing the record's full property dump.
        public override string ToString() => $"{Form} / {Path}";
    }

    /// <summary>
    /// Each corpus row whole: its form, its text, and its cell for every path, decoded.
    /// </summary>
    public static IEnumerable<Func<CorpusRow>> Rows()
    {
        var (paths, rows) = ReadCorpus();
        foreach (var row in rows)
        {
            var cells = new Dictionary<string, string>();
            for (var i = 0; i < paths.Length; i++)
                cells[paths[i]] = Decode(row[i + 2]);

            var corpusRow = new CorpusRow(row[0], Decode(row[1]), cells);
            yield return () => corpusRow;
        }
    }

    public sealed record CorpusRow(string Form, string Text, IReadOnlyDictionary<string, string> Cells)
    {
        public override string ToString() => Form;
    }

    /// <summary>
    /// Today's answer for every cell of a row, in corpus order. Used to write the corpus.
    /// </summary>
    internal static IEnumerable<string> RunRow(string text) => Paths.Select(path => Run(path, text));

    internal static string DecodeText(string cell) => Decode(cell);
}
