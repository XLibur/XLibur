using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Cells;

/// <summary>
/// Equivalence corpus for <see cref="XLCellFormulaShifter"/>: 2,072 (formula, shifted range, shift)
/// combinations with the reference text each must produce.
/// <para>
/// It exists because the shifter was rewritten from a regex onto <c>ClosedXML.Parser</c>, and the
/// behaviour it encodes — <c>#REF!</c> collapse, boundary clamping, absolute markers, axis-only
/// references, string literals that look like addresses, cross-sheet resolution — was only ever
/// pinned by a handful of scattered tests. The corpus was generated from the regex implementation and
/// every divergence reviewed individually; nine were kept as fixes and are listed below. Everything
/// else is byte-identical to the behaviour that shipped.
/// </para>
/// <para>
/// The nine intentional fixes are all one bug: a deletion that removes the *tail* of a reference
/// computed the new bottom edge as <c>last + shift</c> without clamping it to the row above the
/// deletion, so a reference could come back inverted (<c>3:5</c> with rows 5-7 deleted gave
/// <c>3:2</c>) or silently lose a surviving row (<c>A2:A8</c> with rows 5-9 deleted gave
/// <c>A2:A3</c>, dropping row 4, where Excel gives <c>A2:A4</c>).
/// </para>
/// <para>
/// Each row carries both implementations' output. The regex path is still live — it is the fallback
/// for formulas the parser rejects — so it is pinned by the same cases against its own column, which
/// also puts the nine divergences in the data rather than only in this comment.
/// </para>
/// <para>
/// Regenerate with:
/// <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile shiftercorpus</c>
/// — it writes the corpus to stdout and reports divergences between the two columns on stderr.
/// </para>
/// </summary>
public class FormulaShifterCorpusTests
{
    [Test]
    [MethodDataSource(nameof(Corpus))]
    public async Task ShiftMatchesTheCorpus(CorpusCase test)
    {
        var actual = Shift(test, legacy: false);

        await Assert.That(actual).IsEqualTo(test.Expected);
    }

    /// <summary>
    /// The regex implementation is still live — it is the fallback for formulas the parser cannot
    /// parse, such as external workbook references — so it is pinned by the same corpus against its own
    /// recorded output. Nine of the 2,072 rows differ from the parser column, all of them the
    /// tail-deletion clamp described on the class.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(Corpus))]
    public async Task LegacyShiftMatchesTheCorpus(CorpusCase test)
    {
        var actual = Shift(test, legacy: true);

        await Assert.That(actual).IsEqualTo(test.LegacyExpected);
    }

    private static string Shift(CorpusCase test, bool legacy)
    {
        using var wb = new XLWorkbook();
        var shiftedSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("My Sheet");
        var otherSheet = (XLWorksheet)wb.AddWorksheet("Sheet2");

        // A formula on a sheet other than the one being shifted is the case that decides what an
        // unqualified reference means, so the corpus runs every formula from both positions.
        var host = test.ForeignFormula ? otherSheet : shiftedSheet;

        if (test.IsRowShift)
        {
            var range = (XLRange)shiftedSheet.Range(test.First, 1, test.Last, XLHelper.MaxColumnNumber);
            return legacy
                ? XLCellFormulaShifter.ShiftFormulaRowsLegacy(test.Formula, host, range, test.Shift)
                : XLCellFormulaShifter.ShiftFormulaRows(test.Formula, host, range, test.Shift);
        }

        var columnRange = (XLRange)shiftedSheet.Range(1, test.First, XLHelper.MaxRowNumber, test.Last);
        return legacy
            ? XLCellFormulaShifter.ShiftFormulaColumnsLegacy(test.Formula, host, columnRange, test.Shift)
            : XLCellFormulaShifter.ShiftFormulaColumns(test.Formula, host, columnRange, test.Shift);
    }

    public static IEnumerable<Func<CorpusCase>> Corpus()
    {
        // The extractor prefixes "XLibur.Tests.Resource." itself.
        using var stream = TestHelper.GetStreamFromResource("Other.FormulaShifterCorpus.tsv");
        using var reader = new StreamReader(stream);

        while (reader.ReadLine() is { } line)
        {
            if (line.Length == 0 || line[0] == '#')
                continue;

            var parsed = CorpusCase.Parse(line);
            yield return () => parsed;
        }
    }

    public sealed record CorpusCase(
        bool IsRowShift,
        string Formula,
        int First,
        int Last,
        int Shift,
        bool ForeignFormula,
        string Expected,
        string LegacyExpected)
    {
        internal static CorpusCase Parse(string line)
        {
            var f = line.Split('\t');
            return new CorpusCase(
                IsRowShift: f[0] == "row",
                Formula: f[1],
                First: int.Parse(f[2]),
                Last: int.Parse(f[3]),
                Shift: int.Parse(f[4]),
                ForeignFormula: f[5] == "1",
                Expected: f[6],
                LegacyExpected: f[7]);
        }

        // Keeps the test-name column readable instead of showing the record's full property dump.
        public override string ToString() =>
            $"{(IsRowShift ? "row" : "col")} {Formula} @{First}:{Last} {Shift:+#;-#;0}{(ForeignFormula ? " foreign" : "")}";
    }
}
