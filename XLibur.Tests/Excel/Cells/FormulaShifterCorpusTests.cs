using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Parser;
using XLibur.Excel;
using XLibur.Excel.CalcEngine.Visitors;

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

    /// <summary>
    /// The fallback exists for formulas <c>ClosedXML.Parser</c> rejects. This pins both halves of that
    /// claim: an external workbook reference is rejected, and an ordinary formula is accepted — so
    /// narrowing the shifter's catch to <see cref="ParsingException"/> cannot silently reroute anything
    /// that reaches the parser path today.
    /// </summary>
    [Test]
    [Arguments("='[file.xlsx]Sheet'!A1", false)]
    [Arguments("=SUM('[book.xlsx]Data'!A1:A5)", false)]
    [Arguments("=A1+B2", true)]
    [Arguments("=SUM(A1:A5)", true)]
    [Arguments("=Sheet2!A1", true)]
    public async Task The_parser_accepts_only_what_the_fallback_is_not_for(string formula, bool parseable)
    {
        await Assert.That(TryParse(formula)).IsEqualTo(parseable);
    }

    /// <summary>
    /// Every corpus formula must parse, or the corpus is silently testing the regex path through the
    /// shifter's fallback while claiming to test the parser path.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(Corpus))]
    public async Task Every_corpus_formula_is_accepted_by_the_parser(CorpusCase test)
    {
        await Assert.That(TryParse(test.Formula)).IsTrue();
    }

    /// <summary>
    /// External workbook references are what the fallback exists for, and nothing exercised them
    /// through <c>Shift</c> itself — the corpus calls the two implementations directly, and every one of
    /// its formulas parses. These go in the front door and come out the regex path, on both axes and
    /// including the <c>#REF!</c> collapse.
    /// </summary>
    /// <remarks>
    /// The external reference itself never moves — neither implementation shifts a reference whose sheet
    /// is not the shifted one — so every case pairs it with a local reference that does. Without that,
    /// "both paths agree" would be satisfied by both returning the formula untouched.
    /// <para>
    /// These deliberately stay out of the corpus. The extractor records both columns by calling each
    /// implementation directly, so an external-reference row would hold the regex answer twice and say
    /// nothing about routing — and it would break
    /// <see cref="Every_corpus_formula_is_accepted_by_the_parser"/>, which is the guard that keeps the
    /// corpus testing the parser path. This explicit test is the coverage instead.
    /// </para>
    /// </remarks>
    [Test]
    [Arguments(true, "='[file.xlsx]Sheet'!A1+B2", 1, 2, 3, "='[file.xlsx]Sheet'!A1+B5")]
    [Arguments(true, "=SUM('[book.xlsx]Data'!A1:A20)+SUM(C10:C20)", 5, 9, -5,
        "=SUM('[book.xlsx]Data'!A1:A20)+SUM(C5:C15)")]
    [Arguments(true, "='[file.xlsx]Sheet'!A1+D7", 5, 9, -5, "='[file.xlsx]Sheet'!A1+#REF!")]
    [Arguments(false, "='[file.xlsx]Sheet'!A1+B2", 1, 2, 3, "='[file.xlsx]Sheet'!A1+E2")]
    [Arguments(false, "=SUM('[book.xlsx]Data'!A1:A20)+SUM(J10:L10)", 5, 9, -5,
        "=SUM('[book.xlsx]Data'!A1:A20)+SUM(E10:G10)")]
    public async Task An_external_reference_shifts_through_the_fallback(
        bool rowShift, string formula, int first, int last, int shift, string expected)
    {
        using var wb = new XLWorkbook();
        var shiftedSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");

        var range = rowShift
            ? (XLRange)shiftedSheet.Range(first, 1, last, XLHelper.MaxColumnNumber)
            : (XLRange)shiftedSheet.Range(1, first, XLHelper.MaxRowNumber, last);
        var axis = rowShift ? XLCellFormulaShifter.ShiftAxis.Row : XLCellFormulaShifter.ShiftAxis.Column;

        var throughShift = rowShift
            ? XLCellFormulaShifter.ShiftFormulaRows(formula, shiftedSheet, range, shift)
            : XLCellFormulaShifter.ShiftFormulaColumns(formula, shiftedSheet, range, shift);

        var throughFallback = XLCellFormulaShifter.ShiftUnparseable(
            formula, shiftedSheet, range, shift, axis);

        // The pinned value is what makes this more than a tautology: every case moves a reference, so
        // "the two agree" cannot be satisfied by both paths returning the formula untouched.
        await Assert.That(throughShift).IsEqualTo(expected);

        // Reaching the same answer both ways is what proves Shift routed here rather than succeeding on
        // the parser path with a different result.
        await Assert.That(throughShift).IsEqualTo(throughFallback);
    }

    private static bool TryParse(string formula)
    {
        var text = formula.Length > 0 && formula[0] == '=' ? formula[1..] : formula;
        text = FormulaTransformation.ProtectStructuredRefColons(text, out _);
        try
        {
            FormulaParser<object?, object?, object?>.CellFormulaA1(text, null, ProbeFactory.Instance);
            return true;
        }
        catch (ParsingException)
        {
            return false;
        }
    }

    /// <summary>A do-nothing factory: the only thing asked of the parse is whether it throws.</summary>
    private sealed class ProbeFactory : CollectVisitor<object?>
    {
        internal static readonly ProbeFactory Instance = new();
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
