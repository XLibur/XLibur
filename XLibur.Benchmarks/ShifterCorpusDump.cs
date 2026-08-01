using System;
using System.Globalization;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Generates the <c>XLCellFormulaShifter</c> equivalence corpus asserted by
/// <c>FormulaShifterCorpusTests</c>, as tab-separated rows on stdout.
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile shiftercorpus
///
/// This is the supported way to regenerate <c>XLibur.Tests/Resource/Other/FormulaShifterCorpus.tsv</c>
/// after changing the shifter or widening the case matrix — redirect stdout over that file. Each row
/// carries both the parser and the legacy implementation's output, and any disagreement between them is
/// reported on stderr so a change in behaviour has to be looked at rather than silently re-baselined.
/// </summary>
public static class ShifterCorpusDump
{
    private static readonly string[] Formulas =
    [
        "A1", "A3", "A5", "A10", "B3:B7", "A2:A8", "A1:C1", "$A$5", "$A5", "A$5", "A5:A10", "$A$5:$B$10",
        "3:5", "$3:$5", "1:1", "5:5", "4:8", "B:D", "$B:$D", "C:F",
        "Sheet1!A5", "Sheet1!A5:B10", "'My Sheet'!A5", "Sheet2!A5", "Sheet2!A5:B10",
        "SUM(A1:A10)", "SUM(A5:A10)+B3", "IF(A5>0,B6,C7)", "\"A5\"", "\"x\"&A5", "\"A5\"&A5",
        "A5+A5", "SUM(Sheet1!A5:A10,Sheet2!A5:A10)", "A1048576", "A1:A1048576", "XFD5", "A5:XFD5",
    ];

    /// <summary>
    /// Insertions pass the range the caller asked to insert at, and a shift count that is independent
    /// of that range's height (<c>Row(3).InsertRowsAbove(2)</c> passes row 3 and +2). Deletions always
    /// pass <c>|shift| == range height</c>, because XLRangeBase.Delete derives the shift from
    /// <c>RowCount()</c>. Generating the incoherent combinations instead — a 3-row range with a -5
    /// shift — produces inverted output ranges like <c>B3:B2</c> that no caller can ever trigger, and
    /// baking those into an equivalence corpus would constrain the rewrite on inputs that do not exist.
    /// </summary>
    private static readonly (int First, int Last, int Shift)[] RowShifts =
    [
        (3, 3, 1), (3, 3, 2), (3, 3, 5), (1, 1, 1), (1, 1, 3), (5, 5, 1), (5, 5, 2), (6, 6, 1), (5, 7, 2),
        (3, 3, -1), (1, 1, -1), (5, 5, -1), (5, 7, -3), (1, 3, -3), (5, 9, -5), (2, 8, -7),
    ];

    private static readonly (int First, int Last, int Shift)[] ColumnShifts =
    [
        (3, 3, 1), (3, 3, 2), (1, 1, 1), (1, 1, 3), (2, 2, 1), (2, 4, 2),
        (3, 3, -1), (1, 1, -1), (2, 2, -1), (2, 4, -3), (1, 3, -3), (4, 6, -3),
    ];

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        Console.WriteLine("# axis\tformula\tfirst\tlast\tshift\tforeignFormula\texpected\tlegacyExpected");

        foreach (var foreign in new[] { false, true })
        {
            foreach (var (first, last, shift) in RowShifts)
            {
                foreach (var formula in Formulas)
                    Console.WriteLine("row\t" + Row(formula, first, last, shift, foreign));
            }

            foreach (var (first, last, shift) in ColumnShifts)
            {
                foreach (var formula in Formulas)
                    Console.WriteLine("col\t" + Column(formula, first, last, shift, foreign));
            }
        }
    }

    private static string Row(string formula, int first, int last, int shift, bool foreignFormula)
    {
        using var wb = NewWorkbook(out var shiftedSheet, out var otherSheet);
        var shifted = shiftedSheet.Range(first, 1, last, XLHelper.MaxColumnNumber);
        var host = foreignFormula ? otherSheet : shiftedSheet;
        return Format(formula, first, last, shift, foreignFormula,
            Invoke(() => XLCellFormulaShifter.ShiftFormulaRows(formula, host, shifted, shift)),
            Invoke(() => XLCellFormulaShifter.ShiftFormulaRowsLegacy(formula, host, shifted, shift)));
    }

    private static string Column(string formula, int first, int last, int shift, bool foreignFormula)
    {
        using var wb = NewWorkbook(out var shiftedSheet, out var otherSheet);
        var shifted = shiftedSheet.Range(1, first, XLHelper.MaxRowNumber, last);
        var host = foreignFormula ? otherSheet : shiftedSheet;
        return Format(formula, first, last, shift, foreignFormula,
            Invoke(() => XLCellFormulaShifter.ShiftFormulaColumns(formula, host, shifted, shift)),
            Invoke(() => XLCellFormulaShifter.ShiftFormulaColumnsLegacy(formula, host, shifted, shift)));
    }

    /// <summary>
    /// <paramref name="shiftedSheet"/> is the sheet being structurally edited; <paramref name="otherSheet"/>
    /// hosts formulas that merely refer to it. Every worksheet's formulas are visited on a shift, so the
    /// distinction decides what an *unqualified* reference means — and getting it wrong is invisible
    /// unless the corpus runs formulas from a sheet other than the one being shifted.
    /// </summary>
    private static XLWorkbook NewWorkbook(out XLWorksheet shiftedSheet, out XLWorksheet otherSheet)
    {
        var wb = new XLWorkbook();
        shiftedSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("My Sheet");
        otherSheet = (XLWorksheet)wb.AddWorksheet("Sheet2");
        return wb;
    }

    private static string Invoke(Func<string> shift)
    {
        try
        {
            return shift();
        }
        catch (Exception e)
        {
            return "THROWS:" + e.GetType().Name;
        }
    }

    /// <summary>
    /// Emits a tab-separated corpus row carrying <em>both</em> implementations' output, and reports any
    /// disagreement on stderr so every divergence can be reviewed in one pass rather than by diffing two
    /// whole runs.
    /// </summary>
    /// <remarks>
    /// The legacy column is not redundant. The regex implementation is still live — it is the fallback
    /// for formulas the parser rejects — so it needs pinning too, and recording its output beside the
    /// parser's puts the nine known divergences in the data rather than in prose.
    /// </remarks>
    private static string Format(string formula, int first, int last, int shift, bool foreignFormula,
        string result, string legacyResult)
    {
        // Invariant throughout: the corpus is generated on whatever machine runs this and parsed back by
        // int.Parse under the test suite's en-US default, so the shift column — the only one that can be
        // negative — must not pick up a culture-specific minus sign on the way out.
        var firstText = first.ToString(CultureInfo.InvariantCulture);
        var lastText = last.ToString(CultureInfo.InvariantCulture);
        var shiftText = shift.ToString(CultureInfo.InvariantCulture);

        if (result != legacyResult)
        {
            Console.Error.WriteLine(
                $"DIVERGES\t{formula}\t{firstText}\t{lastText}\t{shiftText}\t{foreignFormula}\t{result}\tlegacy={legacyResult}");
        }

        return string.Join('\t', formula, firstText, lastText, shiftText, foreignFormula ? "1" : "0", result,
            legacyResult);
    }
}
