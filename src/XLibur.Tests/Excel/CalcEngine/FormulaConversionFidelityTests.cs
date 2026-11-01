using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Formula text that has to survive a conversion through R1C1 — the form a shared formula is
/// stored in — and a sheet prefix that has to be quoted the way the file format requires.
/// </summary>
/// <remarks>
/// Both are parser-side defects fixed by moving to XLibur.ClosedXML.Parser; see
/// https://github.com/XLibur/XLibur/issues/313. They are pinned here because nothing in the suite
/// exercised them from XLibur's own API, so the swap would otherwise have had no guard.
/// </remarks>
public class FormulaConversionFidelityTests
{
    /// <summary>
    /// A shared formula is written as R1C1 and read back as A1. The parser used to write the
    /// offsets with the current culture's negative sign, so under a culture whose sign is U+2212
    /// it emitted <c>R[−1]C[−1]</c>, which its own reader could not read back.
    /// </summary>
    [Test]
    [SetCulture("sv-SE")]
    public async Task AFormulaConvertsThroughR1C1UnderACultureWithAMinusSign()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("B2").FormulaA1 = "A1+B1";

        await Assert.That(ws.Cell("B2").FormulaR1C1).IsEqualTo("R[-1]C[-1]+R[-1]C");

        ws.Cell("C3").FormulaR1C1 = ws.Cell("B2").FormulaR1C1;

        await Assert.That(ws.Cell("C3").FormulaA1).IsEqualTo("B2+C2");
    }

    /// <summary>
    /// A sheet whose name Excel's formula bar shows unquoted but whose file format requires
    /// quotes. The quotation tables were collected from the formula bar, so the prefix came back
    /// unquoted and Excel refused the workbook.
    /// </summary>
    /// <remarks>
    /// <c>TRUE</c> is the readable case — unquoted, <c>TRUE!A1</c> reads as a logical literal.
    /// U+FF5E is one of the 41 codepoints the tables got wrong in the first position.
    /// </remarks>
    [Test]
    [Arguments("TRUE")]
    [Arguments("ABC～")]
    public async Task ASheetPrefixKeepsItsQuotesThroughAnR1C1RoundTrip(string sheetName)
    {
        using var wb = new XLWorkbook();
        var other = wb.AddWorksheet(sheetName);
        other.Cell("A1").Value = 7;
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("B2").FormulaA1 = $"'{sheetName}'!A1";

        // The prefix survives the trip out to R1C1 and back, quotes intact.
        await Assert.That(ws.Cell("B2").FormulaR1C1).IsEqualTo($"'{sheetName}'!R[-1]C[-1]");
        await Assert.That(ws.Cell("B2").FormulaA1).IsEqualTo($"'{sheetName}'!A1");
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(7);
    }

    /// <summary>
    /// The same prefix has to reach the saved file quoted, since that is where Excel reads it.
    /// </summary>
    /// <remarks>
    /// Unlike the two above, this passed on ClosedXML.Parser 2.0.0 as well: the save path writes
    /// the A1 text the caller authored and never asks the parser to quote a prefix. It is here to
    /// say so — the quoting defect only surfaces once a formula is rewritten, which is what a
    /// shared formula, a shift and a sheet rename all do.
    /// </remarks>
    [Test]
    [Arguments("TRUE")]
    [Arguments("ABC～")]
    public async Task ASheetPrefixIsQuotedInTheSavedFile(string sheetName)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var other = wb.AddWorksheet(sheetName);
            other.Cell("A1").Value = 7;
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("B2").FormulaA1 = $"'{sheetName}'!A1";
            wb.SaveAs(ms, validate: false);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var cell = reloaded.Worksheet("Sheet1").Cell("B2");

        await Assert.That(cell.FormulaA1).IsEqualTo($"'{sheetName}'!A1");
        await Assert.That(cell.Value).IsEqualTo(7);
    }
}
