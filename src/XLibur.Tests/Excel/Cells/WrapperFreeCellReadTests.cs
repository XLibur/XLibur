using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Cells;

/// <summary>
/// <c>XLWorksheet.GetCellValue</c> reads a cell without minting an <c>XLCell</c> wrapper (#842). It must
/// answer exactly what <c>Cell(row, column).Value</c> and <c>GetValue&lt;T&gt;()</c> answer.
/// </summary>
public class WrapperFreeCellReadTests
{
    private static XLWorksheet CreateSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = "text";
        ws.Cell(1, 2).Value = 42.5;
        ws.Cell(1, 3).Value = true;
        ws.Cell(1, 4).Value = new DateTime(2024, 3, 5);
        ws.Cell(1, 5).Value = XLError.DivisionByZero;
        ws.Cell(1, 6).Value = "_x0041_";
        return ws;
    }

    [Test]
    public async Task GetCellValue_MatchesCellValue_ForEveryType()
    {
        var ws = CreateSheet(out var wb);
        using (wb)
        {
            for (var column = 1; column <= 7; column++)
            {
                var expected = ws.Cell(1, column).Value;
                await Assert.That(ws.GetCellValue(1, column)).IsEqualTo(expected);
            }
        }
    }

    [Test]
    public async Task GetCellValueOfT_MatchesGetValueOfT()
    {
        var ws = CreateSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.GetCellValue<string>(1, 1)).IsEqualTo(ws.Cell(1, 1).GetValue<string>());
            await Assert.That(ws.GetCellValue<double>(1, 2)).IsEqualTo(42.5);
            await Assert.That(ws.GetCellValue<int?>(1, 7)).IsNull();
            await Assert.That(ws.GetCellValue<bool>(1, 3)).IsTrue();
            await Assert.That(ws.GetCellValue<DateTime>(1, 4)).IsEqualTo(new DateTime(2024, 3, 5));
            await Assert.That(ws.GetCellValue<XLError>(1, 5)).IsEqualTo(XLError.DivisionByZero);
            await Assert.That(ws.GetCellValue<string>(1, 6)).IsEqualTo("A");
            await Assert.That(ws.GetCellValue<string>(1, 2)).IsEqualTo(ws.Cell(1, 2).GetValue<string>());
        }
    }

    [Test]
    public async Task PublicGetValue_MatchesCellGetValue()
    {
        var ws = CreateSheet(out var wb);
        using (wb)
        {
            IXLWorksheet api = ws;
            for (var column = 1; column <= 7; column++)
            {
                await Assert.That(api.GetValue<string>(1, column)).IsEqualTo(api.Cell(1, column).GetValue<string>());
            }

            await Assert.That(api.GetValue<double>(1, 2)).IsEqualTo(42.5);
            await Assert.That(() => api.GetValue<double>(1, 1)).Throws<InvalidCastException>();
            await Assert.That(() => api.GetValue<double>(0, 1)).Throws<ArgumentOutOfRangeException>();
        }
    }

    [Test]
    public async Task GetCellValue_EvaluatesADirtyFormula()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = 2;
        ws.Cell(1, 2).FormulaA1 = "A1*10";
        await Assert.That(ws.GetCellValue<double>(1, 2)).IsEqualTo(20d);

        ws.Cell(1, 1).Value = 3;
        await Assert.That(ws.GetCellValue<double>(1, 2)).IsEqualTo(30d);
    }

    [Test]
    public async Task TryGetCellValue_ReturnsFalse_ForAnUnconvertibleValue()
    {
        var ws = CreateSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.TryGetCellValue<double>(1, 1, out _)).IsFalse();
            await Assert.That(ws.TryGetCellValue<double>(1, 2, out var number)).IsTrue();
            await Assert.That(number).IsEqualTo(42.5);
        }
    }

    [Test]
    public async Task GetCellValueOfT_ThrowsTheSameInvalidCast_AsGetValue()
    {
        var ws = CreateSheet(out var wb);
        using (wb)
        {
            await Assert.That(() => ws.GetCellValue<double>(1, 1)).Throws<InvalidCastException>();

            var expected = CaptureMessage(() => ws.Cell(1, 1).GetValue<double>());
            var actual = CaptureMessage(() => ws.GetCellValue<double>(1, 1));
            await Assert.That(actual).IsEqualTo(expected);
        }
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(1, 0)]
    [Arguments(XLHelper.MaxRowNumber + 1, 1)]
    [Arguments(1, XLHelper.MaxColumnNumber + 1)]
    public async Task GetCellValue_OutsideTheSheet_Throws(int row, int column)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        await Assert.That(() => ws.GetCellValue(row, column)).Throws<ArgumentOutOfRangeException>();
    }

    private static string? CaptureMessage(Action action)
    {
        try
        {
            action();
            return null;
        }
        catch (InvalidCastException ex)
        {
            return ex.Message;
        }
    }

    [Test]
    public async Task GetCellValue_AllocatesNoWrapper()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        const int rows = 1_000;
        for (var r = 1; r <= rows; r++)
            ws.Cell(r, 1).Value = r;

        // Warm the JIT and the generic instantiation outside the measured window.
        double sum = 0;
        for (var r = 1; r <= rows; r++)
            sum += ws.GetCellValue<double>(r, 1);

        var before = GC.GetAllocatedBytesForCurrentThread();
        for (var r = 1; r <= rows; r++)
            sum += ws.GetCellValue<double>(r, 1);
        var perCall = (GC.GetAllocatedBytesForCurrentThread() - before) / rows;

        await Assert.That(sum).IsGreaterThan(0d);
        // The read itself allocates nothing (measured 0 B; the benchmark shows 18 B over 100k calls).
        // The ceiling is half an XLCell wrapper (48 B): it leaves room for coverage instrumentation and
        // tiering noise on this thread, and still fails if a wrapper is minted per call.
        await Assert.That(perCall).IsLessThan(24L);
    }
}
