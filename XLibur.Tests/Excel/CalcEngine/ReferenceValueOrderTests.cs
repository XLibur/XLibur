using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Functions that read the values of a reference all go through the one sparse iterator,
/// <c>CalcContext.GetNonBlankValues</c>. It visits only the cells that hold something, yet it must
/// still hand them over in row-major order — left to right, then top to bottom — because NPV, IRR
/// and MIRR depend on the order.
/// </summary>
public class ReferenceValueOrderTests
{
    /// <summary>
    /// A 2-column block whose row-major and column-major orders differ, with a blank cell to skip
    /// and one value at the very bottom of the sheet, reached only by a whole-column reference.
    /// </summary>
    private static IXLWorksheet NewSheet(out XLWorkbook wb, string left, string right, params XLCellValue[] values)
    {
        wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(left + "1").Value = values[0];
        ws.Cell(right + "1").Value = values[1];
        ws.Cell(left + "2").Value = values[2];
        ws.Cell(right + "2").Value = values[3];
        // Row 3 of the left column stays blank.
        ws.Cell(right + "3").Value = values[4];
        ws.Cell(left + "1048576").Value = values[5];
        return ws;
    }

    [Test]
    public async Task Npv_ReadsAReferenceInRowMajorOrder()
    {
        var ws = NewSheet(out var wb, "A", "B", 1, 2, 3, 4, 5, 6);
        using (wb)
        {
            // 1/2 + 2/4 + 3/8 + 4/16 + 5/32. Column-major order (1, 3, 2, 4, 5) would give 1.90625.
            await Assert.That((double)ws.Evaluate("NPV(1, A1:B3)")).IsEqualTo(1.78125);
            await Assert.That((double)ws.Evaluate("NPV(1, A1:B3)")).IsEqualTo((double)ws.Evaluate("NPV(1, 1, 2, 3, 4, 5)"));

            // The whole-column reference adds the bottom cell last.
            await Assert.That((double)ws.Evaluate("NPV(1, A:B)")).IsEqualTo((double)ws.Evaluate("NPV(1, 1, 2, 3, 4, 5, 6)"));
        }
    }

    [Test]
    public async Task IrrAndMirr_ReadAReferenceInRowMajorOrder()
    {
        var ws = NewSheet(out var wb, "D", "E", -100, 10, 20, 110, 30, 5);
        using (wb)
        {
            await Assert.That((double)ws.Evaluate("IRR(D1:E3)")).IsEqualTo((double)ws.Evaluate("IRR({-100,10,20,110,30})")).Within(1e-12);
            await Assert.That((double)ws.Evaluate("IRR(D:E)")).IsEqualTo((double)ws.Evaluate("IRR({-100,10,20,110,30,5})")).Within(1e-12);
            await Assert.That((double)ws.Evaluate("IRR(D:E)")).IsNotEqualTo((double)ws.Evaluate("IRR({-100,20,5,10,110,30})"));

            await Assert.That((double)ws.Evaluate("MIRR(D:E, 0.1, 0.12)")).IsEqualTo((double)ws.Evaluate("MIRR({-100,10,20,110,30,5}, 0.1, 0.12)")).Within(1e-12);
        }
    }

    [Test]
    public async Task WholeColumnReferenceToASparseSheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("J1").Value = true;
        ws.Cell("J1048576").Value = false;
        ws.Cell("K1").Value = 2;
        ws.Cell("K1048576").Value = 3;

        // Only the two used cells of each column are read; the million blanks in between are not.
        await Assert.That((bool)ws.Evaluate("AND(J:J)")).IsFalse();
        await Assert.That((bool)ws.Evaluate("OR(J:J)")).IsTrue();
        await Assert.That(ws.Evaluate("MULTINOMIAL(K:K)")).IsEqualTo(10); // 5! / (2! 3!)
        await Assert.That(ws.Evaluate("AND(L:L)")).IsEqualTo(XLError.IncompatibleValue); // No values at all.
    }
}
