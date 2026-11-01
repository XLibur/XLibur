using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The tally functions read a reference through <c>CalcContext.NonBlankValueEnumerator</c>, a
/// struct, instead of the <c>GetNonBlankValues</c> iterator (#686). It must read the areas in
/// their order, each in row-major order, and skip only blank cells.
/// </summary>
public class NonBlankValueEnumeratorTests
{
    [Test]
    public async Task ReadsEveryAreaInOrder()
    {
        using var wb = new XLWorkbook();
        var s1 = wb.AddWorksheet("S1");
        var s2 = wb.AddWorksheet("S2");
        s1.Cell("A1").Value = 1;
        s1.Cell("B1").Value = 2;
        // A2 stays blank.
        s1.Cell("B2").Value = "x";
        s1.Cell("C2").Value = XLError.DivisionByZero;
        s1.Cell("A1048576").Value = 5;
        s2.Cell("A1").Value = 10;
        s2.Cell("A2").Value = true;

        var sheet1 = (XLWorksheet)s1;
        var sheet2 = (XLWorksheet)s2;
        var reference = new Reference(new List<XLRangeAddress>
        {
            new(sheet1, "A1:B2"),
            new(sheet1, "B2:C2"), // Overlaps B2, which is read again.
            new(sheet2, "A1:A2"), // Another sheet.
            new(sheet1, "D5:E6"), // Nothing in it.
            new(sheet1, "A:A"), // Too large to walk cell by cell.
        });

        var ctx = new CalcContext(wb.CalcEngine, CultureInfo.InvariantCulture, wb, sheet1, formulaPoint: null);
        var values = ctx.EnumerateNonBlankValues(reference);
        var read = new List<string>();
        while (values.MoveNext())
            read.Add(values.Current.ToCellValue().ToString(CultureInfo.InvariantCulture));

        await Assert.That(string.Join("|", read)).IsEqualTo("1|2|x|x|#DIV/0!|10|TRUE|1|5");

        // Past the end it stays at the end.
        await Assert.That(values.MoveNext()).IsFalse();
    }
}
