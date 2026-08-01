using System;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Coordinates;

public class PointTests
{
    [Test]
    [Arguments("A1", 1, 1)]
    [Arguments("AA1", 27, 1)]
    [Arguments("AAA1", 703, 1)]
    [Arguments("Z1", 26, 1)]
    [Arguments("ZZ1", 702, 1)]
    [Arguments("XFD1", 16384, 1)]
    [Arguments("A1", 1, 1)]
    [Arguments("A999", 1, 999)]
    [Arguments("XFD1048576", 16384, 1048576)]
    public async Task ParseCellRefsAccordingToGrammar(string cellRef, int columnNumber, int rowNumber)
    {
        var sheetPoint = Point.Parse(cellRef.AsSpan());
        await Assert.That(sheetPoint.Column).IsEqualTo(columnNumber);
        await Assert.That(sheetPoint.Row).IsEqualTo(rowNumber);
    }

    [Test]
    [Arguments("")]
    [Arguments(" ")]
    [Arguments("A")]
    [Arguments("AA")]
    [Arguments("1")]
    [Arguments("11")]
    [Arguments(" A1")]
    [Arguments("A1 ")]
    [Arguments("A 1")]
    [Arguments("@1")] // @ is a char 'A' - 1
    [Arguments("[1")] // [ is a char 'Z' + 1
    [Arguments("A:")] // : is a char '9' + 1
    [Arguments("A/")] // / is a char '0' - 1
    [Arguments("A1:")]
    [Arguments("A1/")]
    [Arguments("A@1")]
    [Arguments("A[1")]
    [Arguments("XFE1")]
    [Arguments("AAAA1")]
    [Arguments("A1048577")]
    [Arguments("A01")]
    [Arguments("A0")]
    [Arguments("A-1")]
    public async Task InvalidInputsAreNotParsed(string cellRef)
    {
        await Assert.That(() => Point.Parse(cellRef.AsSpan())).Throws<FormatException>();
    }

    [Test]
    [Arguments("A1")]
    [Arguments("DE1")]
    [Arguments("D174")]
    [Arguments("XFD1048576")]
    public async Task CanFormatToString(string cellRef)
    {
        var r = Point.Parse(cellRef);
        await Assert.That(r.ToString()).IsEqualTo(cellRef);
    }
}
