using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// Cases where <c>GetFormattedString</c> disagreed with Excel while the app used ExcelNumberFormat 1.1.0
/// (ClosedXML/ClosedXML#2898): exact halves rounded to even, and literals in the text section were dropped.
/// </summary>
public class NumberFormatExcelParityTests
{
    [Test]
    [Arguments(2.5, "0", "3")]
    [Arguments(0.5, "0", "1")]
    [Arguments(-2.5, "0", "-3")]
    [Arguments(1234.5, "#,##0", "1,235")]
    [Arguments(0.125, "0.00", "0.13")]
    [Arguments(-0.125, "0.00", "-0.13")]
    [Arguments(0.625, "0.00", "0.63")]
    public async Task Exact_halves_round_away_from_zero(double value, string format, string expected)
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("S").Cell("A1");
        cell.Value = value;
        cell.Style.NumberFormat.Format = format;

        await Assert.That(cell.GetFormattedString()).IsEqualTo(expected);
    }

    [Test]
    [Skip("Still fails on XLibur.ExcelNumberFormat 1.2.1: literals in the text section are dropped (ClosedXML/ClosedXML#2898). Remove when fixed.")]
    [Arguments("\"Prefix: \"@", "abc", "Prefix: abc")]
    [Arguments("@\" suffix\"", "abc", "abc suffix")]
    [Arguments("0;-0;0;\"Note: \"@", "abc", "Note: abc")]
    public async Task Text_section_keeps_literals(string format, string text, string expected)
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("S").Cell("A1");
        cell.Value = text;
        cell.Style.NumberFormat.Format = format;

        await Assert.That(cell.GetFormattedString()).IsEqualTo(expected);
    }
}
