using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine.Visitors;

namespace XLibur.Tests.Excel.CalcEngine;

public class FormulaExtentTests
{
    [Test]
    public async Task Measures_the_furthest_reference()
    {
        var extent = FormulaExtent.Of("SUM(B3,D2)");

        await Assert.That(extent.MaxRow).IsEqualTo(3);
        await Assert.That(extent.MaxColumn).IsEqualTo(4);
    }

    [Test]
    public async Task A_formula_the_parser_rejects_reaches_the_whole_sheet()
    {
        // A path-qualified external reference is text the parser cannot tokenize. Reporting the
        // whole sheet keeps it in front of the shifter and its fallback.
        var extent = FormulaExtent.Of(@"'C:\x\[book.xlsx]Sheet1'!A1");

        await Assert.That(extent.MaxRow).IsEqualTo(XLHelper.MaxRowNumber);
        await Assert.That(extent.MaxColumn).IsEqualTo(XLHelper.MaxColumnNumber);
    }
}
