using XLibur.Excel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class SearchTests
{
    [Test]
    public async Task TestSearch()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\CellValues.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();

        var foundCells = ws.Search("Initial Value");
        await Assert.That(foundCells.Count()).IsEqualTo(1);
        await Assert.That(foundCells.Single().Address.ToString()).IsEqualTo("B2");
        await Assert.That(foundCells.Single().GetText()).IsEqualTo("Initial Value");

        foundCells = ws.Search("Using");
        await Assert.That(foundCells.Count()).IsEqualTo(2);
        await Assert.That(foundCells.First().Address.ToString()).IsEqualTo("D2");
        await Assert.That(foundCells.First().GetText()).IsEqualTo("Using Get...()");
        await Assert.That(foundCells.Count()).IsEqualTo(2);
        await Assert.That(foundCells.Last().Address.ToString()).IsEqualTo("E2");
        await Assert.That(foundCells.Last().GetText()).IsEqualTo("Using GetValue<T>()");

        foundCells = ws.Search("1234");
        await Assert.That(foundCells.Count()).IsEqualTo(5);
        await Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())).IsEqualTo("B5,C5,D5,E5,F5");

        foundCells = ws.Search("Sep");
        await Assert.That(foundCells.Count()).IsEqualTo(1);
        await Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())).IsEqualTo("G3");

        foundCells = ws.Search("1234", CompareOptions.Ordinal, true);
        await Assert.That(foundCells.Count()).IsEqualTo(5);
        await Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())).IsEqualTo("B5,C5,D5,E5,F5");

        foundCells = ws.Search("test case");
        await Assert.That(foundCells.Count()).IsEqualTo(0);

        foundCells = ws.Search("test case", CompareOptions.OrdinalIgnoreCase);
        await Assert.That(foundCells.Count()).IsEqualTo(6);
    }

    [Test]
    public async Task TestSearch2()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();

        var foundCells = ws.Search("3");
        await Assert.That(foundCells.Count()).IsEqualTo(10);
        await Assert.That(foundCells.First().Address.ToString()).IsEqualTo("C2");

        foundCells = ws.Search("A2", CompareOptions.Ordinal, true);
        await Assert.That(foundCells.Count()).IsEqualTo(6);
        await Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())).IsEqualTo("C2,D2,B6,C6,D6,A11");

        foundCells = ws.Search("RC", CompareOptions.Ordinal, true);
        await Assert.That(foundCells.Count()).IsEqualTo(3);
        await Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())).IsEqualTo("E2,E3,E4");
    }
}
