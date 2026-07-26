using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Coordinates;

public class XLNameTests
{
    [Test]
    public async Task Workbook_scoped_name_is_compared_case_insensitive()
    {
        var lowerCase = new XLName("name");
        var upperCase = new XLName("NAME");

        await Assert.That(upperCase).IsEqualTo(lowerCase);
        await Assert.That(upperCase.GetHashCode()).IsEqualTo(lowerCase.GetHashCode());

        await Assert.That(new XLName("different_name")).IsNotEqualTo(lowerCase);
    }

    [Test]
    public async Task Sheet_scoped_name_is_compared_case_insensitive()
    {
        var lowerCase = new XLName("sheet", "name");
        var upperCase = new XLName("SHEET", "NAME");

        await Assert.That(upperCase).IsEqualTo(lowerCase);
        await Assert.That(upperCase.GetHashCode()).IsEqualTo(lowerCase.GetHashCode());

        await Assert.That(new XLName("Different sheet", "name")).IsNotEqualTo(lowerCase);
        await Assert.That(new XLName("sheet", "different_name")).IsNotEqualTo(lowerCase);
    }
}
