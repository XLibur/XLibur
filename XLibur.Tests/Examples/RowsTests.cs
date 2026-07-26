using XLibur.Examples.Rows;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class RowsTests
{
    [Test]
    public async Task RowCells()
    {
        await TestHelper.RunTestExample<RowCells>(@"Rows\RowCells.xlsx");
    }

    [Test]
    public async Task RowCollection()
    {
        await TestHelper.RunTestExample<RowCollection>(@"Rows\RowCollection.xlsx");
    }

    [Test]
    public async Task RowSettings()
    {
        await TestHelper.RunTestExample<RowSettings>(@"Rows\RowSettings.xlsx");
    }

    //[Test] // Not working yet
    public static async Task InsertRows()
    {
        await TestHelper.RunTestExample<InsertRows>(@"Rows\InsertRows.xlsx");
    }
}
