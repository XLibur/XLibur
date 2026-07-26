using XLibur.Examples.Columns;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class ColumnsTests
{
    [Test]
    public async Task ColumnCells()
    {
        await TestHelper.RunTestExample<ColumnCells>(@"Columns\ColumnCells.xlsx");
    }

    [Test]
    public async Task ColumnCollections()
    {
        await TestHelper.RunTestExample<ColumnCollection>(@"Columns\ColumnCollection.xlsx");
    }

    [Test]
    public async Task ColumnSettings()
    {
        await TestHelper.RunTestExample<ColumnSettings>(@"Columns\ColumnSettings.xlsx");
    }

    [Test]
    public async Task DeletingColumns()
    {
        await TestHelper.RunTestExample<DeletingColumns>(@"Columns\DeletingColumns.xlsx");
    }

    //[Test] // Not working yet
    public static async Task InsertColumns()
    {
        await TestHelper.RunTestExample<InsertColumns>(@"Columns\InsertColumns.xlsx");
    }
}
