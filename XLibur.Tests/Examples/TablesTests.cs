using XLibur.Examples.Tables;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class TablesTests
{
    [Test]
    public async Task InsertingTables()
    {
        await TestHelper.RunTestExample<InsertingTables>(@"Tables\InsertingTables.xlsx");
    }

    [Test]
    public async Task ResizingTables()
    {
        await TestHelper.RunTestExample<ResizingTables>(@"Tables\ResizingTables.xlsx");
    }

    [Test]
    public async Task UsingTables()
    {
        await TestHelper.RunTestExample<UsingTables>(@"Tables\UsingTables.xlsx");
    }
}
