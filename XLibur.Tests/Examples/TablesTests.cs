using XLibur.Examples.Tables;
using NUnit.Framework;

namespace XLibur.Tests.Examples;

[TestFixture]
public class TablesTests
{
    [Test]
    public void InsertingTables()
    {
        TestHelper.RunTestExample<InsertingTables>(@"Tables\InsertingTables.xlsx");
    }

    [Test]
    public void ResizingTables()
    {
        TestHelper.RunTestExample<ResizingTables>(@"Tables\ResizingTables.xlsx");
    }

    [Test]
    public void UsingTables()
    {
        TestHelper.RunTestExample<UsingTables>(@"Tables\UsingTables.xlsx");
    }
}
