using XLibur.Examples;
using NUnit.Framework;
using XLibur.Examples.PivotTables;

namespace XLibur.Tests.Examples;

[TestFixture]
public class PivotTableTests
{
    [Test]
    public void PivotTables()
    {
        TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx");
    }
}
