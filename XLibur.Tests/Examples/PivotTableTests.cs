using XLibur.Examples;
using NUnit.Framework;

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
