using XLibur.Examples.PivotTables;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class PivotTableTests
{
    [Test]
    public async Task PivotTables()
    {
        await TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx");
    }
}
