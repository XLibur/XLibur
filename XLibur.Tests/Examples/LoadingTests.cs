using XLibur.Examples.Loading;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class LoadingTests
{
    [Test]
    public async Task ChangingBasicTable()
    {
        await TestHelper.RunTestExample<ChangingBasicTable>(@"Loading\ChangingBasicTable.xlsx");
    }
}
