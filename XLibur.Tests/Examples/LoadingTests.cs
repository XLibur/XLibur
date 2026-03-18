using XLibur.Examples;
using NUnit.Framework;
using XLibur.Examples.Loading;

namespace XLibur.Tests.Examples;

[TestFixture]
public class LoadingTests
{
    [Test]
    public void ChangingBasicTable()
    {
        TestHelper.RunTestExample<ChangingBasicTable>(@"Loading\ChangingBasicTable.xlsx");
    }
}
