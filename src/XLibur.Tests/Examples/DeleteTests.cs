using XLibur.Examples.Delete;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class DeleteTests
{
    [Test]
    public async Task DeleteFewWorksheets()
    {
        await TestHelper.RunTestExample<DeleteFewWorksheets>(@"Delete\DeleteFewWorksheets.xlsx");
    }

    [Test]
    public async Task RemoveRows()
    {
        await TestHelper.RunTestExample<DeleteRows>(@"Delete\RemoveRows.xlsx");
    }
}
