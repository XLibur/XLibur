using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cubes;

public class CubeTests
{
    [Test]
    public async Task CalLoadAndSaveCubeFromRange()
    {
        // Disable validation, because connection type for range is 102 and validator expects at most 8.
        await TestHelper.LoadAndAssert(async wb =>
        {
            await Assert.That(wb.Worksheets.Count).IsGreaterThan(0);
        }, @"Other\Cubes\CubeFromRange-Input.xlsx");

        await TestHelper.LoadSaveAndCompare(@"Other\Cubes\CubeFromRange-Input.xlsx", @"Other\Cubes\CubeFromRange-Output.xlsx", validate: false);
    }
}
