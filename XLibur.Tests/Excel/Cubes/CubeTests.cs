using NUnit.Framework;

namespace XLibur.Tests.Excel.Cubes;

[TestFixture]
public class CubeTests
{
    [Test]
    public void CalLoadAndSaveCubeFromRange()
    {
        // Disable validation, because connection type for range is 102 and validator expects at most 8.
        TestHelper.LoadAndAssert(wb =>
        {
            Assert.That(wb.Worksheets.Count, Is.GreaterThan(0));
        }, @"Other\Cubes\CubeFromRange-Input.xlsx");

        TestHelper.LoadSaveAndCompare(@"Other\Cubes\CubeFromRange-Input.xlsx", @"Other\Cubes\CubeFromRange-Output.xlsx", validate: false);
    }
}
