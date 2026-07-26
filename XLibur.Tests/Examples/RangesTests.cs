using XLibur.Examples.Ranges;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class RangesTests
{
    [Test]
    public async Task ClearingRanges()
    {
        await TestHelper.RunTestExample<ClearingRanges>(@"Ranges\ClearingRanges.xlsx");
    }

    [Test]
    public async Task CopyingRanges()
    {
        await TestHelper.RunTestExample<CopyingRanges>(@"Ranges\CopyingRanges.xlsx");
    }

    [Test]
    public async Task CurrentRowColumn()
    {
        await TestHelper.RunTestExample<CurrentRowColumn>(@"Ranges\CurrentRowColumn.xlsx");
    }

    [Test]
    public async Task DefiningRanges()
    {
        await TestHelper.RunTestExample<DefiningRanges>(@"Ranges\DefiningRanges.xlsx");
    }

    [Test]
    public async Task DeletingRanges()
    {
        await TestHelper.RunTestExample<DeletingRanges>(@"Ranges\DeletingRanges.xlsx");
    }

    [Test]
    public async Task InsertingDeletingColumns()
    {
        await TestHelper.RunTestExample<InsertingDeletingColumns>(@"Ranges\InsertingDeletingColumns.xlsx");
    }

    [Test]
    public async Task InsertingDeletingRows()
    {
        await TestHelper.RunTestExample<InsertingDeletingRows>(@"Ranges\InsertingDeletingRows.xlsx");
    }

    [Test]
    public async Task MultipleRanges()
    {
        await TestHelper.RunTestExample<MultipleRanges>(@"Ranges\MultipleRanges.xlsx");
    }

    [Test]
    public async Task DefinedNames()
    {
        await TestHelper.RunTestExample<DefinedNames>(@"Ranges\DefinedNames.xlsx");
    }

    [Test]
    public async Task SelectingRanges()
    {
        await TestHelper.RunTestExample<SelectingRanges>(@"Ranges\SelectingRanges.xlsx");
    }

    [Test]
    public async Task ShiftingRanges()
    {
        await TestHelper.RunTestExample<ShiftingRanges>(@"Ranges\ShiftingRanges.xlsx");
    }

    [Test]
    public async Task SortExample()
    {
        await TestHelper.RunTestExample<SortExample>(@"Ranges\SortExample.xlsx");
    }

    [Test]
    public async Task Sorting()
    {
        await TestHelper.RunTestExample<Sorting>(@"Ranges\Sorting.xlsx");
    }

    [Test]
    public async Task TransposeRanges()
    {
        await TestHelper.RunTestExample<TransposeRanges>(@"Ranges\TransposeRanges.xlsx");
    }

    [Test]
    public async Task TransposeRangesPlus()
    {
        await TestHelper.RunTestExample<TransposeRangesPlus>(@"Ranges\TransposeRangesPlus.xlsx");
    }

    [Test]
    public async Task AddingRowToTables()
    {
        await TestHelper.RunTestExample<AddingRowToTables>(@"Ranges\AddingRowToTables.xlsx");
    }

    [Test]
    public async Task WalkingRanges()
    {
        await TestHelper.RunTestExample<WalkingRanges>(@"Ranges\WalkingRanges.xlsx");
    }
}
