using System;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.AutoFilters;

public class DynamicFilterTests
{
    [Test]
    public async Task Average_filter_is_initialized_after_load()
    {
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                var autoFilter = ws.Cell("A1").InsertData(new object[]
                {
                    "Data",
                    1,2,3,4,5,10, // avg. 4.16
                }).SetAutoFilter();
                autoFilter.Column(1).AboveAverage();
            },
            async (_, ws) =>
            {
                ws.AutoFilter.Reapply();
                var filterResult = ws.Rows("2:7").Select(row => !row.IsHidden);
                await Assert.That(filterResult).IsEquivalentTo([false, false, false, false, true, true], CollectionOrdering.Matching);
            });
    }

    [Test]
    public async Task BelowAverage_takes_values_under_avg_value()
    {
        // The average 2 is not included.
        await new AutoFilterTester(f => f.BelowAverage())
            .AddTrue(1)
            .AddFalse(2, 3)
            .AssertVisibility();
    }

    [Test]
    public async Task AboveAverage_takes_values_over_avg_value()
    {
        await new AutoFilterTester(f => f.AboveAverage())
            .AddTrue(3)
            .AddFalse(2, 1)
            .AssertVisibility();
    }

    [Test]
    public async Task Average_ignores_non_unified_numbers()
    {
        await new AutoFilterTester(f => f.BelowAverage())
            .AddTrue(new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified)) // Serial date time 1
            .AddFalse(1.1)
            .AddFalse(1.2)
            .AddFalse(XLError.NoValueAvailable, true, false, "-100", "Text", Blank.Value)
            .AssertVisibility();
    }

    [Test]
    public async Task All_rows_are_hidden_when_column_has_no_number()
    {
        await new AutoFilterTester(f => f.AboveAverage())
            .AddFalse(Blank.Value, true, false, "-100", "Text", XLError.NoValueAvailable)
            .AssertVisibility();
    }
}
