using System;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.AutoFilters;

public class Top10FilterTests
{
    [Test]
    public async Task Top10_filter_is_initialized_after_load()
    {
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                var autoFilter = ws.Cell("A1").InsertData(new object[]
                {
                    "Data",
                    4, 4, 1, 3, 2, 5,
                }).SetAutoFilter();
                autoFilter.Column(1).Top(3);
            },
            async (_, ws) =>
            {
                ws.AutoFilter.Reapply();
                var filterResult = ws.Rows("2:7").Select(row => !row.IsHidden);
                await Assert.That(filterResult).IsEquivalentTo([true, true, false, false, false, true], CollectionOrdering.Matching);
            });
    }

    [Test]
    public async Task Top_items_filter_excludes_non_unified_numbers()
    {
        // Sort and then use cutoff value, it's 4 here and then take all values >= cutoff.
        await new AutoFilterTester(f => f.Top(1))
            .AddTrue(new DateTime(1900, 2, 10, 0, 0, 0, DateTimeKind.Unspecified))
            .AddFalse(11, 10)
            .AddFalse("-1000", "Text", Blank.Value, true, false, XLError.IncompatibleValue)
            .AssertVisibility();
    }

    [Test]
    public async Task Bottom_items_filter_excludes_non_unified_numbers()
    {
        await new AutoFilterTester(f => f.Bottom(1))
            .AddTrue(new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified))
            .AddFalse(2, 3)
            .AddFalse("-1000", "Text", Blank.Value, true, false, XLError.IncompatibleValue)
            .AssertVisibility();
    }

    [Test]
    public async Task Top_items_filter_determines_top_items_by_determining_cut_off_value()
    {
        // Sort and then use cutoff value, it's 4 here and then take all values <= cutoff.
        await new AutoFilterTester(f => f.Top(2))
            .AddTrue(5, 4, 4, 4)
            .AddFalse(3, 2, 1)
            .AssertVisibility();

        // Cutoff is 5 here.
        await new AutoFilterTester(f => f.Top(2))
            .AddTrue(5, 5)
            .AddFalse(4, 4, 4, 3, 2, 1)
            .AssertVisibility();
    }

    [Test]
    public async Task Bottom_items_filter_determines_top_items_by_determining_cut_off_value()
    {
        // Cutoff is 2
        await new AutoFilterTester(f => f.Bottom(2))
            .AddTrue(1, 2, 2, 2)
            .AddFalse(3, 4, 5)
            .AssertVisibility();

        // Cutoff is 5
        await new AutoFilterTester(f => f.Bottom(2))
            .AddTrue(1, 1)
            .AddFalse(2, 2, 2, 3, 4, 5)
            .AssertVisibility();
    }

    [Test]
    public async Task Top_percents_uses_inclusive_percent_value()
    {
        // Autofilter doesn't include value 750, which is at 75%, i.e. right at the border.
        await new AutoFilterTester(f => f.Top(25, XLTopBottomType.Percent))
            .AddFalse(Enumerable.Range(1, 750).Select<int, XLCellValue>(x => x).ToArray())
            .AddTrue(Enumerable.Range(751, 250).Select<int, XLCellValue>(x => x).ToArray())
            .AssertVisibility();
    }

    [Test]
    public async Task Bottom_percents_uses_inclusive_percent_value()
    {
        await new AutoFilterTester(f => f.Bottom(25, XLTopBottomType.Percent))
            .AddTrue(Enumerable.Range(1, 250).Select<int, XLCellValue>(x => x).ToArray())
            .AddFalse(Enumerable.Range(251, 750).Select<int, XLCellValue>(x => x).ToArray())
            .AssertVisibility();
    }

    [Test]
    public async Task Top_percents_always_has_at_least_one_item()
    {
        // Top 1% takes one item that is 33% of all items.
        await new AutoFilterTester(f => f.Top(1, XLTopBottomType.Percent))
            .AddTrue(3)
            .AddFalse(2, 1)
            .AssertVisibility();
    }

    [Test]
    public async Task Bottom_percents_always_has_at_least_one_item()
    {
        await new AutoFilterTester(f => f.Bottom(1, XLTopBottomType.Percent))
            .AddTrue(1)
            .AddFalse(2, 3)
            .AssertVisibility();
    }

    [Test]
    [Arguments(0, true)]
    [Arguments(501, true)]
    [Arguments(0, false)]
    [Arguments(501, false)]
    public async Task Top_and_bottom_filter_value_must_be_between_1_and_500(int value, bool top)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Data";
        ws.Cell("A2").Value = value;
        var autoFilter = ws.Range("A1:A2").SetAutoFilter();
        var filterColumn = autoFilter.Column(1);

        var ex = await Assert.That(() =>
        {
            if (top)
                filterColumn.Top(value);
            else
                filterColumn.Bottom(value);
        }).Throws<ArgumentOutOfRangeException>()!;
        await Assert.That(ex.Message).Contains("Value must be between 1 and 500.");
    }
}
