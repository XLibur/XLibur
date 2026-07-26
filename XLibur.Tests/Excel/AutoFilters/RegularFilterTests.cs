using System;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.AutoFilters;

public class RegularFilterTests
{
    [Test]
    public async Task DateTimeGrouping_and_regular_values_can_be_used_together()
    {
        // OpenXML SDK validator considers filter and dateTimeGroup filter elements together to
        // be an error, but it isn't (XSD allows and Excel reads). Therefore, disable
        // validation for the test.
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                var autoFilter = ws.Cell("A1").InsertData(new object[]
                {
                    "Data",
                    1, 2,
                    new DateTime(2015, 7, 25, 0, 0, 0, DateTimeKind.Unspecified),
                    new DateTime(2015, 8, 25, 0, 0, 0, DateTimeKind.Unspecified),
                }).SetAutoFilter();
                autoFilter.Column(1)
                    .AddFilter(1)
                    .AddDateGroupFilter(new DateTime(2015, 8, 1, 0, 0, 0, DateTimeKind.Unspecified), XLDateTimeGrouping.Month);
            },
            async (_, ws) =>
            {
                ws.AutoFilter.Reapply();
                var dataVisibility = ws.Rows("2:5").Select(row => !row.IsHidden);
                await Assert.That(dataVisibility).IsEquivalentTo([true, false, false, true], CollectionOrdering.Matching);
            }, false);
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Regular_number_value_is_compared_as_text_against_formatted_text()
    {
        await new AutoFilterTester(f => f.AddFilter(1.5))
            .Add(1.5, true)
            .Add("1.5", false)
            .Add("1,5", true)
            .Add("1,50", false)
            .Add(1.5, nf => nf.SetNumberFormatId((int)XLPredefinedFormat.Number.PercentPrecision2), false)
            .Add(700, nf => nf.SetFormat("\"1,5\""), true)
            .AssertVisibility();
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Regular_logical_value_is_compared_as_text_against_formatted_text()
    {
        await new AutoFilterTester(f => f.AddFilter(false))
            .Add(false, true)
            .Add(0, false)
            .Add("FALSE", true)
            .Add("TRUE", false)
            .Add(true, false)
            .Add(77, nf => nf.SetFormat("\"FALSE\""), true)
            .AssertVisibility();
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Regular_error_value_is_compared_as_text_against_formatted_text()
    {
        await new AutoFilterTester(f => f.AddFilter("#VALUE!"))
            .Add(XLError.IncompatibleValue, true)
            .Add(2, false)
            .Add("#VALUE!", true)
            .AssertVisibility();
    }

    [Test]
    public async Task Pattern_is_not_interpreted_as_wildcard()
    {
        await new AutoFilterTester(f => f.AddFilter("A*"))
            .Add("A*", true)
            .Add("A", false)
            .Add("A something", false)
            .AssertVisibility();
    }
}
