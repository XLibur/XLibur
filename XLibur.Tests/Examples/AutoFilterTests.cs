using XLibur.Examples.AutoFilters;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class AutoFilterTests
{
    [Test]
    public async Task CustomAutoFilter()
    {
        await TestHelper.RunTestExample<CustomAutoFilter>(@"AutoFilter\CustomAutoFilter.xlsx");
    }

    [Test]
    public async Task DynamicAutoFilter()
    {
        await TestHelper.RunTestExample<DynamicAutoFilter>(@"AutoFilter\DynamicAutoFilter.xlsx");
    }

    [Test]
    public async Task RegularAutoFilter()
    {
        await TestHelper.RunTestExample<RegularAutoFilter>(@"AutoFilter\RegularAutoFilter.xlsx");
    }

    [Test]
    public async Task TopBottomAutoFilter()
    {
        await TestHelper.RunTestExample<TopBottomAutoFilter>(@"AutoFilter\TopBottomAutoFilter.xlsx");
    }

    [Test]
    public async Task DateTimeGroupAutoFilter()
    {
        await TestHelper.RunTestExample<DateTimeGroupAutoFilter>(@"AutoFilter\DateTimeGroupAutoFilter.xlsx");
    }
}
