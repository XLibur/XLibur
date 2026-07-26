using XLibur.Examples.PageSetup;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class PageSetupTests
{
    [Test]
    public async Task HeaderFooters()
    {
        await TestHelper.RunTestExample<HeaderFooters>(@"PageSetup\HeaderFooters.xlsx");
    }

    [Test]
    public async Task Margins()
    {
        await TestHelper.RunTestExample<Margins>(@"PageSetup\Margins.xlsx");
    }

    [Test]
    public async Task Page()
    {
        await TestHelper.RunTestExample<Page>(@"PageSetup\Page.xlsx");
    }

    [Test]
    public async Task SheetTab()
    {
        await TestHelper.RunTestExample<SheetTab>(@"PageSetup\SheetTab.xlsx");
    }

    [Test]
    public async Task Sheets()
    {
        await TestHelper.RunTestExample<Sheets>(@"PageSetup\Sheets.xlsx");
    }

    [Test]
    public async Task TwoPages()
    {
        await TestHelper.RunTestExample<TwoPages>(@"PageSetup\TwoPages.xlsx");
    }
}
