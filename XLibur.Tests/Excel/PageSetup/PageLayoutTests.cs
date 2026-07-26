using System.Threading.Tasks;

namespace XLibur.Tests.Excel.PageSetup;

public class PageLayoutTests
{
    [Test]
    public async Task FirstPageNumber_can_be_negative()
    {
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) => ws.PageSetup.FirstPageNumber = -3,
            async (_, ws) => await Assert.That(ws.PageSetup.FirstPageNumber).IsEqualTo(-3),
            @"Other\PageSetup\Negative_first_page_number.xlsx");
    }
}
