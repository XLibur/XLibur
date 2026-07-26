using XLibur.Examples.ImageHandling;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class ImageHandlingTests
{
    [Test]
    public async Task ImageAnchors()
    {
        await TestHelper.RunTestExample<ImageAnchors>(@"ImageHandling\ImageAnchors.xlsx");
    }

    [Test]
    public async Task ImageFormats()
    {
        await TestHelper.RunTestExample<ImageFormats>(@"ImageHandling\ImageFormats.xlsx");
    }
}
