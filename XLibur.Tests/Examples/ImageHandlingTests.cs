using XLibur.Examples;
using NUnit.Framework;
using XLibur.Examples.ImageHandling;

namespace XLibur.Tests.Examples;

[TestFixture]
public class ImageHandlingTests
{
    [Test]
    public void ImageAnchors()
    {
        TestHelper.RunTestExample<ImageAnchors>(@"ImageHandling\ImageAnchors.xlsx");
    }

    [Test]
    public void ImageFormats()
    {
        TestHelper.RunTestExample<ImageFormats>(@"ImageHandling\ImageFormats.xlsx");
    }
}
