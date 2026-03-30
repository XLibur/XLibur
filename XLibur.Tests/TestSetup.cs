using NUnit.Framework;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Tests;

[SetUpFixture]
public class TestSetup
{
    [OneTimeSetUp]
    public void GlobalSetup()
    {
        SixLaborsV1FontBootstrap.Register();
    }
}
