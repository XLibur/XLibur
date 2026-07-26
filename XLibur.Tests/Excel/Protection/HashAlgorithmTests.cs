using XLibur.Utils;
using static XLibur.Excel.XLProtectionAlgorithm;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Protection;

public class HashAlgorithmTests
{
    [Test]
    public async Task TestEmptyPassword()
    {
        await Assert.That(CryptographicAlgorithms.GetPasswordHash(Algorithm.SHA512, string.Empty)).IsEmpty();
        await Assert.That(CryptographicAlgorithms.GetPasswordHash(Algorithm.SimpleHash, string.Empty)).IsEmpty();
    }

    [Test]
    public async Task TestSHA512()
    {
        var hash = CryptographicAlgorithms.GetPasswordHash(Algorithm.SHA512, "12345", "aVvPw1DNH3evPqRAd/y3UQ==", 100000);
        await Assert.That(hash).IsEqualTo("E+qAhyIg/HM0dUrPaENfimFOZp7wlOkJsf/sdG+AGHOA9grOv7VLb1ik2vuYohljI9G36e0ea9wnixCK0MMuyQ==");
    }

    [Test]
    public async Task TestSimple()
    {
        var hash = CryptographicAlgorithms.GetPasswordHash(Algorithm.SimpleHash, "12345");
        await Assert.That(hash).IsEqualTo("CA9C");
    }
}
