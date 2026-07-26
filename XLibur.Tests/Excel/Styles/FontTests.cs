using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class FontTests
{
    private readonly XLFontKey _defaultKey = XLFontValue.Default.Key;

    [Test]
    public async Task XLFontKey_GetHashCode_IsCaseInsensitive()
    {
        var fontKey1 = _defaultKey with { FontName = "Arial" };
        var fontKey2 = _defaultKey with { FontName = "Times New Roman" };
        var fontKey3 = _defaultKey with { FontName = "TIMES NEW ROMAN" };

        await Assert.That(fontKey2.GetHashCode()).IsNotEqualTo(fontKey1.GetHashCode());
        await Assert.That(fontKey3.GetHashCode()).IsEqualTo(fontKey2.GetHashCode());
    }

    [Test]
    public async Task XLFontKey_Equals_IsCaseInsensitive()
    {
        var fontKey1 = _defaultKey with { FontName = "Arial" };
        var fontKey2 = _defaultKey with { FontName = "Times New Roman" };
        var fontKey3 = _defaultKey with { FontName = "TIMES NEW ROMAN" };

        await Assert.That(fontKey1.Equals(fontKey2)).IsFalse();
        await Assert.That(fontKey2.Equals(fontKey3)).IsTrue();
    }
}
