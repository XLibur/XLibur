using XLibur.Excel.CalcEngine.Functions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class XLMathTests
{
    [Test]
    public async Task IsEven()
    {
        await Assert.That(XLMath.IsEven(2)).IsTrue();
        await Assert.That(XLMath.IsEven(3)).IsFalse();
    }

    [Test]
    public async Task IsOdd()
    {
        await Assert.That(XLMath.IsOdd(3)).IsTrue();
        await Assert.That(XLMath.IsOdd(2)).IsFalse();
    }
}
