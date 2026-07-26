using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class XLStyleValueTests
{
    [Test]
    public async Task GetHashCode_SameKey_SameHash()
    {
        var key = XLStyleValue.Default.Key;
        var a = XLStyleValue.FromKey(ref key);
        var b = XLStyleValue.FromKey(ref key);

        await Assert.That(b.GetHashCode()).IsEqualTo(a.GetHashCode());
    }

    [Test]
    public async Task GetHashCode_EqualKeys_ProduceSameInstance()
    {
        // The repository interns equal keys, so equal styles must be the same instance.
        var key = XLStyleValue.Default.Key;
        var a = XLStyleValue.FromKey(ref key);
        var b = XLStyleValue.FromKey(ref key);

        await Assert.That(ReferenceEquals(a, b)).IsTrue();
    }

    [Test]
    public async Task Equals_DifferentHash_ReturnsFalse()
    {
        var key1 = XLStyleValue.Default.Key with { IncludeQuotePrefix = false };
        var key2 = XLStyleValue.Default.Key with { IncludeQuotePrefix = true };
        var a = XLStyleValue.FromKey(ref key1);
        var b = XLStyleValue.FromKey(ref key2);

        await Assert.That(a.Equals(b)).IsFalse();
        await Assert.That(b.GetHashCode()).IsNotEqualTo(a.GetHashCode());
    }

    [Test]
    public async Task Equals_DefaultStyle_IsSymmetricAndReflexive()
    {
        var s = XLStyleValue.Default;

        await Assert.That(s.Equals(s)).IsTrue();
        await Assert.That(s.Equals(XLStyleValue.Default)).IsTrue();
        await Assert.That(XLStyleValue.Default.Equals(s)).IsTrue();
    }

    [Test]
    public async Task Equals_Null_ReturnsFalse()
    {
        await Assert.That(XLStyleValue.Default.Equals(null)).IsFalse();
    }

    [Test]
    public async Task ToString_DefaultKey_ReturnsDefault()
    {
        var key = XLStyleValue.Default.Key;
        await Assert.That(key.ToString()).IsEqualTo("Default");
    }

    [Test]
    public async Task ToString_NonDefaultKey_ShowsChangedComponents()
    {
        var key = XLStyleValue.Default.Key with { IncludeQuotePrefix = true };
        var result = key.ToString();

        // Changed component should not say "Default"
        await Assert.That(result).Contains("IncludeQuotePrefix: True");

        // Unchanged components should say "Default"
        await Assert.That(result).Contains("Alignment: Default");
        await Assert.That(result).Contains("Border: Default");
        await Assert.That(result).Contains("Fill: Default");
        await Assert.That(result).Contains("Font: Default");
        await Assert.That(result).Contains("NumberFormat: Default");
        await Assert.That(result).Contains("Protection: Default");
    }
}
