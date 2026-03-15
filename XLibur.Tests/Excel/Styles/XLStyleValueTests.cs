using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.Styles;

[TestFixture]
public class XLStyleValueTests
{
    [Test]
    public void GetHashCode_SameKey_SameHash()
    {
        var key = XLStyleValue.Default.Key;
        var a = XLStyleValue.FromKey(ref key);
        var b = XLStyleValue.FromKey(ref key);

        Assert.AreEqual(a.GetHashCode(), b.GetHashCode());
    }

    [Test]
    public void GetHashCode_EqualKeys_ProduceSameInstance()
    {
        // The repository interns equal keys, so equal styles must be the same instance.
        var key = XLStyleValue.Default.Key;
        var a = XLStyleValue.FromKey(ref key);
        var b = XLStyleValue.FromKey(ref key);

        Assert.IsTrue(ReferenceEquals(a, b));
    }

    [Test]
    public void Equals_DifferentHash_ReturnsFalse()
    {
        var key1 = XLStyleValue.Default.Key with { IncludeQuotePrefix = false };
        var key2 = XLStyleValue.Default.Key with { IncludeQuotePrefix = true };
        var a = XLStyleValue.FromKey(ref key1);
        var b = XLStyleValue.FromKey(ref key2);

        Assert.IsFalse(a.Equals(b));
        Assert.AreNotEqual(a.GetHashCode(), b.GetHashCode());
    }

    [Test]
    public void Equals_DefaultStyle_IsSymmetricAndReflexive()
    {
        var s = XLStyleValue.Default;

        Assert.IsTrue(s.Equals(s));
        Assert.IsTrue(s.Equals(XLStyleValue.Default));
        Assert.IsTrue(XLStyleValue.Default.Equals(s));
    }

    [Test]
    public void Equals_Null_ReturnsFalse()
    {
        Assert.IsFalse(XLStyleValue.Default.Equals(null));
    }
}

