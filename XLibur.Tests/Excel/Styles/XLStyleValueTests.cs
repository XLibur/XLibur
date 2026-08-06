using XLibur.Excel;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class XLStyleValueTests
{
    /// <summary>
    /// Two cells given two different number formats must keep the format each was given, even when
    /// the two format strings happen to share a hash code.
    /// </summary>
    /// <remarks>
    /// The per-style transition cache is keyed on a 32-bit hash of the component key. A lookup that
    /// compares only that hash returns whatever the colliding entry holds, so the second cell
    /// silently inherits the first cell's format.
    /// <para>
    /// The colliding pair is searched for at run time rather than hard-coded because
    /// <see cref="string.GetHashCode()"/> is randomised per process, so no literal pair collides on
    /// every run. That randomisation is also what makes the underlying defect so unpleasant in the
    /// field: an unchanged workbook can round-trip correctly many times and corrupt on the next run.
    /// </para>
    /// <para>
    /// A custom number format sets <c>NumberFormatId</c> to a constant −1, so the key's hash reduces
    /// to the hash of the format string and a birthday search over ~2^16 candidates finds a pair in
    /// milliseconds. The cap is far above the ~65k needed for an even chance, making a fruitless
    /// search vanishingly unlikely; if one ever happens the test skips rather than fails, because
    /// finding no collision is not evidence of correctness.
    /// </para>
    /// </remarks>
    [Test]
    public async Task TransitionCache_FormatsWithCollidingHashes_DoNotOverwriteEachOther()
    {
        if (!TryFindCollidingNumberFormats(out var formatA, out var formatB))
            return;

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");

        // Both start from the default style, so both transitions are cached on the same
        // XLStyleValue instance and hash to the same slot.
        var cellA = sheet.Cell(1, 1);
        var cellB = sheet.Cell(2, 1);

        cellA.Style.NumberFormat.Format = formatA;
        cellB.Style.NumberFormat.Format = formatB;

        await Assert.That(cellA.Style.NumberFormat.Format).IsEqualTo(formatA);
        await Assert.That(cellB.Style.NumberFormat.Format).IsEqualTo(formatB);
    }

    /// <summary>
    /// Searches for two distinct custom number format strings whose hash codes collide. Returns
    /// false if none is found within the cap.
    /// </summary>
    private static bool TryFindCollidingNumberFormats(out string formatA, out string formatB)
    {
        const int candidateCap = 4_000_000;

        var seen = new Dictionary<int, string>();
        for (var i = 0; i < candidateCap; i++)
        {
            var candidate = $"#,##0.00_);[Red]({i})";
            if (seen.TryGetValue(candidate.GetHashCode(), out var previous))
            {
                formatA = previous;
                formatB = candidate;
                return true;
            }

            seen[candidate.GetHashCode()] = candidate;
        }

        formatA = string.Empty;
        formatB = string.Empty;
        return false;
    }

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
