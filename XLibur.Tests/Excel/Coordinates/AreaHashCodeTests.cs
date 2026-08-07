using System.Collections.Generic;
using System.Threading.Tasks;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.Coordinates;

/// <summary>
/// Pins <see cref="Area"/>'s hash distribution, which no correctness test can see.
/// </summary>
/// <remarks>
/// <c>GetHashCode</c> returned <c>FirstPoint ^ LastPoint</c>, which is zero for every single-cell
/// area — the commonest area there is. Every <c>Dictionary</c> keyed on <see cref="Area"/> put all
/// of them in one bucket, and the dependency tree's precedent map degraded to a linear scan, making
/// the tree O(N²) to build. Nothing was ever wrong with the *answers*, so the whole test suite
/// passed throughout.
/// <para>
/// Every case here asserts a <b>distribution</b> — how many distinct hashes a set of inputs
/// produces — and none asserts that a given input hashes to a given value, or that two particular
/// unequal inputs avoid colliding. Both of those would be stricter than the regression needs and
/// would fail a perfectly sound replacement: any hash is allowed to map some input to zero and to
/// collide on some pair. The thresholds sit below the input count for the same reason, so
/// incidental collisions do not fail the suite. A future change to the combining function is free
/// as long as it does not reintroduce a degenerate one.
/// </para>
/// </remarks>
public class AreaHashCodeTests
{
    [Test]
    public async Task SingleCellAreasDoNotAllHashToTheSameValue()
    {
        var hashes = new HashSet<int>();
        for (var row = 1; row <= 1_000; row++)
        {
            var point = new Point(row, 6);
            hashes.Add(new Area(point, point).GetHashCode());
        }

        // The old XOR gave exactly one distinct hash for all thousand.
        await Assert.That(hashes.Count).IsGreaterThan(900);
    }

    /// <summary>
    /// A XOR is symmetric, so it collapsed each rectangle onto its own reversal. Measured over many
    /// pairs rather than one: a symmetric combiner halves the distinct count, which is visible as a
    /// distribution and needs no claim that any particular pair avoids colliding.
    /// </summary>
    [Test]
    public async Task CornerOrderAffectsTheHash()
    {
        var hashes = new HashSet<int>();
        for (var row = 1; row <= 40; row++)
        {
            for (var col = 1; col <= 25; col++)
            {
                var first = new Point(row, col);
                var second = new Point(row + 3, col + 2);
                hashes.Add(new Area(first, second).GetHashCode());
                hashes.Add(new Area(second, first).GetHashCode());
            }
        }

        // 2,000 areas over 1,000 pairs. The old XOR gave one hash per pair, so 1,000.
        await Assert.That(hashes.Count).IsGreaterThan(1_900);
    }

    /// <summary>Equal areas must still agree, which is the part the distribution must not break.</summary>
    [Test]
    public async Task EqualAreasHashAlike()
    {
        var a = new Area(new Point(3, 4), new Point(9, 12));
        var b = new Area(new Point(3, 4), new Point(9, 12));

        await Assert.That(a.GetHashCode()).IsEqualTo(b.GetHashCode());
        await Assert.That(a).IsEqualTo(b);
    }

    /// <summary>
    /// A grid of areas of mixed shapes should spread across buckets, not pile into a few.
    /// </summary>
    [Test]
    public async Task MixedAreaShapesSpreadAcrossBuckets()
    {
        var hashes = new HashSet<int>();
        for (var row = 1; row <= 100; row++)
        {
            for (var col = 1; col <= 10; col++)
            {
                var first = new Point(row, col);
                hashes.Add(new Area(first, first).GetHashCode());
                hashes.Add(new Area(first, new Point(row, col + 4)).GetHashCode());
            }
        }

        await Assert.That(hashes.Count).IsGreaterThan(1_900);
    }
}
