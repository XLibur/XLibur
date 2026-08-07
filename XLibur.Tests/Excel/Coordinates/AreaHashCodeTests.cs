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
/// These tests assert the distribution rather than any particular hash value, so a future change to
/// the combining function is free as long as it does not reintroduce a self-cancelling one.
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

    [Test]
    public async Task SingleCellAreaDoesNotHashToZero()
    {
        var point = new Point(42, 7);

        await Assert.That(new Area(point, point).GetHashCode()).IsNotEqualTo(0);
    }

    /// <summary>
    /// A XOR is symmetric, so it also collapsed the two corners of a rectangle onto each other.
    /// Only the normalised order is ever constructed, but the hash should still tell them apart.
    /// </summary>
    [Test]
    public async Task SwappedCornersDoNotCollide()
    {
        var a = new Area(new Point(1, 1), new Point(2, 2));
        var b = new Area(new Point(2, 2), new Point(1, 1));

        await Assert.That(a.GetHashCode()).IsNotEqualTo(b.GetHashCode());
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
