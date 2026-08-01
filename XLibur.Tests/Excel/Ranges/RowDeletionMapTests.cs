using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Unit tests for the row map a batched delete is built on.
/// <para>
/// <see cref="BatchRowDeleteTests"/> covers the observable behaviour, but two properties here are
/// invisible from outside and would degrade silently. Run coalescing is one: emitting singletons
/// instead of runs is still <em>correct</em>, just one structural pass per row instead of per run,
/// which is most of what the batching is for. The edge-mapping asymmetry is the other.
/// </para>
/// </summary>
public class RowDeletionMapTests
{
    [Test]
    public async Task ConsecutiveRowsCoalesceIntoOneRun()
    {
        var map = XLRowDeletionMap.Create([8, 9, 10, 11, 12])!;

        var runs = map.GetRunsBottomUp();

        await Assert.That(runs.Count).IsEqualTo(1);
        await Assert.That(runs[0]).IsEqualTo((8, 12));
    }

    /// <summary>
    /// Runs come out furthest down the sheet first: a run has to be removed before the runs above it
    /// move, or their row numbers stop meaning what the map says.
    /// </summary>
    [Test]
    public async Task RunsAreEmittedBottomUp()
    {
        var map = XLRowDeletionMap.Create([4, 5, 12, 19, 20, 21, 30])!;

        var runs = map.GetRunsBottomUp();

        await Assert.That(runs).IsEquivalentTo([(30, 30), (19, 21), (12, 12), (4, 5)]);
    }

    [Test]
    public async Task UnsortedInputWithDuplicatesCollapses()
    {
        var map = XLRowDeletionMap.Create([12, 4, 5, 4, 12])!;

        await Assert.That(map.Count).IsEqualTo(3);
        await Assert.That(map.FirstDeletedRow).IsEqualTo(4);
        await Assert.That(map.GetRunsBottomUp()).IsEquivalentTo([(12, 12), (4, 5)]);
    }

    [Test]
    public async Task EmptyInputProducesNoMap()
    {
        await Assert.That(XLRowDeletionMap.Create([])).IsNull();
    }

    /// <summary>
    /// The two edges are counted differently on purpose. A range's top slides up by the rows deleted
    /// strictly above it; its bottom also loses the rows deleted inside it. Mapping both ends the same
    /// way would leave a range that spans a deleted row too long by exactly that count.
    /// </summary>
    [Test]
    public async Task TopAndBottomEdgesMapDifferentlyAcrossAnInteriorDeletion()
    {
        var map = XLRowDeletionMap.Create([2, 12, 15])!;

        // A10:A20 -> one row gone above it, two gone inside it.
        await Assert.That(map.MapFirst(10)).IsEqualTo(9);
        await Assert.That(map.MapLast(20)).IsEqualTo(17);
    }

    [Test]
    public async Task RowsAboveEveryDeletionAreUnmoved()
    {
        var map = XLRowDeletionMap.Create([10, 20, 30])!;

        await Assert.That(map.MapFirst(5)).IsEqualTo(5);
        await Assert.That(map.MapLast(9)).IsEqualTo(9);
    }

    /// <summary>
    /// A range wholly inside the deletion maps to an inverted pair, which is how the shifter recognises
    /// a reference the deletion destroyed.
    /// </summary>
    [Test]
    public async Task FullySwallowedRangeMapsToAnInvertedPair()
    {
        var map = XLRowDeletionMap.Create([10, 11, 12])!;

        await Assert.That(map.MapLast(12)).IsLessThan(map.MapFirst(10));
    }

    /// <summary>
    /// A deleted row shares its mapped position with the first survivor below it — the survivor moves
    /// into the space the deleted row left.
    /// </summary>
    [Test]
    public async Task DeletedRowMapsToWhereItsSuccessorLands()
    {
        var map = XLRowDeletionMap.Create([5])!;

        await Assert.That(map.MapFirst(5)).IsEqualTo(5);
        await Assert.That(map.MapFirst(6)).IsEqualTo(5);
    }
}
