using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.CalcEngine;

public class XLCalculationChainTests
{
    [Test]
    public async Task Enumerating_empty_chain()
    {
        var chain = new XLCalculationChain();
        await Assert.That(GetPoints(chain)).IsEmpty();
    }

    [Test]
    [Arguments(1)]
    [Arguments(2)]
    [Arguments(3)]
    [Arguments(40)]
    public async Task Enumerating_whole_chain(int chainLength)
    {
        var chain = new XLCalculationChain();
        var expectedPoints = new List<SheetPoint>();
        for (var i = 0; i < chainLength; ++i)
        {
            var point = new SheetPoint(1, new Point(1, i));
            chain.AddLast(point);
            expectedPoints.Add(point);
        }

        await Assert.That(GetPoints(chain)).IsEquivalentTo(expectedPoints, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Remove_throws_on_missing_point()
    {
        var chain = new XLCalculationChain();

        await Assert.That(() => chain.Remove(new SheetPoint(1, new Point(1, 1)))).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task Remove_link_from_chain()
    {
        var chain = new XLCalculationChain();
        var a1 = new SheetPoint(1, new Point(1, 1));
        var b1 = new SheetPoint(1, new Point(1, 2));
        var c1 = new SheetPoint(1, new Point(1, 3));
        var d1 = new SheetPoint(1, new Point(1, 4));

        chain.AddLast(a1);
        chain.AddLast(b1);
        chain.AddLast(c1);
        chain.AddLast(d1);

        // Remove point in the middle
        chain.Remove(c1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, b1, d1], CollectionOrdering.Matching);

        // Remove last point in the sequence
        chain.Remove(d1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, b1], CollectionOrdering.Matching);

        // Remove head
        chain.Remove(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([b1], CollectionOrdering.Matching);

        // Remove the only remaining
        chain.Remove(b1);
        await Assert.That(GetPoints(chain)).IsEmpty();
    }

    [Test]
    public async Task AddAfter_adds_point()
    {
        var chain = new XLCalculationChain();
        var a1 = new SheetPoint(1, new Point(1, 1));
        chain.AddLast(a1);

        // Add as tail for single link chain
        var b1 = new SheetPoint(1, new Point(1, 2));
        chain.AddAfter(a1, b1, 0);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, b1], CollectionOrdering.Matching);

        // Add as tail for multi link chain
        var c1 = new SheetPoint(1, new Point(1, 3));
        chain.AddAfter(b1, c1, 0);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, b1, c1], CollectionOrdering.Matching);

        // Add somewhere in the middle
        var d1 = new SheetPoint(1, new Point(1, 4));
        chain.AddAfter(b1, d1, 0);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, b1, d1, c1], CollectionOrdering.Matching);
    }

    [Test]
    public async Task MoveToFront_moves_the_point_to_the_front()
    {
        var chain = new XLCalculationChain();
        var a1 = new SheetPoint(1, new Point(1, 1));
        chain.AddLast(a1);
        var b1 = new SheetPoint(1, new Point(1, 2));
        chain.AddLast(b1);
        var c1 = new SheetPoint(1, new Point(1, 3));
        chain.AddLast(c1);
        var d1 = new SheetPoint(1, new Point(1, 4));
        chain.AddLast(d1);

        await Assert.That(chain.MoveAhead()).IsTrue();
        await Assert.That(chain.Current).IsEqualTo(a1);

        // a,b,c,d -> d,a,b,c
        chain.MoveToCurrent(d1);
        await Assert.That(chain.Current).IsEqualTo(d1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { d1, a1, b1, c1 }, CollectionOrdering.Matching);

        // d,a,b,c -> b,d,a,c
        chain.MoveToCurrent(b1);
        await Assert.That(chain.Current).IsEqualTo(b1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, d1, a1, c1 }, CollectionOrdering.Matching);

        await Assert.That(chain.MoveAhead()).IsTrue();
        await Assert.That(chain.Current).IsEqualTo(d1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, d1, a1, c1 }, CollectionOrdering.Matching);

        // d,a,c -> a,d,c
        chain.MoveToCurrent(a1);
        await Assert.That(chain.Current).IsEqualTo(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, a1, d1, c1 }, CollectionOrdering.Matching);

        // Move A1 to front when it's already at front
        chain.MoveToCurrent(a1);
        await Assert.That(chain.Current).IsEqualTo(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, a1, d1, c1 }, CollectionOrdering.Matching);

        // a,d,c -> c,a,d
        chain.MoveToCurrent(c1);
        await Assert.That(chain.Current).IsEqualTo(c1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, c1, a1, d1 }, CollectionOrdering.Matching);

        await Assert.That(chain.MoveAhead()).IsTrue();
        await Assert.That(chain.Current).IsEqualTo(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, c1, a1, d1 }, CollectionOrdering.Matching);

        // a,d -> d,a
        chain.MoveToCurrent(d1);
        await Assert.That(chain.Current).IsEqualTo(d1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, c1, d1, a1 }, CollectionOrdering.Matching);

        await Assert.That(chain.MoveAhead()).IsTrue();
        await Assert.That(chain.Current).IsEqualTo(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, c1, d1, a1 }, CollectionOrdering.Matching);

        // a -> a
        chain.MoveToCurrent(a1);
        await Assert.That(chain.Current).IsEqualTo(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, c1, d1, a1 }, CollectionOrdering.Matching);

        await Assert.That(chain.MoveAhead()).IsFalse();
        await Assert.That(GetPoints(chain)).IsEquivalentTo(new[] { b1, c1, d1, a1 }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Traversal_detects_cycles()
    {
        var chain = new XLCalculationChain();
        // `=C1+B1`
        var a1 = new SheetPoint(1, new Point(1, 1));
        chain.AddLast(a1);
        // `=A1`
        var b1 = new SheetPoint(1, new Point(1, 2));
        chain.AddLast(b1);
        // `=A1`
        var c1 = new SheetPoint(1, new Point(1, 3));
        chain.AddLast(c1);

        // Move to the first link.
        await Assert.That(chain.MoveAhead()).IsTrue();

        // Cycle a1, c1, when we first encounter c1, we don't know yet that it's a cycle
        chain.MoveToCurrent(c1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, a1, b1], CollectionOrdering.Matching);

        // A1 is marked with a position, because they have been at the current
        // C1 hasn't ben pushed back yet, so it keeps 0.
        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 1, 0], CollectionOrdering.Matching);

        // But then we get A1 again, without any other point being marked
        // as done, therefore we are at cycle.
        chain.MoveToCurrent(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, c1, b1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([1, 1, 0], CollectionOrdering.Matching);
        await Assert.That(chain.IsCurrentInCycle).IsTrue();

        // When we encounter C1 again, it's obviously a cycle.
        chain.MoveToCurrent(c1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, a1, b1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([1, 1, 0], CollectionOrdering.Matching);
        await Assert.That(chain.IsCurrentInCycle).IsTrue();

        // Let's move on and get A1 to the current. Because the C1 has been
        // marked as done, A1 is no longer in cycle.
        chain.MoveAhead();
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, a1, b1], CollectionOrdering.Matching);

        // C1 position has been cleared, because it has moved beyond
        // current and A1 is now current.
        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 1, 0], CollectionOrdering.Matching);

        // A1 is no longer in a current, because current position is 2, but last position
        // of A1 was 1 => there has been a processed node in the meantime.
        await Assert.That(chain.IsCurrentInCycle).IsFalse();

        chain.MoveToCurrent(b1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, b1, a1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 0, 2], CollectionOrdering.Matching);
        await Assert.That(chain.IsCurrentInCycle).IsFalse();

        chain.MoveToCurrent(a1);
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, a1, b1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 2, 2], CollectionOrdering.Matching);
        await Assert.That(chain.IsCurrentInCycle).IsTrue();

        chain.MoveAhead();
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, a1, b1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 0, 2], CollectionOrdering.Matching);
        await Assert.That(chain.IsCurrentInCycle).IsFalse();

        chain.MoveAhead();
        await Assert.That(GetPoints(chain)).IsEquivalentTo([c1, a1, b1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 0, 0], CollectionOrdering.Matching);
    }

    [Test]
    public async Task Reset_clears_positions_ahead_of_current()
    {
        var chain = new XLCalculationChain();
        var a1 = new SheetPoint(1, new Point(1, 1));
        chain.AddLast(a1);
        var b1 = new SheetPoint(1, new Point(1, 2));
        chain.AddLast(b1);
        var c1 = new SheetPoint(1, new Point(1, 3));
        chain.AddLast(c1);

        await Assert.That(chain.MoveAhead()).IsTrue();

        chain.MoveToCurrent(b1);
        chain.MoveToCurrent(a1);
        await Assert.That(chain.IsCurrentInCycle).IsTrue();
        await Assert.That(GetPoints(chain)).IsEquivalentTo([a1, b1, c1], CollectionOrdering.Matching);
        await Assert.That(GetPositions(chain)).IsEquivalentTo([1, 1, 0], CollectionOrdering.Matching);

        chain.Reset();

        await Assert.That(GetPositions(chain)).IsEquivalentTo([0, 0, 0], CollectionOrdering.Matching);
    }

    private static IEnumerable<SheetPoint> GetPoints(XLCalculationChain chain)
    {
        return chain.GetLinks().Select(x => x.Point);
    }

    private static IEnumerable<int> GetPositions(XLCalculationChain chain)
    {
        return chain.GetLinks().Select(x => x.LastPosition);
    }
}
