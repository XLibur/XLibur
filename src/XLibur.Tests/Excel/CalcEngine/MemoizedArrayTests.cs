using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using Array = XLibur.Excel.CalcEngine.Array;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>The cache a <c>LET</c> binds a computed array through (#847).</summary>
public class MemoizedArrayTests
{
    /// <summary>A computed array that counts how often each element is read.</summary>
    private sealed class CountingArray(int width, int height) : Array
    {
        internal int Reads { get; private set; }

        public override int Width => width;

        public override int Height => height;

        public override ScalarValue this[int y, int x]
        {
            get
            {
                Reads++;
                return y * width + x;
            }
        }
    }

    [Test]
    public async Task Each_element_is_computed_once()
    {
        var computed = new CountingArray(2, 3);
        var memoized = MemoizedArray.Of(computed);

        for (var pass = 0; pass < 3; pass++)
        {
            foreach (var _ in memoized)
            {
            }
        }

        await Assert.That(computed.Reads).IsEqualTo(6);
        await Assert.That(memoized[2, 1]).IsEqualTo(ScalarValue.From(5));
    }

    [Test]
    public async Task An_array_that_holds_its_elements_or_is_too_big_is_left_as_it_is()
    {
        var literal = new ScalarArray(ScalarValue.From(1), 2, 2);
        var huge = new CountingArray(1, MemoizedArray.MaxElements + 1);
        var memoized = MemoizedArray.Of(new CountingArray(2, 2));

        await Assert.That(MemoizedArray.Of(literal)).IsSameReferenceAs(literal);
        await Assert.That(MemoizedArray.Of(huge)).IsSameReferenceAs(huge);
        await Assert.That(MemoizedArray.Of(memoized)).IsSameReferenceAs(memoized);
    }

    [Test]
    public async Task An_element_outside_the_array_is_refused()
    {
        var memoized = MemoizedArray.Of(new CountingArray(2, 2));

        await Assert.That(() => memoized[2, 0]).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => memoized[0, -1]).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task A_LET_bound_to_a_whole_column_still_evaluates()
    {
        // Too big to keep, so it stays as lazy as it was.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("C1").SetDynamicFormulaA1("LET(x,A:A*2,SUM(x))");

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(6);
    }
}
