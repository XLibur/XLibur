using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;
using XLibur.Tests.Excel.Tables;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Extensions;

public class EnumerableExtensionsTests
{
    private static readonly int[] SkipLastSingle = [1];
    private static readonly int[] SkipLastTwo = [1, 2];

    [Test]
    public async Task CanGetItemType()
    {
        var array = Array.Empty<int>();
        await Assert.That(array.GetItemType()).IsEqualTo(typeof(int));

        var list = new List<double>();
        await Assert.That(list.GetItemType()).IsEqualTo(typeof(double));
        await Assert.That(list.AsEnumerable().GetItemType()).IsEqualTo(typeof(double));

        IEnumerable<IEnumerable> enumerable = new List<string>();
        await Assert.That(enumerable.GetItemType()).IsEqualTo(typeof(string));

        enumerable = new List<List<string>>();
        await Assert.That(enumerable.GetItemType()).IsEqualTo(typeof(List<string>));

        enumerable = new List<int[]>();
        await Assert.That(enumerable.GetItemType()).IsEqualTo(typeof(int[]));

        var anonymousIterator = new List<TablesTests.TestObjectWithoutAttributes>()
            .Select(o => new { FirstName = o.Column1, LastName = o.Column2 });

        //expectedType can be something like <>f__AnonymousType9`2[System.String,System.String]
        //but since that `9` may differ with new anonymous types declare in the assembly
        //check the beginning and the ending of the actual type
        var expectedTypeStart = "<>f__AnonymousType";
        var expectedTypeEnd = "`2[System.String,System.String]";
        var actualType = anonymousIterator.GetItemType()!.ToString();
        await Assert.That(actualType.StartsWith(expectedTypeStart)).IsTrue();
        await Assert.That(actualType.EndsWith(expectedTypeEnd)).IsTrue();

        IEnumerable<object> obj = anonymousIterator;
        actualType = obj.GetItemType()!.ToString();
        await Assert.That(actualType.StartsWith(expectedTypeStart)).IsTrue();
        await Assert.That(actualType.EndsWith(expectedTypeEnd)).IsTrue();
    }

    [Test]
    public async Task SkipLast_skips_last_element_of_enumerable()
    {
        var empty = Array.Empty<int>().SkipLast();
        await Assert.That(empty).IsEmpty();

        var oneElement = SkipLastSingle.SkipLast();
        await Assert.That(oneElement).IsEmpty();

        var twoElements = SkipLastTwo.SkipLast();
        await Assert.That(twoElements).IsEquivalentTo([1], CollectionOrdering.Matching);
    }

    [Test]
    public async Task WhereNotNull_removes_null_elements()
    {
        var source = new int?[] { 1, null, 2 };

        var result = source.WhereNotNull(x => x);

        await Assert.That(result).IsEquivalentTo([1, 2], CollectionOrdering.Matching);
    }
}
