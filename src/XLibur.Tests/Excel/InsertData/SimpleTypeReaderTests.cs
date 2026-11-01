using XLibur.Excel.InsertData;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class SimpleTypeReaderTests
{
    private static readonly int[] IntData = [1, 2, 3];
    private static readonly List<double> DoubleData = [1.0, 2.0, 3.0];
    private static readonly decimal[] DecimalData = [1.0m, 2.0m, 3.0m];
    private static readonly string[] StringData = ["A", "B", "C"];
    private static readonly char[] CharData = ['A', 'B', 'C'];
    private static readonly DateTime[] DateTimeData = [new(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified)];

    private readonly int[] _data = [1, 2, 3];

    [Test]
    [MethodDataSource(nameof(SimpleSourceNames))]
    // Was a generic method fed by TestCaseData(...).Returns(...). The generic parameter was
    // never load-bearing: with an unconstrained T, IEnumerable<T> only ever converted to the
    // non-generic CreateReader(IEnumerable) overload, so taking IEnumerable directly binds
    // exactly the same call. TUnit data sources return tuples rather than TestCaseData, and
    // Returns(...) becomes an explicit assertion.
    public async Task CanGetPropertyName(IEnumerable data, string expected)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);
        await Assert.That(reader.GetPropertyName(0)).IsEqualTo(expected);
    }

    public static IEnumerable<Func<(IEnumerable Data, string Expected)>> SimpleSourceNames()
    {
        yield return () => (IntData, "Int32");
        yield return () => (DoubleData, "Double");
        yield return () => (DecimalData, "Decimal");
        yield return () => (StringData, "String");
        yield return () => (CharData, "Char");
        yield return () => (DateTimeData, "DateTime");
    }

    [Test]
    public async Task CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetPropertiesCount()).IsEqualTo(1);
    }

    [Test]
    public async Task CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetRecords().Count()).IsEqualTo(3);
    }

    [Test]
    public async Task CanReadValues()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        await Assert.That(enumerable.First().Single()).IsEqualTo(1);
        await Assert.That(enumerable.Last().Single()).IsEqualTo(3);
    }
}
