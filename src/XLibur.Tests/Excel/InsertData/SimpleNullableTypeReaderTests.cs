using XLibur.Excel.InsertData;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class SimpleNullableTypeReaderTests
{
    private readonly int?[] _data = [1, 2, null];

    [Test]
    [MethodDataSource(nameof(SimpleNullableSourceNames))]
    // See SimpleTypeReaderTests.CanGetPropertyName for why dropping the generic parameter
    // preserves the original overload binding.
    public async Task CanGetPropertyName(IEnumerable data, string expected)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);
        await Assert.That(reader.GetPropertyName(0)).IsEqualTo(expected);
    }

    public static IEnumerable<Func<(IEnumerable Data, string Expected)>> SimpleNullableSourceNames()
    {
        yield return () => (new int?[] { 1, 2, null }, "Int32");
        yield return () => (new List<double?> { 1.0, 2.0, null }, "Double");
        yield return () => (new decimal?[] { 1.0m, 2.0m, null }, "Decimal");
        yield return () => (new char?[] { 'A', 'B', null }, "Char");
        yield return () => (new DateTime?[] { new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), null }, "DateTime");
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
        await Assert.That(enumerable.Last().Single()).IsEqualTo(Blank.Value);
    }
}
