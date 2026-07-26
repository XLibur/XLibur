using XLibur.Excel.InsertData;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Tests.Excel.Tables;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class ObjectReaderTests
{
    private static readonly TablesTests.TestObjectWithAttributes[] ObjectWithAttributes =
    [
        new()
        {
            Column1 = "Value 1",
            Column2 = "Value 2",
            UnOrderedColumn = 3,
            MyField = 4,
        },
        new()
        {
            Column1 = "Value 5",
            Column2 = "Value 6",
            UnOrderedColumn = 7,
            MyField = 8,
        }
    ];

    private static readonly TablesTests.TestObjectWithoutAttributes[] ObjectWithoutAttributes =
    [
        new()
        {
            Column1 = "Value 9",
            Column2 = "Value 10"
        },
        new()
        {
            Column1 = "Value 11",
            Column2 = "Value 12"
        }
    ];

    private static readonly TestPoint[] Structs =
    [
        new()
        {
            X = 1,
            Y = 2,
            Z = 3
        },
        new()
    ];

    private static readonly TestPoint?[] NullableStructs =
    [
        new TestPoint
        {
            X = 1,
            Y = 2,
            Z = 3
        },
        new TestPoint(),
        null
    ];

    [Test]
    [MethodDataSource(nameof(ObjectSourceNames))]
    // The data source already handed these over as non-generic IEnumerable, so the generic
    // parameter was decorative; TUnit data sources return tuples instead of TestCaseData.
    public async Task CanGetPropertyName(IEnumerable data, int propertyIndex, string expected)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);
        await Assert.That(reader.GetPropertyName(propertyIndex)).IsEqualTo(expected);
    }

    public static IEnumerable<Func<(IEnumerable Data, int PropertyIndex, string Expected)>> ObjectSourceNames()
    {
        yield return () => (ObjectWithoutAttributes, 0, "Column1");
        yield return () => (ObjectWithoutAttributes, 1, "Column2");

        yield return () => (ObjectWithAttributes, 0, "FirstColumn");
        yield return () => (ObjectWithAttributes, 1, "SecondColumn");
        yield return () => (ObjectWithAttributes, 2, "SomeFieldNotProperty");
        yield return () => (ObjectWithAttributes, 3, "UnOrderedColumn");

        yield return () => (Structs, 0, "X");
        yield return () => (Structs, 1, "Y");
        yield return () => (Structs, 2, "Z");

        yield return () => (NullableStructs, 0, "X");
        yield return () => (NullableStructs, 1, "Y");
        yield return () => (NullableStructs, 2, "Z");
    }

    [Test]
    [MethodDataSource(nameof(PropertyCounts))]
    public async Task CanGetPropertiesCount(IEnumerable data, int expected)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);
        await Assert.That(reader.GetPropertiesCount()).IsEqualTo(expected);
    }

    public static IEnumerable<Func<(IEnumerable Data, int Expected)>> PropertyCounts()
    {
        yield return () => (ObjectWithoutAttributes, 2);
        yield return () => (ObjectWithAttributes, 4);
        yield return () => (Structs, 3);
        yield return () => (NullableStructs, 3);
    }

    [Test]
    public async Task CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(ObjectWithAttributes);
        await Assert.That(reader.GetRecords().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task CanReadValues_FromObject()
    {
        var reader = InsertDataReaderFactory.CreateReader(ObjectWithAttributes);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        var firstRecord = enumerable.First().ToArray();
        var lastRecord = enumerable.Last().ToArray();

        await Assert.That(firstRecord[0]).IsEqualTo("Value 2");
        await Assert.That(firstRecord[1]).IsEqualTo("Value 1");
        await Assert.That(firstRecord[2]).IsEqualTo(4);
        await Assert.That(firstRecord[3]).IsEqualTo(3);

        await Assert.That(lastRecord[0]).IsEqualTo("Value 6");
        await Assert.That(lastRecord[1]).IsEqualTo("Value 5");
        await Assert.That(lastRecord[2]).IsEqualTo(8);
        await Assert.That(lastRecord[3]).IsEqualTo(7);
    }

    [Test]
    public async Task CanReadValues_FromStruct()
    {
        var reader = InsertDataReaderFactory.CreateReader(Structs);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        var firstRecord = enumerable.First().ToArray();
        var lastRecord = enumerable.Last().ToArray();

        await Assert.That(firstRecord[0]).IsEqualTo(1);
        await Assert.That(firstRecord[1]).IsEqualTo(2);
        await Assert.That(firstRecord[2]).IsEqualTo(3);

        await Assert.That(lastRecord[0]).IsEqualTo(0);
        await Assert.That(lastRecord[1]).IsEqualTo(0);
        await Assert.That(lastRecord[2]).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task CanReadValues_FromNullableStruct()
    {
        var reader = InsertDataReaderFactory.CreateReader(NullableStructs);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        var firstRecord = enumerable.First().ToArray();
        var lastRecord = enumerable.Last().ToArray();

        await Assert.That(firstRecord[0]).IsEqualTo(1);
        await Assert.That(firstRecord[1]).IsEqualTo(2);
        await Assert.That(firstRecord[2]).IsEqualTo(3);

        await Assert.That(lastRecord[0]).IsEqualTo(Blank.Value);
        await Assert.That(lastRecord[1]).IsEqualTo(Blank.Value);
        await Assert.That(lastRecord[2]).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task IgnoresIndexers()
    {
        var data = new[] { new TestClassWithIndexer() };
        var reader = InsertDataReaderFactory.CreateReader(data);

        await Assert.That(reader.GetPropertiesCount()).IsEqualTo(1);
        await Assert.That(reader.GetPropertyName(0)).IsEqualTo(nameof(TestClassWithIndexer.Value));
    }

    private record TestClassWithIndexer
    {
        public static int Value => 0;
    }

    private struct TestPoint
    {
        public double X { get; set; }

        public double Y { get; set; }

        public double? Z { get; set; }
    }
}
