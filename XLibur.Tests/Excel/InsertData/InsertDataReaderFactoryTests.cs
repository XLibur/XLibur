using XLibur.Excel.InsertData;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class InsertDataReaderFactoryTests
{
    [Test]
    public async Task CanInstantiateFactory()
    {
        var factory = InsertDataReaderFactory.Instance;

        await Assert.That(factory).IsNotNull();
        await Assert.That(InsertDataReaderFactory.Instance).IsSameReferenceAs(factory);
    }

    [Test]
    [MethodDataSource(nameof(SimpleSources))]
    public async Task CanCreateSimpleReader(IEnumerable data)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);

        await Assert.That(reader).IsAssignableTo<SimpleTypeReader>();
    }

    public static IEnumerable<object> SimpleSources
    {
        get
        {
            yield return new[] { 1, 2, 3 };
            yield return new List<double> { 1.0, 2.0, 3.0 };
            yield return new[] { "A", "B", "C" };
            yield return new[] { "A", "B", "C" };
            yield return new[] { 'A', 'B', 'C' };
        }
    }

    [Test]
    [MethodDataSource(nameof(SimpleNullableSources))]
    public async Task CanCreateSimpleNullableReader(IEnumerable data)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);

        await Assert.That(reader).IsAssignableTo<SimpleNullableTypeReader>();
    }

    public static IEnumerable<object> SimpleNullableSources
    {
        get
        {
            yield return new int?[] { 1, 2, null };
            yield return new List<double?> { 1.0, 2.0, null };
            yield return new char?[] { 'A', 'B', null };
            yield return new DateTime?[] { DateTime.MinValue, DateTime.MaxValue, null };
        }
    }

    [Test]
    [MethodDataSource(nameof(ArraySources))]
    // Was generic. TUnit's source generator could not resolve the type argument for a
    // generic test fed by this data source and silently produced no test cases at all, so
    // the four cases disappeared from the run. The generic parameter was never load-bearing:
    // with an unconstrained T, IEnumerable<T> only converts to the non-generic
    // CreateReader(IEnumerable) overload, which is what this bound to before.
    public async Task CanCreateArrayReader(IEnumerable data)
    {
        var reader = InsertDataReaderFactory.CreateReader(data);

        await Assert.That(reader).IsAssignableTo<ArrayReader>();
    }

    public static IEnumerable<object[]> ArraySources
    {
        get
        {
            yield return
            [
                new int[][]
                {
                    [1, 2, 3],
                    [4, 5, 6]
                }
            ];
            yield return [new List<List<double>> { new List<double> { 1.0, 2.0, 3.0 } }];
            yield return
            [
                (new int[][]
                {
                    [1, 2, 3],
                    [4, 5, 6]
                }).AsEnumerable()
            ];
            yield return
            [
                new[]
                {
                    new decimal[5],
                    new decimal[5],
                }
            ];
        }
    }

    private static readonly int[] SourceArray = [1, 2, 3];
    private static readonly double[] SourceArray0 = [1.0, 2.0, 3.0];

    [Test]
    public async Task CanCreateArrayReaderFromIEnumerableOfIEnumerables()
    {
        IEnumerable<IEnumerable> data = new List<IEnumerable>
        {
            SourceArray.AsEnumerable(),
            SourceArray0.AsEnumerable(),
        };
        var reader = InsertDataReaderFactory.CreateReader(data);

        await Assert.That(reader).IsAssignableTo<ArrayReader>();
    }

    [Test]
    public async Task CanCreateSimpleReaderFromIEnumerableOfString()
    {
        IEnumerable<string> data = new[]
        {
            "String 1",
            "String 2",
        };
        var reader = InsertDataReaderFactory.CreateReader(data);

        await Assert.That(reader).IsAssignableTo<SimpleTypeReader>();
    }

    [Test]
    public async Task CanCreateDataTableReader()
    {
        var dt = new DataTable();
        var reader = InsertDataReaderFactory.CreateReader(dt);

        await Assert.That(reader).IsAssignableTo<XLibur.Excel.InsertData.DataTableReader>();
    }

    [Test]
    public async Task CanCreateDataRecordReader()
    {
        var dataRecords = Array.Empty<IDataRecord>();
        var reader = InsertDataReaderFactory.CreateReader(dataRecords);
        await Assert.That(reader).IsAssignableTo<DataRecordReader>();
    }

    [Test]
    public async Task CanCreateObjectReader()
    {
        var entities = Array.Empty<TestEntity>();
        var reader = InsertDataReaderFactory.CreateReader(entities);
        await Assert.That(reader).IsAssignableTo<ObjectReader>();
    }

    [Test]
    public async Task CanCreateObjectReaderForStruct()
    {
        var entities = Array.Empty<TestStruct>();
        var reader = InsertDataReaderFactory.CreateReader(entities);
        await Assert.That(reader).IsAssignableTo<ObjectReader>();
    }

    [Test]
    public async Task CanCreateUntypedObjectReader()
    {
        var entities = new ArrayList(new object[]
        {
            new TestEntity(),
            "123",
        });
        var reader = InsertDataReaderFactory.CreateReader(entities);
        await Assert.That(reader).IsAssignableTo<UntypedObjectReader>();
    }

    private class TestEntity;

    private struct TestStruct;
}
